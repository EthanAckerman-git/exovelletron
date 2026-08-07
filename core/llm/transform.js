/**
 * Row-wise transformation over a whole column.
 *
 * Some jobs — standardising 500 addresses to USPS format, splitting names, normalising
 * dates — cannot be done from the context sample and cannot be expressed as a formula.
 * They need the model to actually see every row.
 *
 * Feeding 500 rows into one prompt would blow the context and degrade badly toward the
 * end, so rows are processed in small batches. A batch that returns the wrong number of
 * results is split and retried rather than being silently misaligned — writing row 200's
 * answer into row 199 would corrupt data in a way that is very hard to notice.
 *
 * The chunking and parsing here are pure so they can be tested without a model.
 */

/** Rows per request. Small enough to stay reliable, large enough to amortise the prompt. */
export const DEFAULT_BATCH_SIZE = 10;
/** Refuse absurd jobs outright rather than running for an hour. */
export const MAX_TRANSFORM_ROWS = 5000;

/** Split values into index-tagged batches. */
export function batchRows(values, size = DEFAULT_BATCH_SIZE) {
  const batches = [];
  for (let i = 0; i < values.length; i += size) {
    batches.push({ start: i, rows: values.slice(i, i + size) });
  }
  return batches;
}

/**
 * JSON schema the model's output is constrained to.
 *
 * `fieldCount` of 1 yields one string per row (a rewrite); anything higher yields a
 * fixed-width row of strings (a split). Pinning both the row count and the field count
 * in the grammar is what stops a batch coming back misaligned.
 */
export const batchSchema = (count, fieldCount = 1) => ({
  type: "object",
  properties: {
    results: {
      type: "array",
      minItems: count,
      maxItems: count,
      items: fieldCount === 1
        ? { type: "string" }
        : { type: "array", minItems: fieldCount, maxItems: fieldCount, items: { type: "string" } },
    },
  },
  required: ["results"],
});

export function buildBatchPrompt(instruction, rows, startRow, fields = null) {
  const numbered = rows.map((value, i) => `${startRow + i}. ${String(value ?? "")}`).join("\n");

  const shape = fields?.length
    ? `- For each input line return exactly ${fields.length} values, in this order: ${fields.join(", ")}.
- Leave a value as an empty string when the input does not contain it. Never guess or invent one.`
    : `- Return one transformed value per input line.
- If a line is empty or cannot be transformed, return it unchanged.`;

  return `Apply this to each numbered input line:

${instruction}

Rules:
- Return exactly ${rows.length} results, in the same order as the inputs.
${shape}
- Treat each line independently. Do not merge, reorder, drop, or add lines.
- Return only the values, with no numbering and no commentary.

Inputs:
${numbered}`;
}

/**
 * Pull the results array out of a model response.
 * Accepts the constrained object form and tolerates a bare array or fenced JSON.
 * @returns {string[]|null} null when nothing usable was found
 */
export function parseBatchResult(raw, fieldCount = 1) {
  const cell = (v) => (v == null ? "" : String(v));
  const row = (v) => (fieldCount === 1
    ? cell(v)
    : Array.isArray(v) ? v.slice(0, fieldCount).map(cell) : null);

  if (raw == null) return null;
  if (Array.isArray(raw)) {
    const mapped = raw.map(row);
    return mapped.some((v) => v === null) ? null : mapped;
  }
  if (typeof raw === "object" && Array.isArray(raw.results)) {
    const mapped = raw.results.map(row);
    return mapped.some((v) => v === null) ? null : mapped;
  }
  if (typeof raw !== "string") return null;

  const text = raw.trim();
  const fenced = /```(?:json)?\s*([\s\S]*?)```/.exec(text);
  const candidate = fenced ? fenced[1].trim() : text;
  try {
    return parseBatchResult(JSON.parse(candidate), fieldCount);
  } catch {
    return null;
  }
}

/**
 * Fraction of the source text that must survive a split before it is called lossy.
 *
 * A split is meant to redistribute content, not discard it. When the model picks too few
 * output columns — five for data that needs six — it quietly drops whatever did not fit,
 * and the sheet looks perfectly plausible afterwards. Comparing the characters that
 * survive catches that; punctuation and separators are ignored because a split
 * legitimately removes quotes and commas.
 */
export const LOSS_THRESHOLD = 0.8;

const meaningful = (text) => String(text ?? "").replace(/[^a-z0-9]/gi, "").toUpperCase();

/**
 * Rows where a split dropped a meaningful amount of the original text.
 *
 * Splits only. A rewrite is *supposed* to shed characters — "STREET" becoming "ST" is
 * the whole point — so measuring it this way would flag every successful abbreviation.
 * Rows that are not field arrays are skipped for that reason.
 *
 * @param {string[]} values source values
 * @param {Array<string[]>} results per-row field arrays
 * @returns {number[]} row indexes that lost content
 */
export function findLossyRows(values, results) {
  const lossy = [];
  for (let i = 0; i < values.length; i++) {
    if (!Array.isArray(results[i])) continue;
    const source = meaningful(values[i]);
    if (source.length < 12) continue;
    const kept = meaningful(results[i].join(""));
    if (kept.length / source.length < LOSS_THRESHOLD) lossy.push(i);
  }
  return lossy;
}

/**
 * Run a transformation over every value, batching and splitting on mismatch.
 *
 * @param {object} opts
 * @param {string[]} opts.values          source cell values, in row order
 * @param {string} opts.instruction       what to do to each value
 * @param {(prompt:string, schema:object) => Promise<any>} opts.complete  constrained completion
 * @param {(p:{done:number,total:number,results:Array<string|null>}) => void} [opts.onProgress]
 * @param {AbortSignal} [opts.signal]
 * @param {number} [opts.batchSize]
 * @returns {Promise<{results:string[], failed:number[]}>}
 */
export async function transformValues({
  values,
  instruction,
  complete,
  onProgress,
  signal,
  batchSize = DEFAULT_BATCH_SIZE,
  fields = null,
}) {
  const fieldCount = fields?.length ?? 1;
  if (values.length > MAX_TRANSFORM_ROWS) {
    throw new Error(`That is ${values.length} rows; the limit is ${MAX_TRANSFORM_ROWS}.`);
  }

  const results = new Array(values.length).fill(null);
  const failed = [];
  let done = 0;

  /** Process one span, halving it when the model returns the wrong number of results. */
  async function run(start, rows) {
    if (signal?.aborted) throw Object.assign(new Error("Transform cancelled"), { code: "ABORTED" });
    if (!rows.length) return;

    // Blank inputs never need the model.
    if (rows.every((v) => String(v ?? "").trim() === "")) {
      const blank = fieldCount === 1 ? "" : new Array(fieldCount).fill("");
      for (let i = 0; i < rows.length; i++) results[start + i] = blank;
      done += rows.length;
      onProgress?.({ done, total: values.length, results });
      return;
    }

    let parsed = null;
    try {
      parsed = parseBatchResult(
        await complete(buildBatchPrompt(instruction, rows, start + 1, fields), batchSchema(rows.length, fieldCount)),
        fieldCount,
      );
    } catch (err) {
      if (err.code === "ABORTED") throw err;
      parsed = null;
    }

    if (parsed && parsed.length === rows.length) {
      for (let i = 0; i < rows.length; i++) results[start + i] = parsed[i];
      done += rows.length;
      onProgress?.({ done, total: values.length, results });
      return;
    }

    // Misaligned or unparseable. One row cannot be split further, so record the failure
    // and keep the original value rather than writing a guess into the wrong cell.
    if (rows.length === 1) {
      // Keep the original in the first field so nothing is lost, and leave the rest blank.
      results[start] = fieldCount === 1
        ? String(rows[0] ?? "")
        : [String(rows[0] ?? ""), ...new Array(fieldCount - 1).fill("")];
      failed.push(start);
      done += 1;
      onProgress?.({ done, total: values.length, results });
      return;
    }

    const mid = Math.ceil(rows.length / 2);
    await run(start, rows.slice(0, mid));
    await run(start + mid, rows.slice(mid));
  }

  for (const { start, rows } of batchRows(values, batchSize)) {
    await run(start, rows);
  }

  const fallback = (i) => (fieldCount === 1
    ? String(values[i] ?? "")
    : [String(values[i] ?? ""), ...new Array(fieldCount - 1).fill("")]);

  return { results: results.map((v, i) => (v == null ? fallback(i) : v)), failed };
}
