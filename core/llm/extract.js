/**
 * Whole-range record extraction.
 *
 * The row-by-row transform is strictly one-row-in, one-row-out, and its prompt orders
 * the model to treat every line independently. That is exactly right for "standardise
 * each address" and exactly wrong for the other kind of mess: a record whose fields are
 * stacked down five consecutive rows, or a single cell holding a dozen CSV records.
 * Assembling those requires seeing the neighbouring rows — the one thing the transform
 * forbids.
 *
 * So extraction reads the range as one document. The model gets a chunk of lines in
 * their original order and returns however many records it finds — the output row count
 * is deliberately untied from the input row count. Chunks are cut at blank lines
 * whenever possible, because vertically-stacked records are almost always separated by
 * one, and severing a record mid-chunk would split it into two half-records.
 *
 * Everything here is pure so it can be tested without a model.
 */

/** Refuse absurd jobs outright rather than running for an hour. */
export const MAX_EXTRACT_ROWS = 2000;

/**
 * Characters of input per model call. Extraction output repeats every field name for
 * every record, so the reply can run ~2x the input; this keeps prompt plus reply well
 * inside the extraction context.
 */
export const CHUNK_CHAR_BUDGET = 3500;

/** Grammar bound on records per chunk; a chunk of this size cannot legitimately hold more. */
export const MAX_RECORDS_PER_CHUNK = 200;

/**
 * Split lines into chunks under the character budget, preferring to cut at a blank
 * line so a multi-row record is never severed. A single line larger than the budget
 * becomes its own chunk — a line cannot be split.
 *
 * @param {string[]} lines
 * @param {number} [budget]
 * @returns {Array<{start:number, lines:string[]}>}
 */
export function chunkLines(lines, budget = CHUNK_CHAR_BUDGET) {
  const chunks = [];
  let start = 0;
  let current = [];
  let size = 0;
  let lastBlank = -1; // index within `current` of the most recent blank line

  const emit = (count) => {
    if (count <= 0) return;
    chunks.push({ start, lines: current.slice(0, count) });
    current = current.slice(count);
    start += count;
    size = current.reduce((n, l) => n + l.length + 1, 0);
    lastBlank = current.findLastIndex((l) => l.trim() === "");
  };

  for (const raw of lines) {
    const line = String(raw ?? "");
    current.push(line);
    size += line.length + 1;
    if (line.trim() === "") lastBlank = current.length - 1;

    if (size > budget && current.length > 1) {
      // Cut at the last blank line when there is one; otherwise keep everything up to
      // (but not including) the line that overflowed.
      emit(lastBlank > 0 ? lastBlank + 1 : current.length - 1);
    }
  }
  emit(current.length);
  return chunks;
}

/**
 * JSON schema the model's output is constrained to: an array of records, each an
 * object with exactly the requested fields. Objects rather than positional arrays,
 * because a small model maps values onto named keys far more reliably than onto
 * "the third element".
 */
export const extractSchema = (fields, maxRecords = MAX_RECORDS_PER_CHUNK) => ({
  type: "object",
  properties: {
    records: {
      type: "array",
      maxItems: maxRecords,
      items: {
        type: "object",
        properties: Object.fromEntries(fields.map((f) => [f, { type: "string" }])),
        required: [...fields],
      },
    },
  },
  required: ["records"],
});

export function buildExtractPrompt(instruction, fields, lines) {
  return `The numbered lines below are cells from a spreadsheet, in their original order. They hold records, but the layout is messy: one record's fields may be spread over several consecutive lines, one line may pack several complete records together, formats vary, and some lines are blank or irrelevant.

Task:
${instruction}

Rules:
- Read the lines in order and find every distinct record.
- Use neighbouring lines: a lone value usually belongs to the record being assembled around it. A record can start at the end of one line and finish at the start of the next.
- When one line packs several records together, walk its pieces in order, assigning each piece to the current record's next empty field and starting a new record when the pattern repeats.
- Fill every field the record's text contains. When one piece of text holds several of the requested fields — such as a person's name followed by a street address — split it across those fields rather than leaving them joined in one.
- Return one object per record with exactly these fields: ${fields.join(", ")}.
- A field the input does not contain is "" — never guess or invent a value.
- Do not skip records, and do not return the same record twice.
- Return nothing but the records. No commentary.

Lines:
${lines.map((l, i) => `${i + 1}. ${String(l ?? "")}`).join("\n")}`;
}

/**
 * Pull the record list out of a model response, as one array of field values per
 * record, in field order. Accepts the constrained object form and tolerates a bare
 * array or fenced JSON. Records with every field blank are dropped — they are the
 * model's way of saying "nothing here".
 *
 * @returns {string[][]|null} null when nothing usable was found
 */
export function parseExtractResult(raw, fields) {
  const cell = (v) => (v == null ? "" : String(v).trim());

  const record = (item) => {
    if (Array.isArray(item)) return item.slice(0, fields.length).map(cell);
    if (item && typeof item === "object") return fields.map((f) => cell(item[f]));
    return null;
  };

  const fromList = (list) => {
    const mapped = list.map(record);
    if (mapped.some((r) => r === null)) return null;
    return mapped
      .map((r) => (r.length < fields.length ? [...r, ...new Array(fields.length - r.length).fill("")] : r))
      .filter((r) => r.some((v) => v !== ""));
  };

  if (raw == null) return null;
  if (Array.isArray(raw)) return fromList(raw);
  if (typeof raw === "object" && Array.isArray(raw.records)) return fromList(raw.records);
  if (typeof raw !== "string") return null;

  const text = raw.trim();
  const fenced = /```(?:json)?\s*([\s\S]*?)```/.exec(text);
  try {
    return parseExtractResult(JSON.parse(fenced ? fenced[1].trim() : text), fields);
  } catch {
    return null;
  }
}

/**
 * Extract records from every line, chunking and halving failed chunks.
 *
 * @param {object} opts
 * @param {string[]} opts.values        source rows as text, in order
 * @param {string} opts.instruction     what one record is and how to fill each field
 * @param {string[]} opts.fields        output field names, in column order
 * @param {(prompt:string, schema:object) => Promise<any>} opts.complete constrained completion
 * @param {(p:{done:number,total:number,records:number,sample:string[][]}) => void} [opts.onProgress]
 * @param {AbortSignal} [opts.signal]
 * @param {number} [opts.chunkBudget]
 * @returns {Promise<{records:string[][], failed:number[]}>} failed holds the indexes of
 *   lines whose chunk could not be parsed even alone; their content is not in `records`.
 */
export async function extractRecords({
  values,
  instruction,
  fields,
  complete,
  onProgress,
  signal,
  chunkBudget = CHUNK_CHAR_BUDGET,
}) {
  if (values.length > MAX_EXTRACT_ROWS) {
    throw new Error(`That is ${values.length} rows; the limit is ${MAX_EXTRACT_ROWS}.`);
  }

  const records = [];
  const failed = [];
  let done = 0;

  const report = () => onProgress?.({
    done,
    total: values.length,
    records: records.length,
    sample: records.slice(0, 5),
  });

  /** Process one span of lines, halving it when the model returns nothing usable. */
  async function run(start, lines) {
    if (signal?.aborted) throw Object.assign(new Error("Extraction cancelled"), { code: "ABORTED" });
    if (!lines.length) return;

    // Blank lines never need the model.
    if (lines.every((l) => String(l ?? "").trim() === "")) {
      done += lines.length;
      report();
      return;
    }

    let parsed = null;
    try {
      parsed = parseExtractResult(
        await complete(buildExtractPrompt(instruction, fields, lines), extractSchema(fields)),
        fields,
      );
    } catch (err) {
      if (err.code === "ABORTED") throw err;
      parsed = null;
    }

    if (parsed) {
      records.push(...parsed);
      done += lines.length;
      report();
      return;
    }

    if (lines.length === 1) {
      failed.push(start);
      done += 1;
      report();
      return;
    }

    const mid = Math.ceil(lines.length / 2);
    await run(start, lines.slice(0, mid));
    await run(start + mid, lines.slice(mid));
  }

  for (const chunk of chunkLines(values.map((v) => String(v ?? "")), chunkBudget)) {
    await run(chunk.start, chunk.lines);
  }

  return { records, failed };
}
