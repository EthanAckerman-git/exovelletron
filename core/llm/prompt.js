/**
 * System prompt and sheet-context rendering.
 *
 * The model gets a compact, explicitly-truncated view of the worksheet. Being honest
 * about what was elided matters: a model that thinks it saw all 40,000 rows will
 * confidently compute a total from the 30 it actually got. Telling it the data is a
 * sample pushes it toward writing a formula instead, which is the right answer anyway.
 */

export const SYSTEM_PROMPT = `You are the assistant inside "Excel AI Local", running entirely on the user's own Mac. Nothing leaves their computer.

You help with one open Excel workbook. You are given a partial view of it: the user's current selection, column headers, and a sample of rows. Row counts tell you the true size — the sample is never the whole sheet.

How to work:
- Answer in plain language, briefly. Skip preamble and restating the question.
- To change the workbook, call one of the provided tools. Each call becomes a preview the user approves or rejects, so propose the change rather than describing it.
- Prefer a formula over computed literal values whenever the answer derives from sheet data. A formula stays correct when the data changes, and you only ever see a sample of the rows.
- Use exact A1 addresses. When the user says "this column" or "the selection", resolve it against the selection you were given.
- Only reference columns and sheets that exist in the context.
- If a request is ambiguous about where output should go, ask one short question instead of guessing.
- If asked something unrelated to the workbook, just answer it normally.

Formula rules:
- Write the formula for the FIRST cell of the target range; Excel adjusts relative references down the range automatically.
- Size the range to the rows that actually hold data. The used range tells you the last row — if data ends at row 120, fill E2:E120, never E2:E1000. Padding to a round number puts formulas on thousands of empty rows.
- Use standard function names in English, comma-separated arguments.
- Guard divisions that could hit empty cells, e.g. =IFERROR(A2/B2,"").`;

const MAX_CELL_CHARS = 120;

/** Trim one cell for display, marking truncation so the model knows it is partial. */
function renderCell(value) {
  if (value === null || value === undefined || value === "") return "";
  let s = typeof value === "object" ? JSON.stringify(value) : String(value);
  s = s.replace(/[\r\n\t]+/g, " ").trim();
  return s.length > MAX_CELL_CHARS ? `${s.slice(0, MAX_CELL_CHARS - 1)}…` : s;
}

/** Cheap token estimate. Good enough for budgeting; no tokenizer round-trip needed. */
export const estimateTokens = (text) => Math.ceil(text.length / 3.6);

function renderGrid(startRow, startColumnLetters, rows) {
  if (!rows?.length) return "";
  const header = ["row", ...startColumnLetters].join(" | ");
  const divider = ["---", ...startColumnLetters.map(() => "---")].join(" | ");
  const body = rows.map((cells, i) => [String(startRow + i), ...cells.map(renderCell)].join(" | "));
  return [header, divider, ...body].join("\n");
}

/**
 * Render the worksheet context into the text the model sees.
 *
 * @param {object} ctx  produced by taskpane/sheet/context.js
 * @param {number} budgetTokens
 */
export function renderSheetContext(ctx, budgetTokens = 6000) {
  if (!ctx || !ctx.sheetName) return "No workbook data is available.";

  const lines = [];
  lines.push(`Workbook: ${ctx.workbookName || "Untitled"}`);
  lines.push(`Active sheet: ${ctx.sheetName}`);
  if (ctx.sheetNames?.length > 1) lines.push(`All sheets: ${ctx.sheetNames.join(", ")}`);

  if (ctx.usedRange) {
    lines.push(
      `Used range: ${ctx.usedRange.address} (${ctx.usedRange.rowCount} rows x ${ctx.usedRange.columnCount} columns)`,
    );
  }

  if (ctx.selection) {
    const s = ctx.selection;
    lines.push(
      `Current selection: ${s.address} (${s.rowCount} rows x ${s.columnCount} columns${s.isSingleCell ? ", a single cell" : ""})`,
    );
  }

  if (ctx.headers?.length) {
    const named = ctx.headers.map((h, i) => `${ctx.columnLetters[i]}=${renderCell(h) || "(blank)"}`);
    lines.push(`Columns: ${named.join(", ")}`);
  }

  const head = lines.join("\n");
  let used = estimateTokens(head);

  const sections = [];

  if (ctx.selectionSample?.rows?.length) {
    const grid = renderGrid(ctx.selectionSample.startRow, ctx.selectionSample.columnLetters, ctx.selectionSample.rows);
    const text = `\nSelected cells (${ctx.selectionSample.rows.length} of ${ctx.selection?.rowCount ?? ctx.selectionSample.rows.length} rows shown):\n${grid}`;
    if (used + estimateTokens(text) <= budgetTokens) {
      sections.push(text);
      used += estimateTokens(text);
    }
  }

  if (ctx.sample?.rows?.length) {
    const total = ctx.usedRange?.rowCount ?? ctx.sample.rows.length;
    const shown = ctx.sample.rows.length;
    // Drop rows from the sample until the grid fits the remaining budget.
    let rows = ctx.sample.rows;
    let text = "";
    while (rows.length) {
      const grid = renderGrid(ctx.sample.startRow, ctx.sample.columnLetters, rows);
      const note = shown < total ? ` — a sample of ${total} total rows, NOT the full sheet` : "";
      text = `\nSheet data (rows ${ctx.sample.startRow}–${ctx.sample.startRow + rows.length - 1}${note}):\n${grid}`;
      if (used + estimateTokens(text) <= budgetTokens) break;
      rows = rows.slice(0, Math.max(1, Math.floor(rows.length * 0.7)));
      if (rows.length === 1) break;
    }
    if (text) {
      sections.push(text);
      if (rows.length < shown) {
        sections.push(`(${shown - rows.length} more sampled rows omitted to fit the context window.)`);
      }
    }
  }

  if (ctx.formulas?.length) {
    const text = `\nExisting formulas: ${ctx.formulas.slice(0, 12).map((f) => `${f.address}: ${f.formula}`).join("; ")}`;
    if (used + estimateTokens(text) <= budgetTokens) sections.push(text);
  }

  return [head, ...sections].join("\n");
}

/** The user-turn text: worksheet context followed by the actual question. */
export function buildUserTurn(question, sheetContext, budgetTokens) {
  if (!sheetContext) return question;
  return `<worksheet>\n${renderSheetContext(sheetContext, budgetTokens)}\n</worksheet>\n\n${question}`;
}
