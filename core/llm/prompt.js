/**
 * System prompt and sheet-context rendering.
 *
 * The model gets a compact, explicitly-truncated view of the worksheet. Being honest
 * about what was elided matters: a model that thinks it saw all 40,000 rows will
 * confidently compute a total from the 30 it actually got. Telling it the data is a
 * sample pushes it toward writing a formula instead, which is the right answer anyway.
 */

export const SYSTEM_PROMPT = `You are the assistant inside "Exovelletron", running entirely on the user's own Mac. Nothing leaves their computer.

You help with one open Excel workbook. You are given a partial view of it: the user's current selection, column headers, and a sample of rows. Row counts tell you the true size — the sample is never the whole sheet.

How to work:
- Answer in plain language, briefly. Skip preamble and restating the question.
- To change the workbook, call one of the provided tools. Each call becomes a preview the user approves or rejects, so propose the change rather than describing it.
- Prefer a formula over computed literal values whenever the answer derives from sheet data. A formula stays correct when the data changes, and you only ever see a sample of the rows.
- Before splitting a column, count the distinct pieces in the sample values and give each one its own output column. Leaving a piece uncovered throws it away.
- When asked to tidy or "make it look nice", follow the change with format_cells: bold the header row, give it a light fill, and apply a sensible number format (currency, percent, date) to columns that deserve one. Keep it restrained — one accent colour, not a rainbow.
- When one column holds several fields mashed together — a raw CSV line, "NAME 123 MAIN ST CITY ST 12345", "Last, First" — use split_column to break it into proper columns. Do this before trying to tidy the values. split_column requires that each row is exactly one complete record.
- When the data is NOT one record per row — a record's fields stacked down several consecutive rows (name on one row, street on the next, city below that), several complete records packed into one giant cell, or a mix of layouts — use extract_table instead. It reads the whole range as one document, uses neighbouring rows to assemble each record, and writes one clean row per record. split_column and transform_column are strictly one-row-in one-row-out and can never merge rows, so they are the wrong tool for stacked or packed records. Check the sample: if fields of one record appear on separate rows, that is extract_table.
- When every row needs its own judgement that no formula can express — standardising addresses, splitting names, normalising messy dates, categorising free text — use transform_column. It runs over the entire column, including the rows you cannot see, so never write out values row by row yourself and never apologise that you can only see a sample.
- Use exact A1 addresses. When the user says "this column" or "the selection", resolve it against the selection you were given.
- The sheet itself is context: read the headers, the nearby columns, and the shape of the data to work out what the user means before asking. If the sheet already has labelled but empty columns that clearly exist for the output — headers with nothing beneath them — write the results into those columns, in their order, rather than inventing a new location.
- Only reference columns and sheets that exist in the context.
- If a request is ambiguous about where output should go, ask one short question instead of guessing.
- If asked something unrelated to the workbook, just answer it normally.

Formula rules:
- Write the formula for the FIRST cell of the target range; Excel adjusts relative references down the range automatically.
- Size the range to the rows that actually hold data. The used range tells you the last row — if data ends at row 120, fill E2:E120, never E2:E1000. Padding to a round number puts formulas on thousands of empty rows.
- Use standard function names in English, comma-separated arguments.
- Guard divisions that could hit empty cells, e.g. =IFERROR(A2/B2,"").`;

/**
 * Cells are shown almost in full. Messy imports routinely put an entire CSV line into
 * one cell, and truncating those hides exactly the content that needs fixing. The token
 * budget still trims whole rows if the sheet is large, which degrades more predictably
 * than silently clipping every cell.
 */
const MAX_CELL_CHARS = 600;

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

  // The last rows, when the sheet runs past the head sample. Seeing both ends makes the
  // sheet's real size concrete instead of something to infer from a row count.
  if (ctx.tail?.rows?.length) {
    const grid = renderGrid(ctx.tail.startRow, ctx.tail.columnLetters, ctx.tail.rows);
    const text = `\nLast rows of the sheet (rows ${ctx.tail.startRow}–${ctx.tail.startRow + ctx.tail.rows.length - 1}):\n${grid}`;
    if (used + estimateTokens(text) <= budgetTokens) {
      sections.push(text);
      used += estimateTokens(text);
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
