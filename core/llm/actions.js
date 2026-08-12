/**
 * The vocabulary of changes the model may propose to a worksheet.
 *
 * The model never edits anything directly. It calls one of these tools, we validate
 * and queue the result, and the task pane renders it as a preview the user has to
 * approve. That keeps a hallucinated range from quietly overwriting real data.
 *
 * Everything here is pure so it can be unit-tested without a model or Excel.
 */

/** Excel's hard grid limits. */
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;
/** A single approved action should stay reviewable; refuse absurd writes. */
export const MAX_CELLS_PER_ACTION = 50_000;

const A1_CELL = /^\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})$/i;

export function columnToIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function indexToColumn(index) {
  let n = index;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/**
 * Parse an A1 range ("B2", "A1:D20", "Sheet1!A1:B3").
 * @returns {{sheet:string|null,start:{col:number,row:number},end:{col:number,row:number},address:string,rows:number,cols:number}}
 */
export function parseRange(input) {
  if (typeof input !== "string" || !input.trim()) throw new Error("Range is required");
  let ref = input.trim();
  let sheet = null;

  const bang = ref.lastIndexOf("!");
  if (bang !== -1) {
    sheet = ref.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'");
    ref = ref.slice(bang + 1);
    if (!sheet) throw new Error(`Invalid sheet name in "${input}"`);
  }

  const parts = ref.split(":");
  // "A1:" and "A1:B2:C3" are malformed; a bare "A1" is a one-cell range.
  if (parts.length > 2) throw new Error(`"${input}" is not a valid range`);
  const [rawStart, rawEnd] = parts;
  const s = A1_CELL.exec(rawStart ?? "");
  if (!s) throw new Error(`"${input}" is not a valid cell reference`);
  const e = parts.length === 2 ? A1_CELL.exec(rawEnd ?? "") : s;
  if (!e) throw new Error(`"${input}" is not a valid range`);

  const start = { col: columnToIndex(s[1]), row: Number(s[2]) };
  const end = { col: columnToIndex(e[1]), row: Number(e[2]) };
  const norm = {
    start: { col: Math.min(start.col, end.col), row: Math.min(start.row, end.row) },
    end: { col: Math.max(start.col, end.col), row: Math.max(start.row, end.row) },
  };

  if (norm.end.col > MAX_COLS || norm.end.row > MAX_ROWS) throw new Error(`"${input}" is outside the worksheet`);

  const rows = norm.end.row - norm.start.row + 1;
  const cols = norm.end.col - norm.start.col + 1;
  const address = `${indexToColumn(norm.start.col)}${norm.start.row}:${indexToColumn(norm.end.col)}${norm.end.row}`;
  return { sheet, ...norm, address, rows, cols };
}

/** Rectangular check + cell budget, shared by value-writing actions. */
function assertGridFits(values, rows, cols) {
  if (!Array.isArray(values) || !values.length) throw new Error("values must be a non-empty array of rows");
  if (!values.every((r) => Array.isArray(r))) throw new Error("values must be an array of arrays");
  const width = values[0].length;
  if (!width) throw new Error("values rows must not be empty");
  if (!values.every((r) => r.length === width)) throw new Error("every row in values must have the same length");
  if (values.length * width > MAX_CELLS_PER_ACTION) {
    throw new Error(`That would write ${values.length * width} cells; the limit is ${MAX_CELLS_PER_ACTION}`);
  }
  if (values.length !== rows || width !== cols) {
    throw new Error(`values is ${values.length}x${width} but ${rows}x${cols} was requested`);
  }
}

const scalar = (v) => (v === null || v === undefined ? "" : typeof v === "object" ? JSON.stringify(v) : v);

/**
 * Undo one formula habit the model has that Excel's API hard-rejects: wrapping plain
 * cell references in square brackets, `=[B2]*[C2]`. To Excel a leading bracket means
 * an external-workbook reference, so applying such a formula fails with "the argument
 * is invalid". Only a bracket holding nothing but one A1-style reference is unwrapped —
 * structured table references like `Table1[@Units]` or `[[#This Row],[Price]]` carry
 * no bare row number and pass through untouched.
 */
export const sanitizeFormula = (formula) =>
  String(formula).replace(/\[(\$?[A-Za-z]{1,3}\$?\d{1,7})\]/g, "$1");

/**
 * Tool definitions handed to the model, plus a validator that turns raw arguments into
 * a normalized action. Keeping the schema and the validator adjacent means they cannot
 * drift apart.
 */
export const ACTION_SPECS = {
  write_values: {
    description:
      "Write literal values into a cell range. Use for filling in data, labels, or computed results. " +
      "The values grid must exactly match the size of the range.",
    params: {
      type: "object",
      properties: {
        address: { type: "string", description: 'Target range in A1 notation, e.g. "D2:D25".' },
        values: {
          type: "array",
          description: "Rows of cell values, outer array is rows, inner array is columns.",
          items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } },
        },
        sheet: { type: "string", description: "Worksheet name. Omit for the active sheet." },
      },
      required: ["address", "values"],
    },
    validate(args) {
      const range = parseRange(args.sheet ? `${args.sheet}!${args.address}` : args.address);
      const values = args.values.map((row) => row.map(scalar));
      assertGridFits(values, range.rows, range.cols);
      return {
        type: "write_values",
        sheet: range.sheet,
        address: range.address,
        values,
        summary: `Write ${range.rows * range.cols} value${range.rows * range.cols === 1 ? "" : "s"} to ${range.sheet ? `${range.sheet}!` : ""}${range.address}`,
      };
    },
  },

  write_formula: {
    description:
      "Fill an Excel formula down a range, exactly like dragging the fill handle: write it for the " +
      'FIRST cell and relative references adjust per row automatically. Always start with "=". ' +
      "Size the range to the rows that actually contain data — do not pad it out to a round number " +
      "like 1000, which would put formulas on thousands of empty rows.",
    params: {
      type: "object",
      properties: {
        address: { type: "string", description: 'Target range in A1 notation, e.g. "E2:E100".' },
        formula: { type: "string", description: 'The formula, starting with "=", written for the first cell of the range.' },
        sheet: { type: "string", description: "Worksheet name. Omit for the active sheet." },
      },
      required: ["address", "formula"],
    },
    validate(args, context) {
      const range = parseRange(args.sheet ? `${args.sheet}!${args.address}` : args.address);
      const formula = sanitizeFormula(String(args.formula ?? "").trim());
      if (!formula.startsWith("=")) throw new Error('formula must start with "="');
      if (formula.length > 8192) throw new Error("formula is too long");

      // Models habitually pad a fill-down to a round number — "E2:E1000" for six rows of
      // data — which buries the sheet in formulas over empty cells. Clamp to the rows
      // that actually hold data; the preview then shows the honest, smaller range.
      let { end, address } = range;
      let clamped = false;
      const lastDataRow = context?.lastRow;
      if (Number.isInteger(lastDataRow) && end.row > lastDataRow && range.start.row <= lastDataRow) {
        end = { ...end, row: lastDataRow };
        address = `${indexToColumn(range.start.col)}${range.start.row}:${indexToColumn(end.col)}${end.row}`;
        clamped = true;
      }

      const rows = end.row - range.start.row + 1;
      const cells = rows * range.cols;
      if (cells > MAX_CELLS_PER_ACTION) {
        throw new Error(`That would fill ${cells} cells; the limit is ${MAX_CELLS_PER_ACTION}`);
      }

      return {
        type: "write_formula",
        sheet: range.sheet,
        address,
        formula,
        clamped,
        summary:
          `Fill ${cells} cell${cells === 1 ? "" : "s"} in ${range.sheet ? `${range.sheet}!` : ""}${address} with ${formula}` +
          (clamped ? " (trimmed to the rows that contain data)" : ""),
      };
    },
  },

  transform_column: {
    description:
      "Rewrite every value in a column, row by row, using the whole column — not just the rows shown " +
      "in the context. Use this when each row needs its own judgement that no formula can express: " +
      "standardising postal addresses, splitting or reordering names, normalising inconsistent dates, " +
      "categorising free text, cleaning up messy entries. Say precisely what to do to a single value in " +
      "`instruction`; it is applied to each row independently.",
    params: {
      type: "object",
      properties: {
        source: { type: "string", description: 'Column range to read, e.g. "B2:B500".' },
        target: {
          type: "string",
          description: 'Where to write the results, same height as source, e.g. "C2:C500". May equal source to rewrite in place.',
        },
        instruction: {
          type: "string",
          description:
            'What to do to one value, e.g. "Standardise this US street address to USPS format: uppercase, ' +
            'standard suffix abbreviations (STREET->ST, AVENUE->AVE), no punctuation."',
        },
        sheet: { type: "string" },
      },
      required: ["source", "target", "instruction"],
    },
    validate(args, context) {
      const source = parseRange(args.sheet ? `${args.sheet}!${args.source}` : args.source);
      const target = parseRange(args.sheet ? `${args.sheet}!${args.target}` : args.target);

      if (source.cols !== 1 || target.cols !== 1) {
        throw new Error("transform_column works on a single column at a time");
      }

      // Same padding habit as write_formula: keep the job inside the real data.
      let sourceEnd = source.end.row;
      const lastDataRow = context?.lastRow;
      let clamped = false;
      if (Number.isInteger(lastDataRow) && sourceEnd > lastDataRow && source.start.row <= lastDataRow) {
        sourceEnd = lastDataRow;
        clamped = true;
      }

      const rows = sourceEnd - source.start.row + 1;
      if (rows < 1) throw new Error("The source range is empty");
      if (rows > 5000) throw new Error(`That is ${rows} rows; the limit is 5000`);

      const instruction = String(args.instruction ?? "").trim();
      if (instruction.length < 8) throw new Error("instruction must describe the transformation");
      if (instruction.length > 2000) throw new Error("instruction is too long");

      const sourceAddress = `${indexToColumn(source.start.col)}${source.start.row}:${indexToColumn(source.start.col)}${sourceEnd}`;
      const targetAddress = `${indexToColumn(target.start.col)}${target.start.row}:${indexToColumn(target.start.col)}${target.start.row + rows - 1}`;

      return {
        type: "transform_column",
        sheet: source.sheet ?? target.sheet,
        source: sourceAddress,
        address: targetAddress,
        instruction,
        rows,
        inPlace: sourceAddress === targetAddress,
        clamped,
        summary:
          `Rewrite ${rows} row${rows === 1 ? "" : "s"} from ${sourceAddress} into ${targetAddress}` +
          (clamped ? " (trimmed to the rows that contain data)" : ""),
      };
    },
  },

  split_column: {
    description:
      "Split one messy column into several clean columns, row by row, over the whole column. " +
      "Use when a single cell holds several fields jammed together — a raw CSV line, " +
      '"NAME 123 MAIN ST CITY STATE 12345", "Last, First" — and the pieces belong in separate ' +
      "columns. Name the output columns in order. This reads every row, not just the sample. " +
      "IMPORTANT: your columns must account for EVERY part of the input. Look at the sample rows " +
      "and count the distinct pieces before choosing. If one CSV field holds both a person's name " +
      "and a street address, that is TWO columns, not one — anything you leave uncovered is " +
      "silently discarded.",
    params: {
      type: "object",
      properties: {
        source: { type: "string", description: 'Column range to read, e.g. "A2:A500".' },
        target: { type: "string", description: 'Top-left cell of the output block, e.g. "B2".' },
        columns: {
          type: "array",
          description: 'Output column headers, in order, e.g. ["Name","Street","City","State","ZIP"].',
          items: { type: "string" },
        },
        instruction: {
          type: "string",
          description: "How to pull the fields out of one value, including how to handle missing pieces.",
        },
        sheet: { type: "string" },
      },
      required: ["source", "target", "columns", "instruction"],
    },
    validate(args, context) {
      const source = parseRange(args.sheet ? `${args.sheet}!${args.source}` : args.source);
      const target = parseRange(args.sheet ? `${args.sheet}!${args.target}` : args.target);
      if (source.cols !== 1) throw new Error("split_column reads a single column");

      const columns = (Array.isArray(args.columns) ? args.columns : [])
        .map((c) => String(c ?? "").trim())
        .filter(Boolean);
      if (columns.length < 2) throw new Error("split_column needs at least two output columns");
      if (columns.length > 20) throw new Error("split_column supports at most 20 output columns");

      const instruction = String(args.instruction ?? "").trim();
      if (instruction.length < 8) throw new Error("instruction must describe how to split each value");
      if (instruction.length > 2000) throw new Error("instruction is too long");

      let sourceEnd = source.end.row;
      const lastDataRow = context?.lastRow;
      let clamped = false;
      if (Number.isInteger(lastDataRow) && sourceEnd > lastDataRow && source.start.row <= lastDataRow) {
        sourceEnd = lastDataRow;
        clamped = true;
      }

      const rows = sourceEnd - source.start.row + 1;
      if (rows < 1) throw new Error("The source range is empty");
      if (rows > 5000) throw new Error(`That is ${rows} rows; the limit is 5000`);
      if (rows * columns.length > MAX_CELLS_PER_ACTION) {
        throw new Error(`That would write ${rows * columns.length} cells; the limit is ${MAX_CELLS_PER_ACTION}`);
      }

      // The output block would overwrite the column it is reading from.
      const targetEndCol = target.start.col + columns.length - 1;
      if (target.start.col <= source.start.col && targetEndCol >= source.start.col) {
        throw new Error(
          `The output would overwrite the source column ${indexToColumn(source.start.col)}. Pick a target to its right.`,
        );
      }

      const sourceAddress = `${indexToColumn(source.start.col)}${source.start.row}:${indexToColumn(source.start.col)}${sourceEnd}`;
      const targetAddress = `${indexToColumn(target.start.col)}${target.start.row}:${indexToColumn(targetEndCol)}${target.start.row + rows - 1}`;
      const headerAddress = target.start.row > 1
        ? `${indexToColumn(target.start.col)}${target.start.row - 1}:${indexToColumn(targetEndCol)}${target.start.row - 1}`
        : null;

      return {
        type: "split_column",
        sheet: source.sheet ?? target.sheet,
        source: sourceAddress,
        address: targetAddress,
        headerAddress,
        columns,
        instruction,
        rows,
        clamped,
        summary:
          `Split ${rows} row${rows === 1 ? "" : "s"} of ${sourceAddress} into ${columns.length} columns ` +
          `(${columns.join(", ")}) at ${targetAddress}` + (clamped ? " (trimmed to the rows that contain data)" : ""),
      };
    },
  },

  extract_table: {
    description:
      "Rebuild a messy range as a clean table of records. Use when the data is NOT one record " +
      "per row: a record's fields spread down several consecutive rows, several complete records " +
      "packed into one cell, or a mix of layouts. Reads the ENTIRE range as one document — every " +
      "row, with its neighbours for context — finds each distinct record, and writes one row per " +
      "record under the named columns. The number of output rows follows the number of records " +
      "found, not the number of input rows. For data that IS already one record per row, use " +
      "split_column or transform_column instead.",
    params: {
      type: "object",
      properties: {
        source: {
          type: "string",
          description: 'Range to read, e.g. "A2:A40". Include every row the messy data occupies.',
        },
        target: {
          type: "string",
          description:
            'Top-left cell where the first record goes, e.g. "C2". Column headers are written ' +
            "in the row above it. Must not overlap the source columns.",
        },
        columns: {
          type: "array",
          description: 'Output column headers, in order, e.g. ["Name","Street","City","State","ZIP"].',
          items: { type: "string" },
        },
        instruction: {
          type: "string",
          description: "What one record is in this data, and how to fill each column from it.",
        },
        sheet: { type: "string" },
      },
      required: ["source", "target", "columns", "instruction"],
    },
    validate(args, context) {
      const source = parseRange(args.sheet ? `${args.sheet}!${args.source}` : args.source);
      const target = parseRange(args.sheet ? `${args.sheet}!${args.target}` : args.target);

      const columns = (Array.isArray(args.columns) ? args.columns : [])
        .map((c) => String(c ?? "").trim())
        .filter(Boolean);
      if (columns.length < 1) throw new Error("extract_table needs at least one output column");
      if (columns.length > 20) throw new Error("extract_table supports at most 20 output columns");

      const instruction = String(args.instruction ?? "").trim();
      if (instruction.length < 8) throw new Error("instruction must describe what one record is");
      if (instruction.length > 2000) throw new Error("instruction is too long");

      let sourceEnd = source.end.row;
      const lastDataRow = context?.lastRow;
      let clamped = false;
      if (Number.isInteger(lastDataRow) && sourceEnd > lastDataRow && source.start.row <= lastDataRow) {
        sourceEnd = lastDataRow;
        clamped = true;
      }

      const rows = sourceEnd - source.start.row + 1;
      if (rows < 1) throw new Error("The source range is empty");
      if (rows > 2000) throw new Error(`That is ${rows} rows; the limit is 2000`);

      // The record count is unknown until the extraction runs, so any column overlap
      // with the source could overwrite data still being read. Refuse it outright.
      const targetEndCol = target.start.col + columns.length - 1;
      if (target.start.col <= source.end.col && targetEndCol >= source.start.col) {
        throw new Error(
          `The table would overwrite the source columns ` +
          `${indexToColumn(source.start.col)}–${indexToColumn(source.end.col)}. Pick a target to their right.`,
        );
      }

      const sourceAddress =
        `${indexToColumn(source.start.col)}${source.start.row}:${indexToColumn(source.end.col)}${sourceEnd}`;
      const targetCell = `${indexToColumn(target.start.col)}${target.start.row}`;
      const headerAddress = target.start.row > 1
        ? `${indexToColumn(target.start.col)}${target.start.row - 1}:${indexToColumn(targetEndCol)}${target.start.row - 1}`
        : null;

      return {
        type: "extract_table",
        sheet: source.sheet ?? target.sheet,
        source: sourceAddress,
        // The final block size depends on how many records exist; until the run
        // finishes, the target anchor is the honest address to show and reveal.
        address: targetCell,
        targetCol: target.start.col,
        targetRow: target.start.row,
        headerAddress,
        columns,
        instruction,
        rows,
        clamped,
        summary:
          `Read ${sourceAddress} (${rows} row${rows === 1 ? "" : "s"}) and rebuild it as a table ` +
          `(${columns.join(", ")}) starting at ${targetCell}` +
          (clamped ? " (trimmed to the rows that contain data)" : ""),
      };
    },
  },

  format_cells: {
    description:
      "Change how a range looks: number format, bold, italic, or background colour. " +
      "Does not change the underlying values. Pass ONLY the properties you are actually " +
      "changing — omit the rest. Sending fill or bold when the user did not ask for them " +
      "will overwrite the cells' existing styling.",
    params: {
      type: "object",
      properties: {
        address: { type: "string", description: 'Range in A1 notation, e.g. "B2:B50".' },
        numberFormat: { type: "string", description: 'Excel number format, e.g. "$#,##0.00", "0.0%", "yyyy-mm-dd".' },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        fill: { type: "string", description: 'Background colour as a hex code, e.g. "#FFF3CD".' },
        sheet: { type: "string" },
      },
      required: ["address"],
    },
    validate(args) {
      const range = parseRange(args.sheet ? `${args.sheet}!${args.address}` : args.address);
      const out = { type: "format_cells", sheet: range.sheet, address: range.address };
      const bits = [];
      if (typeof args.numberFormat === "string" && args.numberFormat.trim()) {
        out.numberFormat = args.numberFormat.trim().slice(0, 128);
        bits.push(`format ${out.numberFormat}`);
      }
      if (typeof args.bold === "boolean") { out.bold = args.bold; bits.push(args.bold ? "bold" : "not bold"); }
      if (typeof args.italic === "boolean") { out.italic = args.italic; bits.push(args.italic ? "italic" : "not italic"); }
      if (typeof args.fill === "string") {
        const hex = args.fill.trim();
        if (!/^#?[0-9a-f]{6}$/i.test(hex)) throw new Error(`"${args.fill}" is not a 6-digit hex colour`);
        out.fill = hex.startsWith("#") ? hex.toUpperCase() : `#${hex.toUpperCase()}`;
        bits.push(`fill ${out.fill}`);
      }
      if (!bits.length) throw new Error("format_cells needs at least one formatting option");
      out.summary = `Format ${range.sheet ? `${range.sheet}!` : ""}${range.address}: ${bits.join(", ")}`;
      return out;
    },
  },

  insert_column: {
    description:
      "Insert a new empty column between existing data columns, optionally with a header in row 1. " +
      "Rarely needed: split_column, transform_column, and extract_table write their own headers at " +
      "their target, so NEVER insert columns for them — and never insert a column you are about to " +
      "write into anyway. Every insert shifts the address of every column after it.",
    params: {
      type: "object",
      properties: {
        before: { type: "string", description: 'Column letter to insert before, e.g. "D".' },
        header: { type: "string", description: "Optional header text for row 1 of the new column." },
        sheet: { type: "string" },
      },
      required: ["before"],
    },
    validate(args) {
      const letters = String(args.before ?? "").trim().toUpperCase();
      if (!/^[A-Z]{1,3}$/.test(letters)) throw new Error(`"${args.before}" is not a column letter`);
      if (columnToIndex(letters) > MAX_COLS) throw new Error("Column is outside the worksheet");
      const out = { type: "insert_column", sheet: args.sheet?.trim() || null, before: letters };
      if (typeof args.header === "string" && args.header.trim()) out.header = args.header.trim().slice(0, 255);
      out.summary = `Insert a column before ${letters}${out.header ? ` headed "${out.header}"` : ""}`;
      return out;
    },
  },

  sort_range: {
    description: "Sort a range by one of its columns. Include the header row in the range only if hasHeader is true.",
    params: {
      type: "object",
      properties: {
        address: { type: "string", description: 'Range to sort, e.g. "A1:F200".' },
        byColumn: { type: "string", description: 'Column letter to sort on, e.g. "C".' },
        ascending: { type: "boolean", description: "Defaults to true." },
        hasHeader: { type: "boolean", description: "True when the first row of the range is a header." },
        sheet: { type: "string" },
      },
      required: ["address", "byColumn"],
    },
    validate(args) {
      const range = parseRange(args.sheet ? `${args.sheet}!${args.address}` : args.address);
      const letters = String(args.byColumn ?? "").trim().toUpperCase();
      if (!/^[A-Z]{1,3}$/.test(letters)) throw new Error(`"${args.byColumn}" is not a column letter`);
      const colIndex = columnToIndex(letters);
      if (colIndex < range.start.col || colIndex > range.end.col) {
        throw new Error(`Column ${letters} is outside the range ${range.address}`);
      }
      const ascending = args.ascending !== false;
      return {
        type: "sort_range",
        sheet: range.sheet,
        address: range.address,
        byColumn: letters,
        offset: colIndex - range.start.col,
        ascending,
        hasHeader: args.hasHeader === true,
        summary: `Sort ${range.sheet ? `${range.sheet}!` : ""}${range.address} by column ${letters} ${ascending ? "A→Z" : "Z→A"}`,
      };
    },
  },
};

export const ACTION_NAMES = Object.keys(ACTION_SPECS);

/**
 * Drop top-level optional params the model left blank.
 *
 * Constrained decoding nudges models to emit every property in the schema, so they
 * routinely send `fill: ""` or `numberFormat: ""` meaning "not set". Treating those as
 * errors sent the model into a retry loop that burned the whole token budget. Blank
 * means absent.
 *
 * Only top-level scalars are stripped: `""` inside a `values` grid is a real
 * instruction to clear a cell.
 */
export function stripBlankParams(args) {
  const out = {};
  for (const [key, value] of Object.entries(args)) {
    if (value === null || value === undefined) continue;
    if (typeof value === "string" && value.trim() === "") continue;
    out[key] = value;
  }
  return out;
}

/**
 * Last row of the worksheet's used range, or null when unknown.
 * Used to keep a fill-down inside the data rather than padding to a round number.
 */
export function lastRowOf(sheetContext) {
  const address = sheetContext?.usedRange?.address;
  if (typeof address !== "string") return null;
  try {
    return parseRange(address).end.row;
  } catch {
    return null;
  }
}

/**
 * Validate raw tool arguments into a normalized action.
 * @param {string} name
 * @param {object} args
 * @param {{lastRow?: number|null}} [context] worksheet facts used to bound the action
 * @returns {{ok:true,action:object} | {ok:false,error:string}}
 */
export function validateAction(name, args, context = {}) {
  const spec = ACTION_SPECS[name];
  if (!spec) return { ok: false, error: `Unknown action "${name}"` };
  if (!args || typeof args !== "object") return { ok: false, error: `${name} received no arguments` };
  try {
    const cleaned = stripBlankParams(args);
    return {
      ok: true,
      action: {
        id: `${name}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        ...spec.validate(cleaned, context),
      },
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
