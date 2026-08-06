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
      "Write an Excel formula into a range. The same formula is written to every cell, so relative " +
      'references adjust per row exactly like dragging a fill handle. Always start with "=".',
    params: {
      type: "object",
      properties: {
        address: { type: "string", description: 'Target range in A1 notation, e.g. "E2:E100".' },
        formula: { type: "string", description: 'The formula, starting with "=", written for the first cell of the range.' },
        sheet: { type: "string", description: "Worksheet name. Omit for the active sheet." },
      },
      required: ["address", "formula"],
    },
    validate(args) {
      const range = parseRange(args.sheet ? `${args.sheet}!${args.address}` : args.address);
      const formula = String(args.formula ?? "").trim();
      if (!formula.startsWith("=")) throw new Error('formula must start with "="');
      if (formula.length > 8192) throw new Error("formula is too long");
      if (range.rows * range.cols > MAX_CELLS_PER_ACTION) {
        throw new Error(`That would fill ${range.rows * range.cols} cells; the limit is ${MAX_CELLS_PER_ACTION}`);
      }
      return {
        type: "write_formula",
        sheet: range.sheet,
        address: range.address,
        formula,
        summary: `Fill ${range.rows * range.cols} cell${range.rows * range.cols === 1 ? "" : "s"} in ${range.sheet ? `${range.sheet}!` : ""}${range.address} with ${formula}`,
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
    description: "Insert a new empty column before the given column letter, optionally with a header in row 1.",
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
 * Validate raw tool arguments into a normalized action.
 * @returns {{ok:true,action:object} | {ok:false,error:string}}
 */
export function validateAction(name, args) {
  const spec = ACTION_SPECS[name];
  if (!spec) return { ok: false, error: `Unknown action "${name}"` };
  if (!args || typeof args !== "object") return { ok: false, error: `${name} received no arguments` };
  try {
    const cleaned = stripBlankParams(args);
    return { ok: true, action: { id: `${name}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`, ...spec.validate(cleaned) } };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
