/**
 * Reads a bounded, model-sized view of the active worksheet.
 *
 * The important discipline here is never loading the used range's values. A workbook
 * with 200,000 rows would serialise hundreds of megabytes across the Office bridge and
 * hang the host. Instead we load only the *shape* first, then read a small explicit
 * window out of it.
 */

/**
 * Hard ceilings on what we will pull across the bridge, regardless of sheet size.
 *
 * `tailRows` matters more than its size suggests: showing only the first rows of a
 * 40,000-row sheet leaves the model guessing whether the data continues in the same
 * shape. Sampling the end as well makes the scale concrete and catches the common case
 * where later rows are formatted differently from the first screenful.
 */
export const LIMITS = {
  // Half of what it once was: the model can now read_range any rows it is missing,
  // and a leaner prompt is ingested noticeably faster on every single turn.
  sampleRows: 30,
  tailRows: 15,
  columns: 40,
  selectionRows: 40,
  formulaProbeRows: 12,
  /** How many non-active sheets the workbook map describes. */
  mappedSheets: 11,
};

const splitAddress = (address = "") => {
  const bang = address.lastIndexOf("!");
  return bang === -1
    ? { sheet: null, ref: address }
    : { sheet: address.slice(0, bang).replace(/^'|'$/g, ""), ref: address.slice(bang + 1) };
};

const columnLetter = (index) => {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};

/** Excel returns dates as serial numbers; keep cells as plain scalars for the model. */
const plain = (v) => (v === null || v === undefined ? "" : v);

/**
 * @returns {Promise<object|null>} worksheet context, or null when nothing is readable
 */
export async function readSheetContext() {
  return Excel.run(async (ctx) => {
    const workbook = ctx.workbook;
    const sheet = workbook.worksheets.getActiveWorksheet();
    const sheets = workbook.worksheets;
    const used = sheet.getUsedRangeOrNullObject(true);
    const selection = workbook.getSelectedRange();

    workbook.load("name");
    sheet.load("name");
    sheets.load("items/name");
    // Shape only — deliberately not values.
    used.load(["address", "rowCount", "columnCount", "isNullObject", "rowIndex", "columnIndex"]);
    selection.load(["address", "rowCount", "columnCount", "cellCount", "rowIndex", "columnIndex"]);
    await ctx.sync();

    const result = {
      workbookName: workbook.name,
      sheetName: sheet.name,
      sheetNames: sheets.items.map((s) => s.name),
      usedRange: null,
      selection: null,
      headers: null,
      columnLetters: [],
      sample: null,
      selectionSample: null,
      formulas: [],
      sheets: [],
    };

    // The workbook map: shape and header row of every other sheet, so the model knows
    // the whole workbook exists — read_range fills in whatever values it then needs.
    const others = sheets.items
      .filter((s) => s.name !== sheet.name)
      .slice(0, LIMITS.mappedSheets)
      .map((s) => {
        const usedRange = s.getUsedRangeOrNullObject(true);
        usedRange.load(["address", "rowCount", "columnCount", "isNullObject", "rowIndex", "columnIndex"]);
        return { name: s.name, worksheet: s, usedRange };
      });

    {
      const { ref } = splitAddress(selection.address);
      result.selection = {
        address: ref,
        rowCount: selection.rowCount,
        columnCount: selection.columnCount,
        isSingleCell: selection.rowCount === 1 && selection.columnCount === 1,
      };
    }

    // Resolves the other sheets' shapes (already loaded by the last sync), reads each
    // one's first row, and fills result.sheets. One extra sync, only when needed.
    const completeSheetMap = async () => {
      const withData = others.filter((o) => !o.usedRange.isNullObject);
      const headerRanges = withData.map((o) => {
        const u = o.usedRange;
        const header = o.worksheet.getRangeByIndexes(u.rowIndex, u.columnIndex, 1, Math.min(u.columnCount, LIMITS.columns));
        header.load("values");
        return header;
      });
      if (withData.length) await ctx.sync();

      result.sheets = others.map((o) => {
        const u = o.usedRange;
        if (u.isNullObject) return { name: o.name, rowCount: 0 };
        const colCount = Math.min(u.columnCount, LIMITS.columns);
        return {
          name: o.name,
          address: splitAddress(u.address).ref,
          rowCount: u.rowCount,
          columnCount: u.columnCount,
          columnLetters: Array.from({ length: colCount }, (_, i) => columnLetter(u.columnIndex + i)),
          headers: headerRanges[withData.indexOf(o)]?.values?.[0]?.map(plain) ?? [],
        };
      });
    };

    if (used.isNullObject) {
      // The active sheet is empty, but the rest of the workbook may not be.
      await ctx.sync();
      await completeSheetMap();
      return result;
    }

    const { ref: usedRef } = splitAddress(used.address);
    result.usedRange = { address: usedRef, rowCount: used.rowCount, columnCount: used.columnCount };

    const colCount = Math.min(used.columnCount, LIMITS.columns);
    const startRow = used.rowIndex;
    const startCol = used.columnIndex;
    result.columnLetters = Array.from({ length: colCount }, (_, i) => columnLetter(startCol + i));

    // Header row plus a window of body rows, in one bounded read.
    const windowRows = Math.min(used.rowCount, LIMITS.sampleRows + 1);
    const window = sheet.getRangeByIndexes(startRow, startCol, windowRows, colCount);
    window.load(["values"]);

    // When the sheet runs past the head window, also read the last few rows so the model
    // can see how the data actually ends rather than extrapolating from the top.
    let tail = null;
    let tailStartRow = 0;
    const tailCount = Math.min(LIMITS.tailRows, used.rowCount - windowRows);
    if (tailCount > 0) {
      tailStartRow = startRow + used.rowCount - tailCount;
      tail = sheet.getRangeByIndexes(tailStartRow, startCol, tailCount, colCount);
      tail.load(["values"]);
    }

    // Probe a few cells for existing formulas so the model can respect them.
    const probeRows = Math.min(used.rowCount, LIMITS.formulaProbeRows);
    const probe = sheet.getRangeByIndexes(startRow, startCol, probeRows, colCount);
    probe.load(["formulas"]);

    let selectionRange = null;
    if (selection.cellCount > 1 && selection.rowCount <= LIMITS.selectionRows && selection.columnCount <= LIMITS.columns) {
      selectionRange = selection;
      selectionRange.load(["values"]);
    }

    await ctx.sync();
    await completeSheetMap();

    const rows = window.values.map((r) => r.map(plain));
    const looksLikeHeader =
      rows.length > 1 && rows[0].every((c) => typeof c === "string" || c === "") && rows[0].some((c) => c !== "");

    if (looksLikeHeader) {
      result.headers = rows[0];
      result.sample = { startRow: startRow + 2, columnLetters: result.columnLetters, rows: rows.slice(1) };
    } else {
      result.sample = { startRow: startRow + 1, columnLetters: result.columnLetters, rows };
    }

    if (tail) {
      result.tail = {
        startRow: tailStartRow + 1,
        columnLetters: result.columnLetters,
        rows: tail.values.map((r) => r.map(plain)),
      };
    }

    for (let r = 0; r < probe.formulas.length; r++) {
      for (let c = 0; c < probe.formulas[r].length; c++) {
        const f = probe.formulas[r][c];
        if (typeof f === "string" && f.startsWith("=")) {
          result.formulas.push({ address: `${result.columnLetters[c]}${startRow + r + 1}`, formula: f });
        }
      }
      if (result.formulas.length >= 12) break;
    }

    if (selectionRange) {
      const { ref } = splitAddress(selection.address);
      const firstRow = Number((/\d+/.exec(ref) || [String(selection.rowIndex + 1)])[0]);
      result.selectionSample = {
        startRow: firstRow,
        columnLetters: Array.from({ length: selection.columnCount }, (_, i) => columnLetter(selection.columnIndex + i)),
        rows: selectionRange.values.map((r) => r.map(plain)),
      };
    }

    return result;
  });
}

/** Short label for the selection strip, e.g. "Sales!E2:E501". */
export function describeSelection(context) {
  if (!context?.selection) return { address: "No selection", meta: "" };
  const { address, rowCount, columnCount, isSingleCell } = context.selection;
  return {
    address: `${context.sheetName}!${address}`,
    meta: isSingleCell ? "1 cell" : `${rowCount} × ${columnCount}`,
  };
}
