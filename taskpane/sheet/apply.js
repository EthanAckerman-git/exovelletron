/**
 * Applies an approved action to the workbook, capturing enough prior state to undo it.
 *
 * Excel's own undo stack does not reliably include add-in writes, so we snapshot what
 * we are about to overwrite and offer our own one-step undo. A user who approves a
 * 500-cell fill and then regrets it must be able to get their sheet back.
 */

import { offsetFormula } from "../ui/formula.js";
import {
  loadHeaderFormat,
  snapshotHeaderFormat,
  headerStyleFrom,
  applyHeaderStyle,
  headerReferenceCell,
} from "./style.js";

const targetSheet = (ctx, name) =>
  name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();

/**
 * Undo must never destroy work that arrived after the apply. Every content undo
 * therefore re-reads its range and only restores the snapshot when the cells still
 * hold exactly what the apply wrote; anything else means a later change — another
 * card's output, or the user's own typing — now lives there, and the undo refuses.
 */
export const UNDO_CHANGED_MESSAGE =
  "the sheet has changed here since this was applied — undo would overwrite that newer work, so nothing was touched";

/** Deep-compare two formula grids the way Excel hands them back. */
export const gridsMatch = (a, b) =>
  Array.isArray(a) && Array.isArray(b) && a.length === b.length &&
  a.every((row, r) => Array.isArray(row) && Array.isArray(b[r]) && row.length === b[r].length &&
    row.every((cell, c) => String(cell) === String(b[r][c])));

/** Read a range's formulas back after a write, so undo can prove nothing moved. */
async function readBack(action, address) {
  return Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(address);
    range.load("formulas");
    await ctx.sync();
    return range.formulas;
  });
}

/** Restore a snapshot only when the range still holds exactly what was written. */
async function restoreIfUnchanged(action, address, written, snapshot) {
  await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(address);
    range.load("formulas");
    await ctx.sync();
    if (!gridsMatch(range.formulas, written)) throw new Error(UNDO_CHANGED_MESSAGE);
    range.formulas = snapshot;
    await ctx.sync();
  });
}

const supportsCopyFrom = () => {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", "1.9");
  } catch {
    return false;
  }
};

const dims = (address) => {
  const [a, b = a] = address.split(":");
  const cell = (ref) => {
    const m = /^([A-Z]+)(\d+)$/i.exec(ref);
    let col = 0;
    for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
    return { col, row: Number(m[2]) };
  };
  const s = cell(a);
  const e = cell(b);
  return { rows: Math.abs(e.row - s.row) + 1, cols: Math.abs(e.col - s.col) + 1 };
};

/**
 * @param {object} action  a validated action from the engine
 * @returns {Promise<{undo: (() => Promise<void>)|null, message: string}>}
 */
export async function applyAction(action) {
  switch (action.type) {
    case "write_values":
      return applyWrite(action, (range) => { range.values = action.values; });

    case "write_formula":
      return applyFormulaFill(action);

    case "format_cells":
      return applyFormat(action);

    case "insert_column":
      return applyInsertColumn(action);

    case "sort_range":
      return applySort(action);

    default:
      throw new Error(`Cannot apply unsupported action "${action.type}"`);
  }
}

/**
 * Fill a formula down a range with per-cell relative-reference adjustment.
 *
 * Assigning the same string to every cell of `range.formulas` does NOT behave like
 * dragging the fill handle — Excel writes each formula verbatim, so every row ends up
 * computing the anchor row and shows an identical result. The fix is to write the anchor
 * cell and let Excel copy it across, which applies its own relative-reference rules.
 *
 * `copyFrom` needs ExcelApi 1.9; on anything older we compute the offsets ourselves.
 */
async function applyFormulaFill(action) {
  const { rows, cols } = dims(action.address);

  const snapshot = await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(action.address);
    range.load(["formulas"]);
    await ctx.sync();
    return range.formulas;
  });

  await Excel.run(async (ctx) => {
    const sheet = targetSheet(ctx, action.sheet);
    const target = sheet.getRange(action.address);

    if (supportsCopyFrom()) {
      const anchor = target.getCell(0, 0);
      anchor.formulas = [[action.formula]];
      target.copyFrom(anchor, Excel.RangeCopyType.formulas);
    } else {
      target.formulas = Array.from({ length: rows }, (_, r) =>
        Array.from({ length: cols }, (_, c) => offsetFormula(action.formula, r, c)));
    }
    await ctx.sync();
  });

  const written = await readBack(action, action.address);

  return {
    message: action.summary,
    undo: () => restoreIfUnchanged(action, action.address, written, snapshot),
  };
}

/** Shared path for the two value-writing actions: snapshot formulas, then write. */
async function applyWrite(action, mutate) {
  const snapshot = await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(action.address);
    // Snapshot formulas rather than values: restoring formulas restores values too,
    // whereas restoring values would silently flatten a formula into a literal.
    range.load(["formulas"]);
    await ctx.sync();
    return range.formulas;
  });

  await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(action.address);
    mutate(range);
    await ctx.sync();
  });

  const written = await readBack(action, action.address);

  return {
    message: action.summary,
    undo: () => restoreIfUnchanged(action, action.address, written, snapshot),
  };
}

async function applyFormat(action) {
  const snapshot = await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(action.address);
    range.load(["numberFormat"]);
    range.format.fill.load("color");
    range.format.font.load(["bold", "italic"]);
    await ctx.sync();
    return {
      numberFormat: range.numberFormat,
      fill: range.format.fill.color,
      bold: range.format.font.bold,
      italic: range.format.font.italic,
    };
  });

  await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(action.address);
    if (action.numberFormat !== undefined) {
      const { rows, cols } = dims(action.address);
      range.numberFormat = Array.from({ length: rows }, () => Array.from({ length: cols }, () => action.numberFormat));
    }
    if (action.bold !== undefined) range.format.font.bold = action.bold;
    if (action.italic !== undefined) range.format.font.italic = action.italic;
    if (action.fill !== undefined) range.format.fill.color = action.fill;
    await ctx.sync();
  });

  return {
    message: action.summary,
    undo: async () => {
      await Excel.run(async (ctx) => {
        const range = targetSheet(ctx, action.sheet).getRange(action.address);
        range.numberFormat = snapshot.numberFormat;
        range.format.font.bold = snapshot.bold;
        range.format.font.italic = snapshot.italic;
        // A null fill means "no fill"; clear() is the only way back to that state.
        if (snapshot.fill) range.format.fill.color = snapshot.fill;
        else range.format.fill.clear();
        await ctx.sync();
      });
    },
  };
}

async function applyInsertColumn(action) {
  await Excel.run(async (ctx) => {
    const sheet = targetSheet(ctx, action.sheet);
    sheet.getRange(`${action.before}:${action.before}`).insert(Excel.InsertShiftDirection.right);
    if (action.header) {
      // The new header should look like the sheet's existing ones. The cell left of
      // the inserted column keeps its position through the shift, so sample it.
      const refAddress = headerReferenceCell(`${action.before}1`);
      const ref = refAddress ? sheet.getRange(refAddress) : null;
      if (ref) loadHeaderFormat(ref);
      await ctx.sync();

      const header = sheet.getRange(`${action.before}1`);
      header.values = [[action.header]];
      applyHeaderStyle(header, headerStyleFrom(ref ? snapshotHeaderFormat(ref) : null));
    }
    await ctx.sync();
  });

  return {
    message: action.summary,
    // Deleting a column erases whatever it holds AND shifts every column after it.
    // That is only a clean revert while the column still contains nothing but what
    // this action created — the header, or nothing at all. Once anything else has
    // been written into it (a later card's output, the user's typing), undo refuses.
    undo: async () => {
      await Excel.run(async (ctx) => {
        const sheet = targetSheet(ctx, action.sheet);
        const column = sheet.getRange(`${action.before}:${action.before}`);
        const used = column.getUsedRangeOrNullObject(true);
        used.load(["isNullObject", "formulas", "rowIndex", "rowCount"]);
        await ctx.sync();

        if (!used.isNullObject) {
          const onlyTheHeader =
            action.header &&
            used.rowIndex === 0 &&
            used.rowCount === 1 &&
            String(used.formulas[0][0]) === String(action.header);
          const empty = used.formulas.every((row) => row.every((cell) => String(cell) === ""));
          if (!onlyTheHeader && !empty) {
            throw new Error(
              "the new column has been filled since — deleting it now would erase that work, so nothing was touched",
            );
          }
        }

        column.delete(Excel.DeleteShiftDirection.left);
        await ctx.sync();
      });
    },
  };
}

async function applySort(action) {
  // Sorting permutes rows irreversibly, so the snapshot is the whole range.
  const snapshot = await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(action.address);
    range.load(["formulas"]);
    await ctx.sync();
    return range.formulas;
  });

  await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(action.address);
    range.sort.apply(
      [{ key: action.offset, ascending: action.ascending }],
      false,
      action.hasHeader,
      Excel.SortOrientation.rows,
    );
    await ctx.sync();
  });

  const written = await readBack(action, action.address);

  return {
    message: action.summary,
    undo: () => restoreIfUnchanged(action, action.address, written, snapshot),
  };
}

/** Scroll to and select the range an action touched, so the user can see the result. */
export async function revealRange(action) {
  if (!action.address && !action.before) return;
  const address = action.address ?? `${action.before}:${action.before}`;
  await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(address);
    range.select();
    await ctx.sync();
  });
}
