/**
 * Runs a whole-column transformation: read the source out of Excel, send every value to
 * the model, write the results back.
 *
 * The read and the write are separate `Excel.run` calls with the model work in between,
 * because that work can take minutes and holding a batch open across it would stall the
 * host.
 */
import { streamTransform, streamExtract } from "../api.js";
import { gridsMatch, UNDO_CHANGED_MESSAGE } from "./apply.js";
import {
  loadHeaderFormat,
  snapshotHeaderFormat,
  headerStyleFrom,
  applyHeaderStyle,
  restoreHeader,
  headerReferenceCell,
} from "./style.js";

const targetSheet = (ctx, name) =>
  name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();

const flatten = (values) => values.map((row) => (row[0] == null ? "" : String(row[0])));

const indexToColumn = (index) => {
  let n = index;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};

/**
 * Read the block (and header) back after writing, so undo can prove the cells still
 * hold exactly this before restoring — and refuse instead of wiping newer work.
 */
async function readBackWritten(action, address) {
  return Excel.run(async (ctx) => {
    const sheet = targetSheet(ctx, action.sheet);
    const range = sheet.getRange(address);
    range.load("formulas");
    const header = action.headerAddress ? sheet.getRange(action.headerAddress) : null;
    if (header) header.load("formulas");
    await ctx.sync();
    return { body: range.formulas, header: header ? header.formulas : null };
  });
}

/** The shared verify-then-restore undo for whole-range writes. */
async function undoWrittenBlock(action, address, written, snapshot) {
  await Excel.run(async (ctx) => {
    const sheet = targetSheet(ctx, action.sheet);
    const range = sheet.getRange(address);
    range.load("formulas");
    const header = action.headerAddress ? sheet.getRange(action.headerAddress) : null;
    if (header) header.load("formulas");
    await ctx.sync();

    if (!gridsMatch(range.formulas, written.body)) throw new Error(UNDO_CHANGED_MESSAGE);
    if (header && written.header && !gridsMatch(header.formulas, written.header)) {
      throw new Error(UNDO_CHANGED_MESSAGE);
    }

    range.formulas = snapshot.body;
    if (header && snapshot.header) restoreHeader(header, snapshot.header);
    await ctx.sync();
  });
}

/**
 * One snapshot pass before anything is written: the target block's formulas (for
 * Undo), the header row's current look (for Undo), and the style the new header
 * should take — continued from the sheet's own headers when it has any.
 */
async function snapshotBeforeWrite(action, address) {
  return Excel.run(async (ctx) => {
    const sheet = targetSheet(ctx, action.sheet);
    const range = sheet.getRange(address);
    range.load(["formulas"]);
    const header = action.headerAddress ? sheet.getRange(action.headerAddress) : null;
    if (header) loadHeaderFormat(header);
    const refAddress = action.headerAddress ? headerReferenceCell(action.headerAddress) : null;
    const ref = refAddress ? sheet.getRange(refAddress) : null;
    if (ref) loadHeaderFormat(ref);
    await ctx.sync();
    return {
      body: range.formulas,
      header: header ? snapshotHeaderFormat(header) : null,
      headerStyle: headerStyleFrom(ref ? snapshotHeaderFormat(ref) : null),
    };
  });
}

/**
 * @param {object} action  a validated transform_column action
 * @param {object} handlers { onProgress({done,total,sample}), onPhase(name) }
 * @returns {Promise<{undo:Function, message:string, failed:number[]}>}
 */
export async function runTransform(action, handlers = {}) {
  handlers.onPhase?.("reading");

  const source = await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(action.source);
    range.load(["values", "rowCount"]);
    await ctx.sync();
    return flatten(range.values);
  });

  // All-blank is as empty as zero rows — rewriting it would "succeed" by writing
  // blanks, which reads as data loss. The classic cause: columns were inserted
  // earlier in the same reply and the real data shifted right of the address.
  if (!source.length || source.every((v) => v.trim() === "")) {
    throw new Error(
      `the source ${action.source} is empty — if columns were just inserted, the data meant here has shifted right`,
    );
  }

  // Snapshot the target before touching it, so Undo restores whatever was there —
  // including the source itself when the transform rewrites in place.
  const snapshot = await snapshotBeforeWrite(action, action.address);

  handlers.onPhase?.("transforming");

  const { results, failed, lossy } = await new Promise((resolve, reject) => {
    const cancel = streamTransform(
      { values: source, instruction: action.instruction, fields: action.columns ?? null },
      {
        onProgress: (p) => handlers.onProgress?.(p),
        onDone: (payload) => resolve(payload),
        onError: (message) => reject(new Error(message)),
      },
    );
    handlers.onCancelHandle?.(cancel);
  });

  if (results.length !== source.length) {
    throw new Error(`The model returned ${results.length} rows for ${source.length} inputs; nothing was written.`);
  }

  handlers.onPhase?.("writing");

  await Excel.run(async (ctx) => {
    const sheet = targetSheet(ctx, action.sheet);
    // A split writes a grid and labels it; a rewrite writes a single column.
    sheet.getRange(action.address).values = action.columns
      ? results.map((row) => (Array.isArray(row) ? row : [row]))
      : results.map((v) => [v]);
    if (action.headerAddress) {
      const header = sheet.getRange(action.headerAddress);
      header.values = [action.columns];
      applyHeaderStyle(header, snapshot.headerStyle);
    }
    await ctx.sync();
  });

  const written = await readBackWritten(action, action.address);

  return {
    message: action.summary,
    failed,
    lossy: lossy ?? [],
    results,
    undo: () => undoWrittenBlock(action, action.address, written, snapshot),
  };
}

/**
 * Runs a whole-range record extraction: read the source out of Excel, let the model
 * find every record in it, write them as a table at the target.
 *
 * The output size is unknown until the model finishes — a range where each record
 * spans five rows shrinks, and a cell packing a dozen records grows. So the undo
 * snapshot is taken AFTER the model runs but BEFORE anything is written: by then the
 * record count, and therefore the exact target block, is known, and the sheet is
 * still untouched.
 *
 * @param {object} action  a validated extract_table action
 * @param {object} handlers { onProgress({done,total,records,sample}), onPhase(name), onCancelHandle(fn) }
 */
export async function runExtract(action, handlers = {}) {
  handlers.onPhase?.("reading");

  // Multi-column sources are read row by row; cells in a row are joined so each
  // line the model sees corresponds to one sheet row.
  const source = await Excel.run(async (ctx) => {
    const range = targetSheet(ctx, action.sheet).getRange(action.source);
    range.load(["values"]);
    await ctx.sync();
    return range.values.map((row) =>
      row.map((cell) => (cell == null ? "" : String(cell))).filter((s) => s.trim() !== "").join(" | "));
  });

  if (!source.some((line) => line.trim() !== "")) throw new Error("The source range is empty.");

  handlers.onPhase?.("transforming");

  const { records, failed } = await new Promise((resolve, reject) => {
    const cancel = streamExtract(
      { values: source, instruction: action.instruction, fields: action.columns },
      {
        onProgress: (p) => handlers.onProgress?.(p),
        onDone: (payload) => resolve(payload),
        onError: (message) => reject(new Error(message)),
      },
    );
    handlers.onCancelHandle?.(cancel);
  });

  if (!records.length) {
    throw new Error("No records were found in the source range. Nothing was written.");
  }

  const width = action.columns.length;
  const endCol = indexToColumn(action.targetCol + width - 1);
  const address = `${indexToColumn(action.targetCol)}${action.targetRow}:${endCol}${action.targetRow + records.length - 1}`;

  const snapshot = await snapshotBeforeWrite(action, address);

  handlers.onPhase?.("writing");

  await Excel.run(async (ctx) => {
    const sheet = targetSheet(ctx, action.sheet);
    sheet.getRange(address).values = records;
    if (action.headerAddress) {
      const header = sheet.getRange(action.headerAddress);
      header.values = [action.columns];
      applyHeaderStyle(header, snapshot.headerStyle);
    }
    await ctx.sync();
  });

  // The block is real now; point reveal, history, and the card at all of it.
  action.address = address;

  const written = await readBackWritten(action, address);

  return {
    message: `${records.length} record${records.length === 1 ? "" : "s"} written to ${address}`,
    failed,
    results: records,
    undo: () => undoWrittenBlock(action, address, written, snapshot),
  };
}
