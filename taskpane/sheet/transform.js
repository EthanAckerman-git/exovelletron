/**
 * Runs a whole-column transformation: read the source out of Excel, send every value to
 * the model, write the results back.
 *
 * The read and the write are separate `Excel.run` calls with the model work in between,
 * because that work can take minutes and holding a batch open across it would stall the
 * host.
 */
import { streamTransform } from "../api.js";

const targetSheet = (ctx, name) =>
  name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();

const flatten = (values) => values.map((row) => (row[0] == null ? "" : String(row[0])));

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

  if (!source.length) throw new Error("The source column is empty.");

  // Snapshot the target before touching it, so Undo restores whatever was there —
  // including the source itself when the transform rewrites in place.
  const snapshot = await Excel.run(async (ctx) => {
    const sheet = targetSheet(ctx, action.sheet);
    const range = sheet.getRange(action.address);
    range.load(["formulas"]);
    const header = action.headerAddress ? sheet.getRange(action.headerAddress) : null;
    header?.load(["formulas"]);
    await ctx.sync();
    return { body: range.formulas, header: header ? header.formulas : null };
  });

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
    if (action.headerAddress) sheet.getRange(action.headerAddress).values = [action.columns];
    await ctx.sync();
  });

  return {
    message: action.summary,
    failed,
    lossy: lossy ?? [],
    results,
    undo: async () => {
      await Excel.run(async (ctx) => {
        const sheet = targetSheet(ctx, action.sheet);
        sheet.getRange(action.address).formulas = snapshot.body;
        if (action.headerAddress && snapshot.header) {
          sheet.getRange(action.headerAddress).formulas = snapshot.header;
        }
        await ctx.sync();
      });
    },
  };
}
