/**
 * Executes the model's mid-turn workbook reads.
 *
 * Same discipline as context.js: never load values whose size is unknown. The
 * requested range is first intersected with the sheet's used range — so "A1:Z10000"
 * on a 200-row sheet reads 200 rows, not ten thousand — and the cell count is checked
 * before any values cross the Office bridge.
 *
 * Never throws: every failure comes back as { ok: false, error } in words the model
 * can act on, because the text is relayed straight into its turn.
 */

/** The most cells one read may pull across the bridge. */
export const MAX_READ_CELLS = 2000;

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

const plain = (v) => (v === null || v === undefined ? "" : v);

/**
 * Split a possibly sheet-qualified address ("Sheet2!A1:B2") into its parts. The model
 * is told to pass the sheet separately, but qualifying the address is a habit worth
 * absorbing rather than rejecting.
 */
export function splitReadAddress({ sheet, address }) {
  const raw = String(address ?? "").trim();
  const bang = raw.lastIndexOf("!");
  if (bang === -1) return { sheet: sheet || null, ref: raw };
  return {
    sheet: raw.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'") || sheet || null,
    ref: raw.slice(bang + 1).trim(),
  };
}

/**
 * @param {{sheet:string|null, address:string}} request
 * @returns {Promise<{ok:true, sheet:string, address:string, startRow:number,
 *   columnLetters:string[], rows:any[][]} | {ok:false, error:string}>}
 */
export async function readRangeForModel(request) {
  const { sheet, ref } = splitReadAddress(request);
  if (!ref) return { ok: false, error: "no range was given — pass A1 notation like \"A2:D40\"" };

  try {
    return await Excel.run(async (ctx) => {
      const ws = sheet
        ? ctx.workbook.worksheets.getItemOrNullObject(sheet)
        : ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      const used = ws.getUsedRangeOrNullObject(true);
      used.load("isNullObject");
      await ctx.sync();

      if (ws.isNullObject) {
        return { ok: false, error: `there is no sheet named "${sheet}" — the sheet list in your context has the real names` };
      }
      if (used.isNullObject) {
        return { ok: true, sheet: ws.name, address: ref, startRow: 1, columnLetters: [], rows: [] };
      }

      // The sync below is also where Excel rejects a malformed address, so the shape
      // is known — and bounded — before any values are requested.
      const target = ws.getRange(ref).getIntersectionOrNullObject(used);
      target.load(["address", "isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount", "cellCount"]);
      await ctx.sync();

      if (target.isNullObject) {
        return { ok: true, sheet: ws.name, address: ref, startRow: 1, columnLetters: [], rows: [] };
      }
      if (target.cellCount > MAX_READ_CELLS) {
        return {
          ok: false,
          error: `that is ${target.cellCount} cells of data; read at most ${MAX_READ_CELLS} at a time — split the range`,
        };
      }

      target.load("values");
      await ctx.sync();

      return {
        ok: true,
        sheet: ws.name,
        address: (target.address.split("!").pop() ?? ref).trim(),
        startRow: target.rowIndex + 1,
        columnLetters: Array.from({ length: target.columnCount }, (_, i) => columnLetter(target.columnIndex + i)),
        rows: target.values.map((r) => r.map(plain)),
      };
    });
  } catch (err) {
    const error = err?.code === "InvalidArgument" || err?.code === "ItemNotFound"
      ? `"${ref}" is not a range Excel understands — use A1 notation like "A2:D40"`
      : err?.message || "the read failed";
    return { ok: false, error };
  }
}
