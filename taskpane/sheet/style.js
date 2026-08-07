/**
 * Header styling for tables the app writes.
 *
 * A header row the AI adds should look like it belongs to the sheet it lands on. So the
 * style is sampled from the sheet itself — the cell just left of the new headers, which
 * on any sheet that already has styled headers is one of them — and only when there is
 * nothing to continue does the default kick in. The default is Excel's own table-header
 * blue with white bold text, so even a bare sheet gets a properly coloured top row
 * rather than plain text.
 *
 * Pure functions plus Office.js proxy helpers; the pure parts are unit-tested.
 */

export const DEFAULT_HEADER_STYLE = Object.freeze({
  fill: "#4472C4",
  fontColor: "#FFFFFF",
  bold: true,
});

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
 * The cell just left of a header row — on a sheet that already has styled headers,
 * that is one of them, and its look is what the new headers should continue.
 * @returns {string|null} e.g. "A1" for a header row starting at B1; null in column A
 */
export function headerReferenceCell(headerAddress) {
  const m = /^([A-Z]+)(\d+)/i.exec(headerAddress ?? "");
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return col > 1 ? `${indexToColumn(col - 1)}${m[2]}` : null;
}

/** Queue the format properties needed to snapshot or sample a range. */
export function loadHeaderFormat(range) {
  range.load(["formulas"]);
  range.format.fill.load("color");
  range.format.font.load(["bold", "color"]);
}

/** Capture a loaded range's look so Undo can put it back exactly. */
export const snapshotHeaderFormat = (range) => ({
  formulas: range.formulas,
  fill: range.format.fill.color,
  bold: range.format.font.bold,
  fontColor: range.format.font.color,
});

/** True when a sampled fill is a real colour rather than "no fill" (read back as white). */
const isRealFill = (fill) =>
  typeof fill === "string" && fill.trim() !== "" && !/^#?FFFFFF$/i.test(fill.trim());

/**
 * Decide the style for a new header row from a sampled neighbour.
 *
 * @param {{fill?:string, fontColor?:string, bold?:boolean}|null} reference
 *   snapshot of the cell left of the new headers, or null when there is none
 * @returns {{fill:string, fontColor:string, bold:boolean}}
 */
export function headerStyleFrom(reference) {
  if (reference && isRealFill(reference.fill)) {
    return {
      fill: reference.fill,
      fontColor: typeof reference.fontColor === "string" && reference.fontColor ? reference.fontColor : "#FFFFFF",
      bold: reference.bold !== false,
    };
  }
  return DEFAULT_HEADER_STYLE;
}

/** Paint a header range with the chosen style. */
export function applyHeaderStyle(range, style) {
  range.format.font.bold = style.bold;
  range.format.font.color = style.fontColor;
  range.format.fill.color = style.fill;
}

/** Put a header range back exactly as the snapshot recorded it. */
export function restoreHeader(range, saved) {
  range.formulas = saved.formulas;
  range.format.font.bold = saved.bold;
  if (saved.fontColor) range.format.font.color = saved.fontColor;
  if (isRealFill(saved.fill)) range.format.fill.color = saved.fill;
  else range.format.fill.clear();
}
