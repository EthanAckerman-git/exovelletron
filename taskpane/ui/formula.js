/**
 * Display-only formula shifting for the change preview.
 *
 * When a formula is filled down a column, Excel adjusts relative row references per
 * row. The preview shows the formula each row will actually receive, so the user can
 * check row 3 rather than extrapolating from row 2 in their head.
 *
 * This is presentation only — the write itself uses Excel's own fill semantics.
 */

/**
 * Matches an A1 reference while refusing the lookalikes that break naive versions:
 *  - `LOG10(` — a function name ending in digits, rejected by the trailing `(` guard
 *  - `A1B2`   — rejected because a letter may not follow the digits
 *  - `$A$1`   — captured with its absolute markers so they can be honoured
 */
const REF = /(?<![A-Z0-9_$.])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\d(A-Z_])/gi;

/** Split a formula into quoted-string and code segments so literals stay untouched. */
function segments(formula) {
  const out = [];
  let buf = "";
  let inString = false;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '"') {
      // "" inside a string is an escaped quote, not a terminator.
      if (inString && formula[i + 1] === '"') {
        buf += '""';
        i++;
        continue;
      }
      buf += ch;
      if (inString) {
        out.push({ text: buf, code: false });
        buf = "";
        inString = false;
      } else {
        if (buf.length > 1) out.push({ text: buf.slice(0, -1), code: true });
        buf = '"';
        inString = true;
      }
      continue;
    }
    buf += ch;
  }
  if (buf) out.push({ text: buf, code: !inString });
  return out;
}

/**
 * Shift relative references by whole rows/columns.
 * @param {string} formula
 * @param {number} rowDelta
 * @param {number} [colDelta]
 */
export function offsetFormula(formula, rowDelta, colDelta = 0) {
  if (typeof formula !== "string" || (!rowDelta && !colDelta)) return formula;

  return segments(formula)
    .map(({ text, code }) => {
      if (!code) return text;
      return text.replace(REF, (match, colAbs, col, rowAbs, row) => {
        let nextCol = col;
        let nextRow = row;

        if (!rowAbs && rowDelta) {
          const shifted = Number(row) + rowDelta;
          if (shifted < 1 || shifted > 1_048_576) return "#REF!";
          nextRow = String(shifted);
        }
        if (!colAbs && colDelta) {
          let index = 0;
          for (const ch of col.toUpperCase()) index = index * 26 + (ch.charCodeAt(0) - 64);
          const shifted = index + colDelta;
          if (shifted < 1 || shifted > 16384) return "#REF!";
          let n = shifted;
          nextCol = "";
          while (n > 0) {
            const rem = (n - 1) % 26;
            nextCol = String.fromCharCode(65 + rem) + nextCol;
            n = Math.floor((n - 1) / 26);
          }
        }
        return `${colAbs}${nextCol}${rowAbs}${nextRow}`;
      });
    })
    .join("");
}
