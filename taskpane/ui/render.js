/**
 * DOM construction for the transcript.
 *
 * Everything is built with createElement and textContent rather than innerHTML — model
 * output is untrusted text and must never be parsed as markup.
 */
import { offsetFormula } from "./formula.js";

export const el = (tag, className, text) => {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
};

/** How many rows of the target range to show before collapsing into a count. */
const PREVIEW_ROWS = 5;

const rangeParts = (address) => {
  const [a, b = a] = address.split(":");
  const cell = (ref) => {
    const m = /^([A-Z]+)(\d+)$/i.exec(ref) ?? ["", "A", "1"];
    let col = 0;
    for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
    return { col, row: Number(m[2]), letters: m[1].toUpperCase() };
  };
  const start = cell(a);
  const end = cell(b);
  return { start, end, rows: end.row - start.row + 1, cols: end.col - start.col + 1 };
};

const columnLetter = (index) => {
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
 * Renders assistant prose. Supports only paragraphs and `inline code`, which is all the
 * model is asked to produce — no HTML is ever interpreted.
 */
export function renderProse(container, text) {
  container.textContent = "";
  for (const block of text.split(/\n{2,}/)) {
    if (!block.trim()) continue;
    const p = el("p");
    // Split on backtick spans, keeping the delimiters' contents.
    const parts = block.split(/(`[^`]+`)/g);
    for (const part of parts) {
      if (part.startsWith("`") && part.endsWith("`") && part.length > 2) {
        p.appendChild(el("code", null, part.slice(1, -1)));
      } else if (part) {
        p.appendChild(document.createTextNode(part));
      }
    }
    container.appendChild(p);
  }
}

/** The mini-spreadsheet: the point is that approving a change means seeing it. */
function buildGrid(action) {
  const { start, rows, cols } = rangeParts(action.address);
  const wrap = el("div", "grid");
  const table = el("table");

  const headRow = el("tr");
  headRow.appendChild(el("th", null, ""));
  const shownCols = Math.min(cols, 6);
  for (let c = 0; c < shownCols; c++) headRow.appendChild(el("th", null, columnLetter(start.col + c)));
  if (cols > shownCols) headRow.appendChild(el("th", null, "…"));
  table.appendChild(el("thead")).appendChild(headRow);

  const body = el("tbody");
  const shownRows = Math.min(rows, PREVIEW_ROWS);

  for (let r = 0; r < shownRows; r++) {
    const tr = el("tr");
    tr.appendChild(el("td", "grid__rownum", String(start.row + r)));
    for (let c = 0; c < shownCols; c++) {
      let value = "";
      if (action.type === "write_formula") value = offsetFormula(action.formula, r, c);
      else if (action.type === "write_values") value = String(action.values?.[r]?.[c] ?? "");
      const td = el("td", "grid__target", value);
      td.title = value;
      tr.appendChild(td);
    }
    if (cols > shownCols) tr.appendChild(el("td", "grid__target", "…"));
    body.appendChild(tr);
  }
  table.appendChild(body);
  wrap.appendChild(table);

  const hidden = rows - shownRows;
  if (hidden > 0) {
    wrap.appendChild(el("div", "grid__more", `+ ${hidden.toLocaleString()} more row${hidden === 1 ? "" : "s"}`));
  }
  return wrap;
}

/** Compact property list for actions that change appearance or structure, not content. */
function buildDetails(action) {
  const list = el("div", "proposal__summary");
  const bits = [];
  if (action.numberFormat) bits.push(`Number format ${action.numberFormat}`);
  if (action.bold !== undefined) bits.push(action.bold ? "Bold" : "Not bold");
  if (action.italic !== undefined) bits.push(action.italic ? "Italic" : "Not italic");
  if (action.fill) bits.push(`Fill ${action.fill}`);
  if (action.header) bits.push(`Header "${action.header}"`);
  if (action.byColumn) bits.push(`Sort on column ${action.byColumn}, ${action.ascending ? "A→Z" : "Z→A"}`);
  list.textContent = bits.length ? bits.join(" · ") : action.summary;
  return list;
}

/**
 * Build the proposed-change card.
 * @param {object} action
 * @param {{onApply:Function,onDismiss:Function,onUndo:Function}} handlers
 */
export function buildProposal(action, handlers) {
  const card = el("div", "proposal");
  card.dataset.status = "pending";

  const head = el("div", "proposal__head");
  head.appendChild(el("span", "eyebrow proposal__label", "Proposed change"));
  const addr = el("span", "proposal__addr", action.sheet ? `${action.sheet}!${action.address ?? ""}` : (action.address ?? ""));
  head.appendChild(addr);
  card.appendChild(head);

  const showsGrid = action.type === "write_formula" || action.type === "write_values";
  card.appendChild(showsGrid ? el("div", "proposal__summary", action.summary) : buildDetails(action));
  if (showsGrid) card.appendChild(buildGrid(action));

  const actions = el("div", "proposal__actions");
  const apply = el("button", "btn btn--primary", "Apply");
  const dismiss = el("button", "btn btn--quiet", "Dismiss");
  actions.append(apply, dismiss);
  card.appendChild(actions);

  const finish = (label, undo) => {
    card.dataset.status = "applied";
    head.querySelector(".proposal__label").textContent = label;
    actions.remove();
    const receipt = el("div", "proposal__receipt");
    receipt.appendChild(el("span", null, label === "Applied" ? "Change written to the sheet" : "Change dismissed"));
    if (undo) {
      const undoBtn = el("button", "link", "Undo");
      undoBtn.addEventListener("click", async () => {
        undoBtn.disabled = true;
        undoBtn.textContent = "Undoing…";
        try {
          await undo();
          receipt.textContent = "";
          receipt.appendChild(el("span", null, "Change undone"));
          card.dataset.status = "pending";
          head.querySelector(".proposal__label").textContent = "Undone";
        } catch (err) {
          undoBtn.disabled = false;
          undoBtn.textContent = "Undo";
          receipt.querySelector("span").textContent = `Could not undo: ${err.message}`;
        }
      });
      receipt.appendChild(undoBtn);
    }
    card.appendChild(receipt);
  };

  apply.addEventListener("click", async () => {
    apply.disabled = true;
    dismiss.disabled = true;
    apply.textContent = "Applying…";
    try {
      const { undo } = await handlers.onApply(action);
      finish("Applied", undo);
    } catch (err) {
      apply.disabled = false;
      dismiss.disabled = false;
      apply.textContent = "Apply";
      const error = card.querySelector(".proposal__error") ?? el("div", "proposal__summary proposal__error");
      error.textContent = `Could not apply: ${err.message}`;
      error.style.color = "var(--danger)";
      card.insertBefore(error, actions);
    }
  });

  dismiss.addEventListener("click", () => {
    handlers.onDismiss?.(action);
    finish("Dismissed", null);
  });

  return card;
}

/** A user or assistant turn. Returns the body node so streaming can append into it. */
export function buildTurn(role) {
  const turn = el("div", `turn turn--${role}`);
  const body = el("div", "turn__body");
  turn.appendChild(body);
  return { turn, body };
}
