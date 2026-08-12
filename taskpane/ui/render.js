/**
 * DOM construction for the transcript.
 *
 * Everything is built with createElement and textContent rather than innerHTML — model
 * output is untrusted text and must never be parsed as markup.
 */
import { offsetFormula } from "./formula.js";
import { renderMarkdown } from "./markdown.js";

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
 * Renders an assistant reply as Markdown.
 *
 * Built entirely from DOM nodes — model output is untrusted and is never parsed as HTML.
 */
export function renderProse(container, text) {
  renderMarkdown(container, text);
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

/** One result row as a single readable line: split fields are joined with a separator. */
const formatSampleRow = (value) =>
  (Array.isArray(value) ? value.map((v) => v || "—").join("  ·  ") : value) || "(blank)";

/**
 * A whole-column rewrite has no preview grid to show up front: the results do not exist
 * until the job runs. So the card states exactly what will happen, then fills in real
 * output as rows complete.
 */
function buildTransformBody(action) {
  const wrap = el("div");
  wrap.appendChild(el("div", "proposal__summary", action.summary));

  const instruction = el("div", "proposal__instruction");
  instruction.appendChild(el("span", "eyebrow", action.columns ? "New columns" : "Rule applied to every row"));
  if (action.columns) instruction.appendChild(el("p", "mono", action.columns.join("  ·  ")));
  instruction.appendChild(el("p", null, action.instruction));
  wrap.appendChild(instruction);

  const progress = el("div", "transform__progress");
  progress.hidden = true;
  const bar = el("div", "progress");
  bar.appendChild(el("i"));
  const label = el("div", "transform__count mono", "");
  progress.append(bar, label);
  wrap.appendChild(progress);

  const sample = el("div", "transform__sample");
  sample.hidden = true;
  wrap.appendChild(sample);

  return { wrap, progress, bar: bar.firstChild, label, sample };
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
 * Enable the action buttons on every card still waiting for its reply to finish.
 *
 * Cards are built while the reply is streaming, but the engine is single-threaded:
 * clicking Apply mid-stream always fails with "the model is busy". So cards start
 * disarmed and only go live when the turn completes.
 */
export function armProposals(root) {
  for (const card of root.querySelectorAll(".proposal[data-waiting]")) {
    delete card.dataset.waiting;
    card.querySelector(".proposal__wait")?.remove();
    for (const btn of card.querySelectorAll(".proposal__actions .btn")) btn.disabled = false;
  }
}

const isBatchedType = (type) =>
  type === "transform_column" || type === "split_column" || type === "extract_table";

/** The button label a change deserves when it stands alone. */
const applyLabelFor = (action) =>
  action.type === "extract_table"
    ? "Extract records"
    : isBatchedType(action.type)
      ? `${action.columns ? "Split" : "Rewrite"} ${action.rows} rows`
      : "Apply";

/**
 * The body of one change inside the grouped card. Instant changes show their summary
 * (and preview grid); long-running ones carry their instruction block plus a progress
 * bar and sample area that come alive during the run.
 */
function buildActionSection(action) {
  const section = el("div", "proposal__section");
  section.appendChild(el("div", "proposal__sectionaddr mono",
    action.sheet ? `${action.sheet}!${action.address ?? ""}` : (action.address ?? "")));

  if (isBatchedType(action.type)) {
    const transform = buildTransformBody(action);
    section.appendChild(transform.wrap);
    return { section, transform };
  }

  const showsGrid = action.type === "write_formula" || action.type === "write_values";
  section.appendChild(showsGrid ? el("div", "proposal__summary", action.summary) : buildDetails(action));
  if (showsGrid) section.appendChild(buildGrid(action));
  return { section, transform: null };
}

/**
 * ONE card for everything a reply proposes — instant writes and long-running jobs
 * alike — behind ONE Apply, with ONE Undo.
 *
 * A reply that splits a column and cleans two others is one plan; asking the user to
 * approve it three times is noise. Changes apply in order (long jobs streaming their
 * progress into their own section, stoppable mid-run), roll back together when one
 * fails, and undo in reverse with one click. Every undo verifies the cells still hold
 * exactly what was written and refuses rather than wiping newer work.
 *
 * Returns { card, add }; `add` absorbs each action as it streams in.
 */
export function buildProposalGroup(handlers, { deferActions = false } = {}) {
  const card = el("div", "proposal");
  card.dataset.status = "pending";
  card.dataset.multi = "false";

  const head = el("div", "proposal__head");
  const label = el("span", "eyebrow proposal__label", "Proposed change");
  const addr = el("span", "proposal__addr", "");
  head.append(label, addr);
  card.appendChild(head);

  const sections = el("div", "proposal__sections");
  card.appendChild(sections);

  const actionsRow = el("div", "proposal__actions");
  const apply = el("button", "btn btn--primary", "Apply");
  const dismiss = el("button", "btn btn--quiet", "Dismiss");
  actionsRow.append(apply, dismiss);
  if (deferActions) {
    apply.disabled = true;
    dismiss.disabled = true;
    card.dataset.waiting = "true";
    actionsRow.appendChild(el("span", "proposal__wait", "Available when the reply finishes…"));
  }
  card.appendChild(actionsRow);

  /** [{action, transform}] in arrival order. */
  const entries = [];
  /** Set while a long job runs so Dismiss can act as Stop. */
  let cancelRunningJob = null;

  const add = (action) => {
    const built = buildActionSection(action);
    entries.push({ action, transform: built.transform });
    sections.appendChild(built.section);
    if (entries.length === 1) {
      addr.textContent = action.sheet ? `${action.sheet}!${action.address ?? ""}` : (action.address ?? "");
      apply.textContent = applyLabelFor(action);
    } else {
      card.dataset.multi = "true";
      label.textContent = `Proposed changes (${entries.length})`;
      apply.textContent = `Apply all ${entries.length}`;
      addr.textContent = "";
    }
  };

  const showError = (text) => {
    const error = el("div", "proposal__summary proposal__error", text);
    error.style.color = "var(--danger)";
    card.insertBefore(error, actionsRow);
  };

  /** Fill a section's transform UI from a finished job. */
  const settleTransformUi = (entry, { results, lossy, failed }) => {
    const t = entry.transform;
    t.progress.hidden = true;
    if (results?.length) {
      t.sample.hidden = false;
      t.sample.textContent = "";
      t.sample.appendChild(el("span", "eyebrow", "Result"));
      for (const value of results.slice(0, 4)) {
        t.sample.appendChild(el("div", "transform__row mono", formatSampleRow(value)));
      }
      if (results.length > 4) {
        t.sample.appendChild(el("div", "grid__more", `+ ${(results.length - 4).toLocaleString()} more rows`));
      }
    }
    if (lossy?.length) {
      const warn = el("div", "proposal__warning");
      warn.appendChild(el("span", "eyebrow", "Check these rows"));
      warn.appendChild(el("p", null,
        `${lossy.length} row${lossy.length === 1 ? "" : "s"} lost some of the original text — the columns chosen ` +
        `may not cover everything in the source. Undo and ask for more columns if so.`));
      entry.transform.wrap.appendChild(warn);
    }
    if (failed?.length) {
      const warn = el("div", "proposal__summary",
        `${failed.length} row${failed.length === 1 ? "" : "s"} could not be transformed and were left unchanged.`);
      warn.style.color = "var(--amber)";
      entry.transform.wrap.appendChild(warn);
    }
  };

  const finish = (undos) => {
    card.dataset.status = "applied";
    label.textContent = "Applied";
    actionsRow.remove();
    const receipt = el("div", "proposal__receipt");
    receipt.appendChild(el("span", null,
      entries.length === 1 ? "Change written to the sheet" : `${entries.length} changes written to the sheet`));
    if (undos.length) {
      const undoBtn = el("button", "link", "Undo");
      undoBtn.addEventListener("click", async () => {
        undoBtn.disabled = true;
        undoBtn.textContent = "Undoing…";
        let done = 0;
        try {
          for (const undo of [...undos].reverse()) {
            await undo();
            done += 1;
          }
          receipt.textContent = "";
          receipt.appendChild(el("span", null, entries.length === 1 ? "Change undone" : "All changes undone"));
          card.dataset.status = "pending";
          label.textContent = "Undone";
        } catch (err) {
          undoBtn.disabled = false;
          undoBtn.textContent = "Undo";
          receipt.querySelector("span").textContent = done
            ? `Undid ${done} of ${undos.length}, then stopped: ${err.message}`
            : `Could not undo: ${err.message}`;
        }
      });
      receipt.appendChild(undoBtn);
    }
    card.appendChild(receipt);
  };

  apply.addEventListener("click", async () => {
    card.querySelector(".proposal__error")?.remove();

    const blocked = handlers.canApply?.();
    if (typeof blocked === "string") {
      const notice = el("div", "proposal__summary proposal__error", blocked);
      notice.style.color = "var(--amber)";
      card.insertBefore(notice, actionsRow);
      return;
    }

    const applyLabel = apply.textContent;
    apply.disabled = true;
    dismiss.disabled = true;
    apply.textContent = "Applying…";

    const undos = [];
    let current = null;
    try {
      for (const entry of entries) {
        current = entry;
        const t = entry.transform;
        if (t) {
          // A long job reports progress and can be stopped; Dismiss becomes Stop.
          t.progress.hidden = false;
          dismiss.disabled = false;
          dismiss.textContent = "Stop";
          dismiss.dataset.mode = "stop";
        }

        const result = await handlers.onApply(entry.action, {
          onCancelHandle: (fn) => { cancelRunningJob = fn; },
          onPhase: (phase) => {
            if (!t) return;
            t.label.textContent =
              phase === "reading" ? "Reading the column…" : phase === "writing" ? "Writing results…" : "Starting…";
          },
          onProgress: ({ done, total, sample, records }) => {
            if (!t) return;
            t.bar.style.width = `${Math.max(2, (done / total) * 100).toFixed(1)}%`;
            t.label.textContent = `${done.toLocaleString()} of ${total.toLocaleString()} rows`
              + (records != null ? ` · ${records.toLocaleString()} record${records === 1 ? "" : "s"} found` : "");
            if (sample?.length) {
              t.sample.hidden = false;
              t.sample.textContent = "";
              t.sample.appendChild(el("span", "eyebrow", "First results"));
              for (const value of sample.slice(0, 4)) {
                t.sample.appendChild(el("div", "transform__row mono", formatSampleRow(value)));
              }
            }
          },
        });

        if (t) {
          settleTransformUi(entry, result ?? {});
          dismiss.disabled = true;
          dismiss.textContent = "Dismiss";
          delete dismiss.dataset.mode;
          cancelRunningJob = null;
        }
        if (result?.undo) undos.push(result.undo);
      }
      finish(undos);
    } catch (err) {
      // Roll back whatever landed, most recent first, so the sheet holds all of the
      // proposal or none of it.
      for (const undo of undos.reverse()) {
        try { await undo(); } catch { /* keep unwinding */ }
      }
      if (current?.transform) {
        current.transform.progress.hidden = true;
      }
      cancelRunningJob = null;
      apply.disabled = false;
      dismiss.disabled = false;
      dismiss.textContent = "Dismiss";
      delete dismiss.dataset.mode;
      apply.textContent = applyLabel;
      showError(
        `Could not apply${entries.length > 1 && current ? ` "${current.action.summary}"` : ""}: ${err.message}. ` +
        `${undos.length || entries.indexOf(current) > 0 ? "Everything that had been applied was rolled back." : "The sheet was left unchanged."}`,
      );
    }
  });

  dismiss.addEventListener("click", () => {
    if (dismiss.dataset.mode === "stop") {
      cancelRunningJob?.();
      dismiss.disabled = true;
      dismiss.textContent = "Stopping…";
      return;
    }
    for (const entry of entries) handlers.onDismiss?.(entry.action);
    card.dataset.status = "applied";
    label.textContent = "Dismissed";
    actionsRow.remove();
    const receipt = el("div", "proposal__receipt");
    receipt.appendChild(el("span", null, entries.length === 1 ? "Change dismissed" : "Changes dismissed"));
    card.appendChild(receipt);
  });

  return { card, add };
}

/**
 * Build the proposed-change card.
 * @param {object} action
 * @param {{onApply:Function,onDismiss:Function,canApply?:() => true|string}} handlers
 *   `canApply` is consulted at click time; a string return blocks the apply and is
 *   shown as the reason (the engine can only do one thing at a time).
 * @param {{deferActions?:boolean}} [opts]  start with the buttons disabled until
 *   `armProposals` runs — used while the reply is still streaming.
 */
export function buildProposal(action, handlers, { deferActions = false } = {}) {
  const card = el("div", "proposal");
  card.dataset.status = "pending";

  const head = el("div", "proposal__head");
  head.appendChild(el("span", "eyebrow proposal__label", "Proposed change"));
  const addr = el("span", "proposal__addr", action.sheet ? `${action.sheet}!${action.address ?? ""}` : (action.address ?? ""));
  head.appendChild(addr);
  card.appendChild(head);

  const isTransform = action.type === "transform_column" || action.type === "split_column" || action.type === "extract_table";
  const showsGrid = action.type === "write_formula" || action.type === "write_values";

  let transform = null;
  if (isTransform) {
    transform = buildTransformBody(action);
    card.appendChild(transform.wrap);
  } else {
    card.appendChild(showsGrid ? el("div", "proposal__summary", action.summary) : buildDetails(action));
    if (showsGrid) card.appendChild(buildGrid(action));
  }

  const applyLabel = action.type === "extract_table"
    ? "Extract records"
    : isTransform ? `${action.columns ? "Split" : "Rewrite"} ${action.rows} rows` : "Apply";

  const actions = el("div", "proposal__actions");
  const apply = el("button", "btn btn--primary", applyLabel);
  const dismiss = el("button", "btn btn--quiet", "Dismiss");
  actions.append(apply, dismiss);
  if (deferActions) {
    apply.disabled = true;
    dismiss.disabled = true;
    card.dataset.waiting = "true";
    actions.appendChild(el("span", "proposal__wait", "Available when the reply finishes…"));
  }
  card.appendChild(actions);

  /** Set while a long transform is running so Dismiss can act as Stop. */
  let cancelRunningJob = null;

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
    // Clear any failure from a previous attempt; leaving it up next to a success
    // receipt tells the user two contradictory things at once.
    card.querySelector(".proposal__error")?.remove();

    // The engine does one thing at a time. Refusing here, with the reason, beats
    // firing the request and surfacing "the model is busy" as a failure.
    const blocked = handlers.canApply?.();
    if (typeof blocked === "string") {
      const notice = el("div", "proposal__summary proposal__error", blocked);
      notice.style.color = "var(--amber)";
      card.insertBefore(notice, actions);
      return;
    }

    apply.disabled = true;
    dismiss.disabled = true;
    apply.textContent = isTransform ? "Working…" : "Applying…";

    // A column rewrite runs for minutes, so it reports progress and can be stopped.
    if (isTransform) {
      transform.progress.hidden = false;
      dismiss.disabled = false;
      dismiss.textContent = "Stop";
      dismiss.dataset.mode = "stop";
    }

    try {
      const { undo, failed, results, lossy } = await handlers.onApply(action, {
        onCancelHandle: (fn) => { cancelRunningJob = fn; },
        onPhase: (phase) => {
          if (!isTransform) return;
          transform.label.textContent =
            phase === "reading" ? "Reading the column…" : phase === "writing" ? "Writing results…" : "Starting…";
        },
        onProgress: ({ done, total, sample, records }) => {
          if (!isTransform) return;
          transform.bar.style.width = `${Math.max(2, (done / total) * 100).toFixed(1)}%`;
          transform.label.textContent = `${done.toLocaleString()} of ${total.toLocaleString()} rows`
            + (records != null ? ` · ${records.toLocaleString()} record${records === 1 ? "" : "s"} found` : "");
          if (sample?.length) {
            transform.sample.hidden = false;
            transform.sample.textContent = "";
            transform.sample.appendChild(el("span", "eyebrow", "First results"));
            for (const value of sample.slice(0, 4)) {
              transform.sample.appendChild(el("div", "transform__row mono", formatSampleRow(value)));
            }
          }
        },
      });

      if (isTransform) {
        transform.progress.hidden = true;
        delete dismiss.dataset.mode;
        if (results?.length) {
          transform.sample.hidden = false;
          transform.sample.textContent = "";
          transform.sample.appendChild(el("span", "eyebrow", "Result"));
          for (const value of results.slice(0, 4)) {
            transform.sample.appendChild(el("div", "transform__row mono", formatSampleRow(value)));
          }
          if (results.length > 4) {
            transform.sample.appendChild(el("div", "grid__more", `+ ${(results.length - 4).toLocaleString()} more rows`));
          }
        }
        if (lossy?.length) {
          const rows = lossy.slice(0, 5).map((i) => i + (action.rows ? 0 : 0));
          const warn = el("div", "proposal__warning");
          warn.appendChild(el("span", "eyebrow", "Check these rows"));
          warn.appendChild(el("p", null,
            `${lossy.length} row${lossy.length === 1 ? "" : "s"} lost some of the original text — the columns chosen ` +
            `may not cover everything in the source. Undo and ask for more columns if so.`));
          card.appendChild(warn);
        }
        if (failed?.length) {
          const warn = el("div", "proposal__summary", `${failed.length} row${failed.length === 1 ? "" : "s"} could not be transformed and were left unchanged.`);
          warn.style.color = "var(--amber)";
          card.appendChild(warn);
        }
      }

      finish("Applied", undo);
    } catch (err) {
      if (isTransform) {
        transform.progress.hidden = true;
        delete dismiss.dataset.mode;
        dismiss.textContent = "Dismiss";
      }
      apply.disabled = false;
      dismiss.disabled = false;
      apply.textContent = applyLabel;
      const error = card.querySelector(".proposal__error") ?? el("div", "proposal__summary proposal__error");
      error.textContent = `Could not apply: ${err.message}`;
      error.style.color = "var(--danger)";
      card.insertBefore(error, actions);
    }
  });

  dismiss.addEventListener("click", () => {
    if (dismiss.dataset.mode === "stop") {
      cancelRunningJob?.();
      dismiss.disabled = true;
      dismiss.textContent = "Stopping…";
      return;
    }
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
