/**
 * Task pane entry point: wires Office, the local engine, and the transcript together.
 */
import * as api from "./api.js";
import { readSheetContext, describeSelection } from "./sheet/context.js";
import { applyAction, revealRange } from "./sheet/apply.js";
import { runTransform, runExtract } from "./sheet/transform.js";
import { el, buildTurn, buildProposal, armProposals, renderProse } from "./ui/render.js";
import { createHistoryPanel } from "./ui/history.js";

const dom = {};
let history = null;

const state = {
  ready: false,
  busy: false,
  cancelStream: null,
  sheetContext: null,
  hasTurns: false,
  /** Persisted transcript. Kept alongside the DOM so history survives a reopen. */
  conversationId: null,
  messages: [],
};

/**
 * Starting points, written as the sentence the user would actually send. Each one is a
 * complete request rather than a topic, so clicking it does something real immediately
 * instead of leaving a half-finished prompt to edit.
 */
const SUGGESTIONS = [
  {
    title: "Explain this data",
    hint: "What am I looking at?",
    prompt: "Explain what this data is and what each column holds.",
  },
  {
    title: "Split a messy column",
    hint: "One cell holding several fields",
    prompt: "This column has several fields crammed into each cell. Split it into separate, properly named columns.",
  },
  {
    title: "Clean up the values",
    hint: "Inconsistent names, dates, addresses",
    prompt: "The values in this column are inconsistently formatted. Standardise every row.",
  },
  {
    title: "Add a calculated column",
    hint: "A formula across every row",
    prompt: "Add a new column calculated from the existing ones, and tell me what you chose.",
  },
  {
    title: "Make it look good",
    hint: "Headers, number formats, colour",
    prompt: "Format this sheet nicely: bold the header row with a light fill, and apply sensible number formats.",
  },
  {
    title: "Find problems",
    hint: "Blanks, duplicates, odd values",
    prompt: "Look through this data and tell me about any blanks, duplicates, or values that look wrong.",
  },
];

/* ------------------------------------------------------------------ bootstrapping */

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    document.body.textContent = "Exovelletron runs inside Microsoft Excel.";
    return;
  }
  cacheDom();
  applyHostTheme();
  bindEvents();
  api.setSessionExpiredHandler(recoverSession);
  refreshSelection();
  watchStatus();
  restoreAfterReload();
});

/** Where the open conversation is parked across a token-recovery reload. */
const RESUME_KEY = "eal.resume-conversation";

/**
 * Recover from a restarted desktop app.
 *
 * The token in this page belongs to a server run that has ended, so every call now 401s.
 * Reloading re-fetches the pane with a fresh token. The open conversation is already
 * saved, so it is parked by id and reopened on the way back — the user sees a blink,
 * not an error they have to understand.
 */
let recovering = false;
function recoverSession() {
  if (recovering) return;
  recovering = true;
  try {
    if (state.conversationId) sessionStorage.setItem(RESUME_KEY, state.conversationId);
  } catch { /* private mode; the reload still fixes the token */ }
  setTimeout(() => location.reload(), 150);
}

async function restoreAfterReload() {
  let id = null;
  try {
    id = sessionStorage.getItem(RESUME_KEY);
    sessionStorage.removeItem(RESUME_KEY);
  } catch { /* nothing to restore */ }
  if (id) await loadConversation(id).catch(() => {});
}

function cacheDom() {
  for (const id of ["dot", "modelName", "thread", "input", "send", "selAddr", "selMeta", "settings", "newChat", "historyBtn"]) {
    dom[id] = document.getElementById(id);
  }
  history = createHistoryPanel({ onOpen: loadConversation });
  document.getElementById("app").appendChild(history.node);
}

/** Follow Excel's own theme when it exposes one; otherwise fall back to the OS. */
function applyHostTheme() {
  try {
    const theme = Office.context?.officeTheme;
    if (!theme?.bodyBackgroundColor) return;
    const hex = theme.bodyBackgroundColor.replace("#", "");
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    document.documentElement.dataset.theme = luminance < 0.5 ? "dark" : "light";
  } catch { /* fall back to prefers-color-scheme */ }
}

function bindEvents() {
  dom.input.addEventListener("input", autoGrow);
  dom.input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      submit();
    }
  });
  dom.send.addEventListener("click", () => (state.busy ? stop() : submit()));
  dom.newChat.addEventListener("click", newConversation);
  dom.historyBtn.addEventListener("click", () => history.toggle());
  dom.settings.addEventListener("click", openControlPanel);

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => refreshSelection(),
  );
}

function autoGrow() {
  dom.input.style.height = "auto";
  dom.input.style.height = `${Math.min(dom.input.scrollHeight, 132)}px`;
  dom.send.disabled = !state.busy && (!dom.input.value.trim() || !state.ready);
}

/* ------------------------------------------------------------------------ status */

function setEngineState(status) {
  const engineState = status?.engine?.state ?? "idle";
  dom.dot.dataset.state = engineState;
  state.ready = engineState === "ready";

  if (status?.model) {
    dom.modelName.textContent = status.model.name;
    dom.dot.title = `${status.model.name} — ${engineState}`;
  } else {
    dom.modelName.textContent = engineState === "loading" ? "Loading model…" : "No model loaded";
  }

  dom.input.disabled = false;
  dom.input.placeholder = state.ready
    ? "Ask about this sheet…"
    : engineState === "loading"
      ? "Starting the model…"
      : "Open Exovelletron to set up a model";
  autoGrow();

  if (!state.ready && !state.hasTurns) showSetupNotice(status);
  else if (state.ready && !state.hasTurns) showEmptyState();
}

async function watchStatus() {
  try {
    setEngineState(await api.getStatus());
  } catch {
    dom.dot.dataset.state = "error";
    dom.modelName.textContent = "Local engine not running";
    showConnectionError();
    return;
  }

  api.subscribeEvents({
    engine: async () => {
      try { setEngineState(await api.getStatus()); } catch { /* transient */ }
    },
    progress: (p) => updateDownloadProgress(p),
    installed: async () => {
      try { setEngineState(await api.getStatus()); } catch { /* transient */ }
    },
  });
}

function showConnectionError() {
  dom.thread.textContent = "";
  const notice = el("div", "notice");
  notice.appendChild(el("p", "notice__title", "The local engine isn't running"));
  notice.appendChild(el("p", "notice__body", "Open the Exovelletron app to start it, then reopen this pane."));
  dom.thread.appendChild(notice);
}

function showSetupNotice(status) {
  dom.thread.textContent = "";
  const notice = el("div", "notice");
  const downloading = status?.downloading;

  if (downloading) {
    notice.appendChild(el("p", "notice__title", "Downloading the model"));
    notice.appendChild(el("p", "notice__body", "This happens once. You can leave this open."));
    const bar = el("div", "progress");
    bar.appendChild(el("i"));
    bar.id = "dlbar";
    notice.appendChild(bar);
    notice.appendChild(el("p", "eyebrow", "Starting…")).id = "dltext";
  } else if (status?.engine?.state === "loading") {
    notice.appendChild(el("p", "notice__title", "Starting the model"));
    notice.appendChild(el("p", "notice__body", "This takes a couple of seconds."));
  } else {
    notice.appendChild(el("p", "notice__title", "No model yet"));
    notice.appendChild(el("p", "notice__body", "Open the Exovelletron app to download one. It's a one-time step."));
    const btn = el("button", "btn btn--primary", "Open Exovelletron");
    btn.addEventListener("click", openControlPanel);
    notice.appendChild(btn);
  }
  dom.thread.appendChild(notice);
}

function updateDownloadProgress(p) {
  const bar = document.getElementById("dlbar");
  const text = document.getElementById("dltext");
  if (bar) bar.firstChild.style.width = `${Math.max(2, p.percent).toFixed(1)}%`;
  if (text) {
    const mb = (n) => `${(n / 1e9).toFixed(2)} GB`;
    text.textContent = p.phase === "verifying"
      ? "Verifying download…"
      : `${p.percent.toFixed(0)}% · ${mb(p.received)} of ${mb(p.total)}`;
  }
}

function openControlPanel() {
  // The pane cannot launch a native app directly; tell the user plainly.
  const existing = document.getElementById("openhint");
  if (existing) return;
  const hint = el("div", "notice");
  hint.id = "openhint";
  hint.appendChild(el("p", "notice__title", "Open the app from your Dock"));
  hint.appendChild(el("p", "notice__body", "Exovelletron runs as a menu-bar app. Click its icon in the menu bar, or open it from Applications."));
  dom.thread.appendChild(hint);
  dom.thread.scrollTop = dom.thread.scrollHeight;
}

/* --------------------------------------------------------------------- selection */

async function refreshSelection() {
  try {
    state.sheetContext = await readSheetContext();
    const { address, meta } = describeSelection(state.sheetContext);
    dom.selAddr.textContent = address;
    dom.selMeta.textContent = meta;
  } catch {
    dom.selAddr.textContent = "No workbook open";
    dom.selMeta.textContent = "";
  }
}

/* ------------------------------------------------------------------- empty state */

function showEmptyState() {
  dom.thread.textContent = "";
  const empty = el("div", "empty");
  empty.appendChild(el("p", "empty__title", "Ask about this sheet"));
  empty.appendChild(el("p", "empty__sub", "Pick one, or just describe what you want. Runs on this Mac — nothing is uploaded."));

  const list = el("div", "suggestions");
  for (const s of SUGGESTIONS) {
    const btn = el("button", "suggestion");
    btn.appendChild(document.createTextNode(s.title));
    btn.appendChild(el("span", null, s.hint));
    btn.addEventListener("click", () => {
      dom.input.value = s.prompt ?? s.title;
      autoGrow();
      submit();
    });
    list.appendChild(btn);
  }
  empty.appendChild(list);
  dom.thread.appendChild(empty);
}

/* -------------------------------------------------------------------- conversing */

/* ----------------------------------------------------------------- persistence */

/** Append to the transcript we persist. Returns the stored record so it can be updated. */
function recordMessage(role, text, extra = {}) {
  const message = { id: `${role}-${Date.now()}-${state.messages.length}`, role, text, at: Date.now(), ...extra };
  state.messages.push(message);
  return message;
}

let saveTimer = null;
/**
 * Save the transcript, coalescing the bursts that arrive while a reply streams.
 * Failures are deliberately quiet: losing history is a nuisance, but an error banner
 * over a working conversation is worse.
 */
function persistConversation() {
  if (!state.messages.length) return;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    try {
      const { conversation } = await api.saveConversation({
        id: state.conversationId ?? undefined,
        messages: state.messages,
      });
      state.conversationId = conversation.id;
    } catch { /* history is best-effort */ }
  }, 400);
}

/** Re-render a saved conversation and put the model back in its context. */
async function loadConversation(id) {
  if (state.busy) stop();
  let conversation;
  try {
    ({ conversation } = await api.getConversation(id));
  } catch (err) {
    return toastError(err.message);
  }

  state.conversationId = conversation.id;
  state.messages = conversation.messages;
  state.hasTurns = true;
  dom.thread.textContent = "";

  for (const message of conversation.messages) {
    if (message.role === "error") {
      const turn = buildTurn("error");
      turn.body.textContent = message.text;
      dom.thread.appendChild(turn.turn);
      continue;
    }
    const turn = buildTurn(message.role);
    if (message.role === "assistant") renderProse(turn.body, message.text);
    else turn.body.textContent = message.text;
    dom.thread.appendChild(turn.turn);

    // Actions are shown as a record of what happened. They are not re-appliable: the
    // workbook has moved on, and the undo snapshot died with the session that made it.
    for (const action of message.actions ?? []) {
      const past = el("div", "proposal");
      past.dataset.status = "applied";
      const head = el("div", "proposal__head");
      head.appendChild(el("span", "eyebrow proposal__label", action.status === "applied" ? "Applied" : "Not applied"));
      head.appendChild(el("span", "proposal__addr", action.address ?? ""));
      past.appendChild(head);
      past.appendChild(el("div", "proposal__summary", action.summary ?? ""));
      turn.turn.appendChild(past);
    }
  }

  try { await api.resumeConversation(id); } catch { /* the transcript still reads fine */ }
  dom.thread.scrollTop = dom.thread.scrollHeight;
}

function toastError(message) {
  const turn = buildTurn("error");
  turn.body.textContent = message;
  dom.thread.appendChild(turn.turn);
  dom.thread.scrollTop = dom.thread.scrollHeight;
}

async function newConversation() {
  if (state.busy) stop();
  try { await api.resetConversation(); } catch { /* non-fatal */ }
  state.hasTurns = false;
  state.conversationId = null;
  state.messages = [];
  history?.refresh().catch(() => {});
  showEmptyState();
}

function setBusy(busy) {
  state.busy = busy;
  dom.send.dataset.mode = busy ? "stop" : "send";
  dom.send.setAttribute("aria-label", busy ? "Stop generating" : "Send message");
  dom.send.disabled = busy ? false : !dom.input.value.trim() || !state.ready;
}

function stop() {
  state.cancelStream?.();
  api.abortChat().catch(() => {});
  setBusy(false);
}

async function submit() {
  const message = dom.input.value.trim();
  if (!message || state.busy || !state.ready) return;

  if (!state.hasTurns) {
    dom.thread.textContent = "";
    state.hasTurns = true;
  }

  dom.input.value = "";
  autoGrow();

  const userTurn = buildTurn("user");
  userTurn.body.textContent = message;
  dom.thread.appendChild(userTurn.turn);
  recordMessage("user", message);

  const answer = buildTurn("assistant");
  const pending = el("div", "thinking");
  pending.append(el("i"), el("i"), el("i"));
  answer.body.appendChild(pending);
  dom.thread.appendChild(answer.turn);
  scrollToEnd();

  setBusy(true);
  await refreshSelection();

  let text = "";
  let firstToken = true;
  const caret = el("span", "caret");

  state.cancelStream = api.streamChat(
    { message, sheetContext: state.sheetContext },
    {
      onToken: (chunk) => {
        if (firstToken) {
          answer.body.textContent = "";
          firstToken = false;
        }
        text += chunk;
        renderProse(answer.body, text);
        answer.body.lastElementChild?.appendChild(caret);
        scrollToEnd();
      },
      onAction: (action) => {
        caret.remove();
        if (firstToken) {
          answer.body.textContent = "";
          firstToken = false;
        }
        // The card appears while the reply is still streaming, but its buttons stay
        // disabled until the turn ends — the engine can only do one thing at a time,
        // so an early click would just be refused as busy.
        answer.turn.appendChild(buildProposal(
          action,
          { onApply: applyProposal, canApply: whyApplyIsBlocked },
          { deferActions: true },
        ));
        scrollToEnd();
      },
      onDone: ({ text: finalText, stats, actions }) => {
        caret.remove();
        armProposals(answer.turn);
        if (finalText) renderProse(answer.body, finalText);
        else if (firstToken) answer.body.remove();
        recordMessage("assistant", finalText || text, {
          actions: (actions ?? []).map((a) => ({ ...a, status: "pending" })),
          stats,
        });
        persistConversation();
        history?.refresh().catch(() => {});
        if (stats?.tokensPerSecond) {
          const meta = el("div", "composer__hint");
          meta.appendChild(el("span", "mono", `${stats.tokensPerSecond} tok/s · ${stats.seconds}s`));
          answer.turn.appendChild(meta);
        }
        setBusy(false);
        scrollToEnd();
      },
      onError: (failure) => {
        caret.remove();
        pending.remove();
        // The engine is free again, so any proposals that did arrive are usable.
        armProposals(answer.turn);
        if (firstToken) answer.turn.remove();
        const errorTurn = buildTurn("error");
        errorTurn.body.textContent = failure;
        dom.thread.appendChild(errorTurn.turn);
        recordMessage("error", failure);
        persistConversation();
        setBusy(false);
        scrollToEnd();
      },
    },
  );
}

/**
 * Carry out an approved action.
 *
 * A column rewrite goes down a different path from the single-shot edits: it streams
 * every row through the model, so it reports progress and can be stopped.
 */
/** Actions that stream every row through the model rather than writing in one shot. */
const BATCHED_ACTIONS = new Set(["transform_column", "split_column"]);
/** Reads a whole range through the model and writes one row per record it finds. */
const EXTRACT_ACTION = "extract_table";

/**
 * Why an Apply click cannot run right now, or true when it can.
 * Old cards stay clickable across turns, so this is checked at click time too.
 */
function whyApplyIsBlocked() {
  if (state.busy) return "Wait for the current reply to finish, then apply.";
  return true;
}

async function applyProposal(action, handlers) {
  const result = action.type === EXTRACT_ACTION
    ? await runExtract(action, handlers)
    : BATCHED_ACTIONS.has(action.type)
      ? await runTransform(action, handlers)
      : await applyAction(action);

  revealRange(action).catch(() => {});
  refreshSelection();
  persistConversation();
  return result;
}

function scrollToEnd() {
  // Only follow the stream when the user has not scrolled up to read something.
  const nearBottom = dom.thread.scrollHeight - dom.thread.scrollTop - dom.thread.clientHeight < 120;
  if (nearBottom) dom.thread.scrollTop = dom.thread.scrollHeight;
}
