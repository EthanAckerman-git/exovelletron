/**
 * Task pane entry point: wires Office, the local engine, and the transcript together.
 */
import * as api from "./api.js";
import { readSheetContext, describeSelection } from "./sheet/context.js";
import { applyAction, revealRange } from "./sheet/apply.js";
import { el, buildTurn, buildProposal, renderProse } from "./ui/render.js";

const dom = {};
const state = {
  ready: false,
  busy: false,
  cancelStream: null,
  sheetContext: null,
  hasTurns: false,
};

const SUGGESTIONS = [
  { title: "Explain this selection", hint: "What am I looking at?" },
  { title: "Write a formula", hint: "Add a column that…" },
  { title: "Find problems in this data", hint: "Blanks, duplicates, odd values" },
];

/* ------------------------------------------------------------------ bootstrapping */

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    document.body.textContent = "Excel AI Local runs inside Microsoft Excel.";
    return;
  }
  cacheDom();
  applyHostTheme();
  bindEvents();
  refreshSelection();
  watchStatus();
});

function cacheDom() {
  for (const id of ["dot", "modelName", "thread", "input", "send", "selAddr", "selMeta", "settings", "newChat"]) {
    dom[id] = document.getElementById(id);
  }
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
      : "Open Excel AI Local to set up a model";
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
  notice.appendChild(el("p", "notice__body", "Open the Excel AI Local app to start it, then reopen this pane."));
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
    notice.appendChild(el("p", "notice__body", "Open the Excel AI Local app to download one. It's a one-time step."));
    const btn = el("button", "btn btn--primary", "Open Excel AI Local");
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
  hint.appendChild(el("p", "notice__body", "Excel AI Local runs as a menu-bar app. Click its icon in the menu bar, or open it from Applications."));
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
  empty.appendChild(el("p", "empty__sub", "Runs on this Mac. Nothing is uploaded."));

  const list = el("div", "suggestions");
  for (const s of SUGGESTIONS) {
    const btn = el("button", "suggestion");
    btn.appendChild(document.createTextNode(s.title));
    btn.appendChild(el("span", null, s.hint));
    btn.addEventListener("click", () => {
      dom.input.value = s.title;
      autoGrow();
      submit();
    });
    list.appendChild(btn);
  }
  empty.appendChild(list);
  dom.thread.appendChild(empty);
}

/* -------------------------------------------------------------------- conversing */

async function newConversation() {
  if (state.busy) stop();
  try { await api.resetConversation(); } catch { /* non-fatal */ }
  state.hasTurns = false;
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
        answer.turn.appendChild(
          buildProposal(action, {
            onApply: async (a) => {
              const result = await applyAction(a);
              revealRange(a).catch(() => {});
              refreshSelection();
              return result;
            },
          }),
        );
        scrollToEnd();
      },
      onDone: ({ text: finalText, stats }) => {
        caret.remove();
        if (finalText) renderProse(answer.body, finalText);
        else if (firstToken) answer.body.remove();
        if (stats?.tokensPerSecond) {
          const meta = el("div", "composer__hint");
          meta.appendChild(el("span", "mono", `${stats.tokensPerSecond} tok/s · ${stats.seconds}s`));
          answer.turn.appendChild(meta);
        }
        setBusy(false);
        scrollToEnd();
      },
      onError: (message) => {
        caret.remove();
        pending.remove();
        if (firstToken) answer.turn.remove();
        const errorTurn = buildTurn("error");
        errorTurn.body.textContent = message;
        dom.thread.appendChild(errorTurn.turn);
        setBusy(false);
        scrollToEnd();
      },
    },
  );
}

function scrollToEnd() {
  // Only follow the stream when the user has not scrolled up to read something.
  const nearBottom = dom.thread.scrollHeight - dom.thread.scrollTop - dom.thread.clientHeight < 120;
  if (nearBottom) dom.thread.scrollTop = dom.thread.scrollHeight;
}
