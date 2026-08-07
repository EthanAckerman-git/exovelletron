/**
 * The saved-conversations panel.
 *
 * Slides over the transcript rather than replacing it, so closing it returns you exactly
 * where you were.
 */
import { el } from "./render.js";
import * as api from "../api.js";

/** "3 minutes ago" is more useful here than a timestamp nobody reads. */
export function relativeTime(then, now = Date.now()) {
  const seconds = Math.max(0, Math.round((now - then) / 1000));
  if (seconds < 60) return "just now";
  const minutes = Math.round(seconds / 60);
  if (minutes < 60) return `${minutes} minute${minutes === 1 ? "" : "s"} ago`;
  const hours = Math.round(minutes / 60);
  if (hours < 24) return `${hours} hour${hours === 1 ? "" : "s"} ago`;
  const days = Math.round(hours / 24);
  if (days < 7) return `${days} day${days === 1 ? "" : "s"} ago`;
  return new Date(then).toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

/**
 * @param {object} handlers
 * @param {(id:string) => Promise<void>} handlers.onOpen
 * @param {() => void} handlers.onNew
 */
export function createHistoryPanel(handlers) {
  const panel = el("aside", "history");
  panel.hidden = true;

  const head = el("div", "history__head");
  head.appendChild(el("span", "eyebrow", "Conversations"));
  const close = el("button", "status__btn", "");
  close.setAttribute("aria-label", "Close history");
  close.innerHTML = '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"><path d="M4.5 4.5l7 7M11.5 4.5l-7 7"/></svg>';
  head.appendChild(close);
  panel.appendChild(head);

  const list = el("div", "history__list");
  panel.appendChild(list);

  const hide = () => { panel.hidden = true; };
  close.addEventListener("click", hide);

  async function render() {
    list.textContent = "";
    let conversations = [];
    try {
      ({ conversations } = await api.listConversations());
    } catch {
      list.appendChild(el("p", "history__empty", "History is unavailable right now."));
      return;
    }

    if (!conversations.length) {
      list.appendChild(el("p", "history__empty", "Conversations you have here are saved automatically."));
      return;
    }

    for (const conversation of conversations) {
      const row = el("div", "history__item");

      const open = el("button", "history__open");
      open.appendChild(el("span", "history__title", conversation.title));
      const meta = el("span", "history__meta mono");
      meta.textContent = `${relativeTime(conversation.updatedAt)} · ${conversation.messageCount} message${conversation.messageCount === 1 ? "" : "s"}`;
      open.appendChild(meta);
      open.addEventListener("click", async () => {
        open.disabled = true;
        await handlers.onOpen(conversation.id);
        hide();
      });

      const remove = el("button", "history__delete", "");
      remove.setAttribute("aria-label", `Delete ${conversation.title}`);
      remove.innerHTML = '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"><path d="M3.5 4.5h9M6.5 4.5v-1h3v1M5 4.5l.5 8h5l.5-8"/></svg>';
      remove.addEventListener("click", async (event) => {
        event.stopPropagation();
        remove.disabled = true;
        try {
          await api.deleteConversation(conversation.id);
          row.remove();
          if (!list.querySelector(".history__item")) await render();
        } catch {
          remove.disabled = false;
        }
      });

      row.append(open, remove);
      list.appendChild(row);
    }
  }

  return {
    node: panel,
    hide,
    async toggle() {
      if (!panel.hidden) return hide();
      await render();
      panel.hidden = false;
    },
    refresh: render,
  };
}
