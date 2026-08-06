/**
 * Client for the local server. Same-origin, so the session token in the page is the
 * only credential needed.
 */

const TOKEN = window.__EAL_TOKEN__;
const HEADERS = { "Content-Type": "application/json", "x-eal-token": TOKEN };

async function request(path, options = {}) {
  const res = await fetch(path, { ...options, headers: { ...HEADERS, ...(options.headers || {}) } });
  if (!res.ok) {
    let message = `Request failed (${res.status})`;
    try {
      message = (await res.json()).error || message;
    } catch { /* keep the status message */ }
    throw new Error(message);
  }
  return res.status === 204 ? null : res.json();
}

export const getStatus = () => request("/api/status");
export const listModels = () => request("/api/models");
export const selectModel = (id) => request("/api/models/select", { method: "POST", body: JSON.stringify({ id }) });
export const downloadModel = (id) => request("/api/models/download", { method: "POST", body: JSON.stringify({ id }) });
export const resetConversation = () => request("/api/engine/reset", { method: "POST" });
export const abortChat = () => request("/api/chat/abort", { method: "POST" });

/**
 * Stream a chat turn.
 *
 * Parses SSE by hand rather than using EventSource, because EventSource cannot issue a
 * POST or carry the session token header.
 *
 * @param {object} body            { message, sheetContext }
 * @param {object} handlers        { onToken, onAction, onDone, onError }
 * @returns {() => void}           call to abort the stream
 */
export function streamChat(body, handlers) {
  const controller = new AbortController();

  (async () => {
    try {
      const res = await fetch("/api/chat", {
        method: "POST",
        headers: HEADERS,
        body: JSON.stringify(body),
        signal: controller.signal,
      });

      if (!res.ok) {
        let message = `Request failed (${res.status})`;
        try { message = (await res.json()).error || message; } catch { /* keep status */ }
        handlers.onError?.(message);
        return;
      }

      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let buffer = "";

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });

        // Frames are separated by a blank line; anything after the last one is partial.
        const frames = buffer.split("\n\n");
        buffer = frames.pop() ?? "";

        for (const frame of frames) {
          if (!frame.trim() || frame.startsWith(":")) continue;
          let event = "message";
          const dataLines = [];
          for (const line of frame.split("\n")) {
            if (line.startsWith("event:")) event = line.slice(6).trim();
            else if (line.startsWith("data:")) dataLines.push(line.slice(5).trimStart());
          }
          if (!dataLines.length) continue;

          let payload;
          try { payload = JSON.parse(dataLines.join("\n")); } catch { continue; }

          if (event === "token") handlers.onToken?.(payload.text);
          else if (event === "action") handlers.onAction?.(payload);
          else if (event === "done") handlers.onDone?.(payload);
          else if (event === "error") handlers.onError?.(payload.message);
        }
      }
    } catch (err) {
      if (err.name !== "AbortError") handlers.onError?.(err.message || "Connection to the local engine was lost.");
    }
  })();

  return () => controller.abort();
}

/** Subscribe to model-download and engine-state events. Returns an unsubscribe fn. */
export function subscribeEvents(handlers) {
  const controller = new AbortController();

  (async () => {
    try {
      const res = await fetch("/api/models/events", { headers: HEADERS, signal: controller.signal });
      if (!res.ok) return;
      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let buffer = "";
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const frames = buffer.split("\n\n");
        buffer = frames.pop() ?? "";
        for (const frame of frames) {
          if (!frame.trim() || frame.startsWith(":")) continue;
          let event = "message";
          const dataLines = [];
          for (const line of frame.split("\n")) {
            if (line.startsWith("event:")) event = line.slice(6).trim();
            else if (line.startsWith("data:")) dataLines.push(line.slice(5).trimStart());
          }
          if (!dataLines.length) continue;
          try { handlers[event]?.(JSON.parse(dataLines.join("\n"))); } catch { /* skip bad frame */ }
        }
      }
    } catch { /* stream closed */ }
  })();

  return () => controller.abort();
}
