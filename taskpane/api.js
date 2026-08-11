/**
 * Client for the local server. Same-origin, so the session token in the page is the
 * only credential needed.
 */

const TOKEN = window.__EAL_TOKEN__;
const HEADERS = { "Content-Type": "application/json", "x-eal-token": TOKEN };

/**
 * The session token is minted per server run and baked into this page when it was
 * served. If the desktop app restarts, every call starts failing with 401 and the pane
 * looks broken for no reason the user can see or act on. Registering a handler lets the
 * pane reload itself and pick up a fresh token instead.
 */
let onSessionExpired = null;
export const setSessionExpiredHandler = (fn) => { onSessionExpired = fn; };

/** True when the response means "your token is from a previous run of the app". */
const isStaleSession = (status) => status === 401;

async function request(path, options = {}) {
  const res = await fetch(path, { ...options, headers: { ...HEADERS, ...(options.headers || {}) } });
  if (!res.ok) {
    if (isStaleSession(res.status)) onSessionExpired?.();
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
export const setSettings = (patch) => request("/api/settings", { method: "POST", body: JSON.stringify(patch) });

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
        if (isStaleSession(res.status)) onSessionExpired?.();
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
          else if (event === "search") handlers.onSearch?.(payload);
          else if (event === "fetch") handlers.onFetch?.(payload);
          else if (event === "read_request") handlers.onReadRequest?.(payload);
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

/** Answer a mid-turn workbook read the model requested over the chat stream. */
export const postReadResult = (body) => request("/api/chat/read", { method: "POST", body: JSON.stringify(body) });

/** Conversation history. */
export const listConversations = () => request("/api/conversations");
export const getConversation = (id) => request(`/api/conversations/detail?id=${encodeURIComponent(id)}`);
export const saveConversation = (body) => request("/api/conversations/save", { method: "POST", body: JSON.stringify(body) });
export const deleteConversation = (id) => request("/api/conversations/delete", { method: "POST", body: JSON.stringify({ id }) });
export const resumeConversation = (id) => request("/api/conversations/resume", { method: "POST", body: JSON.stringify({ id }) });

/**
 * Stream a long model job (column transform, record extraction) as SSE.
 * @returns {(body:object, handlers:object) => () => void} the returned fn aborts
 */
function streamJob(path, failureLabel) {
  return (body, handlers) => {
    const controller = new AbortController();

    (async () => {
      try {
        const res = await fetch(path, {
          method: "POST",
          headers: HEADERS,
          body: JSON.stringify(body),
          signal: controller.signal,
        });
        if (!res.ok) {
          if (isStaleSession(res.status)) onSessionExpired?.();
          let message = `Request failed (${res.status})`;
          try { message = (await res.json()).error || message; } catch { /* keep status */ }
          handlers.onError?.(message);
          return;
        }
        await readSse(res, {
          progress: (p) => handlers.onProgress?.(p),
          done: (p) => handlers.onDone?.(p),
          error: (p) => handlers.onError?.(p.message),
        });
      } catch (err) {
        if (err.name !== "AbortError") handlers.onError?.(err.message || failureLabel);
      }
    })();

    return () => controller.abort();
  };
}

/** Stream a whole-column transformation. Returns an abort fn. */
export const streamTransform = streamJob("/api/transform", "Transform failed");
/** Stream a whole-range record extraction. Returns an abort fn. */
export const streamExtract = streamJob("/api/extract", "Extraction failed");

/**
 * Read an SSE body, dispatching each frame to a handler by event name.
 * Frames are separated by a blank line; whatever trails the last one is a partial frame
 * and has to be carried into the next read.
 */
async function readSse(res, handlers) {
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
