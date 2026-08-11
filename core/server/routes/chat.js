/**
 * Chat endpoint, streamed over Server-Sent Events.
 *
 * SSE rather than a websocket: it is one-directional, survives the Office webview
 * cleanly, and needs no extra protocol handling.
 */

/** Serialise one SSE frame. Newlines in data must be split across `data:` lines. */
export function sseFrame(event, data) {
  const payload = JSON.stringify(data);
  return `event: ${event}\ndata: ${payload}\n\n`;
}

/** How long the pane gets to answer a workbook read before the model moves on. */
export const READ_TIMEOUT_MS = 20_000;

export function registerChatRoutes(routes, { engine, json, readJsonBody }) {
  /**
   * Workbook reads in flight, id → {resolve, timer}. SSE is one-way, so a mid-turn
   * read is a round trip in two halves: the engine parks on a promise while a
   * `read_request` frame rides the stream out, and the pane POSTs the cells back to
   * resolve it. Only one generation runs at a time, so one map serves the route.
   */
  const pendingReads = new Map();

  const settleRead = (id, result) => {
    const entry = pendingReads.get(id);
    if (!entry) return false;
    pendingReads.delete(id);
    clearTimeout(entry.timer);
    entry.resolve(result);
    return true;
  };

  const flushReads = (reason) => {
    for (const id of [...pendingReads.keys()]) settleRead(id, { ok: false, error: reason });
  };

  routes.set("POST /api/chat", async (req, res) => {
    let body;
    try {
      body = await readJsonBody(req);
    } catch (err) {
      return json(res, 400, { error: err.message });
    }

    const message = typeof body.message === "string" ? body.message.trim() : "";
    if (!message) return json(res, 400, { error: "message is required" });
    if (message.length > 20000) return json(res, 400, { error: "message is too long" });

    const status = engine.status();
    if (status.state === "loading") return json(res, 503, { error: "The model is still loading." });
    if (status.state !== "ready") return json(res, 503, { error: status.error || "No model is loaded." });
    if (status.busy) return json(res, 409, { error: "Already answering another message." });

    res.writeHead(200, {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-store",
      Connection: "keep-alive",
      "X-Accel-Buffering": "no",
    });
    res.write(sseFrame("start", { at: Date.now() }));

    // If the pane goes away mid-generation, stop burning GPU on an answer nobody
    // wants — and unpark any read the engine is waiting on, or the abort would have
    // to wait out the read timeout first. The listener sits on the RESPONSE: on
    // current Node an IncomingMessage's `close` fires when its body has been read,
    // which is the start of every request here, not a disconnect. `res` closing with
    // an unfinished body is the actual "the pane went away" signal.
    const onClose = () => {
      if (res.writableEnded) return;
      engine.abort();
      flushReads("the reply was cancelled");
    };
    res.on("close", onClose);

    const onAction = (action) => res.write(sseFrame("action", action));
    engine.on("action", onAction);
    // Surfaced so the pane can show exactly when — and for what — the model searched
    // or opened a page.
    const onSearch = (payload) => res.write(sseFrame("search", payload));
    engine.on("search", onSearch);
    const onFetch = (payload) => res.write(sseFrame("fetch", payload));
    engine.on("fetch", onFetch);

    const readSheet = (request) =>
      new Promise((resolve) => {
        const id = crypto.randomUUID();
        const timer = setTimeout(
          () => settleRead(id, { ok: false, error: "the workbook did not answer in time" }),
          READ_TIMEOUT_MS,
        );
        pendingReads.set(id, { resolve, timer });
        res.write(sseFrame("read_request", { id, sheet: request.sheet, address: request.address }));
      });

    try {
      const result = await engine.chat({
        message,
        sheetContext: body.sheetContext ?? null,
        onToken: (text) => res.write(sseFrame("token", { text })),
        readSheet,
      });
      res.write(sseFrame("done", { text: result.text, actions: result.actions, stats: result.stats }));
    } catch (err) {
      res.write(sseFrame("error", { message: err.message || "Generation failed" }));
    } finally {
      flushReads("the reply was cancelled");
      engine.off("action", onAction);
      engine.off("search", onSearch);
      engine.off("fetch", onFetch);
      res.off("close", onClose);
      res.end();
    }
  });

  // The second half of a workbook read: the pane answering a `read_request` frame.
  routes.set("POST /api/chat/read", async (req, res) => {
    let body;
    try {
      body = await readJsonBody(req);
    } catch (err) {
      return json(res, 400, { error: err.message });
    }

    const id = typeof body.id === "string" ? body.id : "";
    const result = body.ok
      ? {
          ok: true,
          sheet: typeof body.sheet === "string" ? body.sheet : null,
          address: typeof body.address === "string" ? body.address : "",
          startRow: Number.isInteger(body.startRow) ? body.startRow : 1,
          columnLetters: Array.isArray(body.columnLetters) ? body.columnLetters : [],
          rows: Array.isArray(body.rows) ? body.rows : [],
        }
      : { ok: false, error: typeof body.error === "string" ? body.error : "the read failed" };

    if (!settleRead(id, result)) {
      return json(res, 404, { error: "No read with that id is waiting." });
    }
    json(res, 200, { ok: true });
  });

  routes.set("POST /api/chat/abort", async (_req, res) => {
    engine.abort();
    json(res, 200, { ok: true });
  });
}
