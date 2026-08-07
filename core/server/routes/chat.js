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

export function registerChatRoutes(routes, { engine, json, readJsonBody }) {
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

    // If the pane goes away mid-generation, stop burning GPU on an answer nobody wants.
    const onClose = () => engine.abort();
    req.on("close", onClose);

    const onAction = (action) => res.write(sseFrame("action", action));
    engine.on("action", onAction);
    // Surfaced so the pane can show exactly when — and for what — the model searched.
    const onSearch = (payload) => res.write(sseFrame("search", payload));
    engine.on("search", onSearch);

    try {
      const result = await engine.chat({
        message,
        sheetContext: body.sheetContext ?? null,
        onToken: (text) => res.write(sseFrame("token", { text })),
      });
      res.write(sseFrame("done", { text: result.text, actions: result.actions, stats: result.stats }));
    } catch (err) {
      res.write(sseFrame("error", { message: err.message || "Generation failed" }));
    } finally {
      engine.off("action", onAction);
      engine.off("search", onSearch);
      req.off("close", onClose);
      res.end();
    }
  });

  routes.set("POST /api/chat/abort", async (_req, res) => {
    engine.abort();
    json(res, 200, { ok: true });
  });
}
