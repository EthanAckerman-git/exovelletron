/** Saved conversation history. */

export function registerConversationRoutes(routes, { history, engine, json, readJsonBody }) {
  routes.set("GET /api/conversations", async (_req, res) => {
    json(res, 200, { conversations: await history.list() });
  });

  routes.set("GET /api/conversations/detail", async (_req, res, ctx) => {
    const id = ctx.url.searchParams.get("id");
    if (!id) return json(res, 400, { error: "id is required" });
    const conversation = await history.get(id);
    if (!conversation) return json(res, 404, { error: "No such conversation" });
    json(res, 200, { conversation });
  });

  routes.set("POST /api/conversations/save", async (req, res) => {
    const body = await readJsonBody(req);
    try {
      json(res, 200, { conversation: await history.save(body) });
    } catch (err) {
      json(res, 400, { error: err.message });
    }
  });

  routes.set("POST /api/conversations/delete", async (req, res) => {
    const { id } = await readJsonBody(req);
    if (!id) return json(res, 400, { error: "id is required" });
    await history.remove(id);
    json(res, 200, { ok: true });
  });

  /**
   * Re-prime the model with a saved conversation so follow-up questions still make
   * sense after reopening it. Without this the transcript would be visible but the
   * model would have no memory of any of it.
   */
  routes.set("POST /api/conversations/resume", async (req, res) => {
    const { id } = await readJsonBody(req);
    const conversation = await history.get(id);
    if (!conversation) return json(res, 404, { error: "No such conversation" });
    await engine.restoreConversation(conversation.messages);
    json(res, 200, { ok: true, messages: conversation.messages.length });
  });
}
