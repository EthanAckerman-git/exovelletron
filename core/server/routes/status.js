/** Health and capability reporting for the task pane. */
import os from "node:os";
import { DEFAULT_MODEL_ID, getModel } from "../../models/catalog.js";

export function registerStatusRoutes(routes, { engine, models, appInfo, settings, json, readJsonBody }) {
  routes.set("GET /api/status", async (_req, res) => {
    const engineStatus = engine.status();
    const active = engineStatus.modelId ? getModel(engineStatus.modelId) : null;
    json(res, 200, {
      ok: true,
      version: appInfo.version ?? "1.0.0",
      engine: engineStatus,
      model: active ? { id: active.id, name: active.name, params: active.params, contextTokens: engineStatus.contextTokens } : null,
      defaultModelId: DEFAULT_MODEL_ID,
      machine: { totalRamGb: Math.round(os.totalmem() / 1024 ** 3), arch: os.arch() },
      downloading: models.activeDownloadId,
      webSearch: settings ? settings.get().webSearch === true : false,
    });
  });

  routes.set("POST /api/engine/reset", async (_req, res) => {
    await engine.resetConversation();
    json(res, 200, { ok: true });
  });

  /** The one preference the pane can change: whether the model may search the web. */
  routes.set("POST /api/settings", async (req, res) => {
    if (!settings) return json(res, 501, { error: "Settings are not available." });
    let body;
    try {
      body = await readJsonBody(req, { limit: 4096 });
    } catch (err) {
      return json(res, 400, { error: err.message });
    }
    if (typeof body.webSearch !== "boolean") {
      return json(res, 400, { error: "webSearch must be true or false" });
    }
    const next = await settings.set({ webSearch: body.webSearch });
    json(res, 200, { ok: true, webSearch: next.webSearch === true });
  });
}
