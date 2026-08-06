/** Health and capability reporting for the task pane. */
import os from "node:os";
import { DEFAULT_MODEL_ID, getModel } from "../../models/catalog.js";

export function registerStatusRoutes(routes, { engine, models, appInfo, json }) {
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
    });
  });

  routes.set("POST /api/engine/reset", async (_req, res) => {
    await engine.resetConversation();
    json(res, 200, { ok: true });
  });
}
