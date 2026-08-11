/** Model listing, download control, and activation. */
import { getModel, DEFAULT_MODEL_ID } from "../../models/catalog.js";
import { recommendTiers } from "../../models/recommend.js";
import { detectChip } from "../../models/machine.js";
import { sseFrame } from "./chat.js";

export function registerModelRoutes(routes, { engine, models, json, readJsonBody, settings }) {
  // The chip never changes while the server runs; detect once per process.
  let chipPromise = null;
  const chip = () => (chipPromise ??= detectChip());

  routes.set("GET /api/models", async (_req, res) => {
    const memory = await engine.memoryInfo();
    const list = await models.list({ availableBytes: memory.total, contextTokens: engine.status().contextTokens });
    json(res, 200, {
      models: list,
      tiers: recommendTiers(list),
      machine: { chip: (await chip()).chip, availableBytes: memory.total },
      activeId: engine.modelId,
      downloadingId: models.activeDownloadId,
      memory,
    });
  });

  /** Live download progress. Kept separate from the chat stream so either can run alone. */
  routes.set("GET /api/models/events", async (req, res) => {
    res.writeHead(200, {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-store",
      Connection: "keep-alive",
    });
    res.write(sseFrame("hello", { at: Date.now() }));

    const forward = (event) => (payload) => res.write(sseFrame(event, payload));
    const handlers = {
      progress: forward("progress"),
      installed: forward("installed"),
      removed: forward("removed"),
      error: forward("error"),
    };
    for (const [name, fn] of Object.entries(handlers)) models.on(name, fn);

    const onEngineState = (s) => res.write(sseFrame("engine", s));
    engine.on("state", onEngineState);

    // Some proxies and webviews drop an idle stream; a periodic comment keeps it warm.
    const keepAlive = setInterval(() => res.write(": ping\n\n"), 20000);

    // The response's close is the reliable "subscriber left" signal; the request's
    // own close can fire as soon as its (empty) body is done.
    res.on("close", () => {
      clearInterval(keepAlive);
      for (const [name, fn] of Object.entries(handlers)) models.off(name, fn);
      engine.off("state", onEngineState);
      res.end();
    });
  });

  routes.set("POST /api/models/download", async (req, res) => {
    const { id } = await readJsonBody(req);
    if (!getModel(id)) return json(res, 400, { error: `Unknown model: ${id}` });
    if (models.activeDownloadId) return json(res, 409, { error: `Already downloading ${models.activeDownloadId}.` });

    // Respond immediately; the client follows progress on the events stream.
    models.download(id).catch(() => { /* surfaced via the error event */ });
    json(res, 202, { ok: true, id });
  });

  routes.set("POST /api/models/cancel", async (req, res) => {
    const { id } = await readJsonBody(req);
    json(res, 200, { ok: models.cancel(id) });
  });

  routes.set("POST /api/models/remove", async (req, res) => {
    const { id } = await readJsonBody(req);
    const model = getModel(id);
    if (!model) return json(res, 400, { error: `Unknown model: ${id}` });
    if (engine.modelId === id) {
      if (engine.isBusy) return json(res, 409, { error: "That model is answering a message right now." });
      await engine.unload();
    }
    await models.remove(id);
    // A config pointing at a deleted model would silently boot with nothing.
    if (settings && settings.get().modelId === id) await settings.set({ modelId: DEFAULT_MODEL_ID });
    json(res, 200, { ok: true });
  });

  routes.set("POST /api/models/select", async (req, res) => {
    const { id } = await readJsonBody(req);
    const model = getModel(id);
    if (!model) return json(res, 400, { error: `Unknown model: ${id}` });
    if (!(await models.isInstalled(id))) return json(res, 409, { error: `${model.name} is not downloaded yet.` });
    if (engine.isBusy) return json(res, 409, { error: "Finish the current message before switching models." });

    await engine.load(id, models.pathFor(id));
    // Persist the choice, exactly as the desktop panel's path does — a switch made
    // from the task pane used to silently revert on the next launch.
    if (settings) await settings.set({ modelId: id });
    json(res, 200, { ok: true, activeId: engine.modelId });
  });
}
