/**
 * Row-by-row column transformation, streamed.
 *
 * The task pane reads the source values out of Excel and posts them here; the model
 * rewrites each one and the results stream back with progress. A long job needs to show
 * that it is alive, so progress frames are emitted as batches land.
 */
import { sseFrame } from "./chat.js";
import { MAX_TRANSFORM_ROWS } from "../../llm/transform.js";

export function registerTransformRoutes(routes, { engine, json, readJsonBody }) {
  routes.set("POST /api/transform", async (req, res) => {
    let body;
    try {
      body = await readJsonBody(req, { limit: 16_000_000 });
    } catch (err) {
      return json(res, 400, { error: err.message });
    }

    const values = Array.isArray(body.values) ? body.values.map((v) => (v == null ? "" : String(v))) : null;
    const instruction = typeof body.instruction === "string" ? body.instruction.trim() : "";
    const fields = Array.isArray(body.fields) && body.fields.length > 1
      ? body.fields.map((f) => String(f)).slice(0, 20)
      : null;

    if (!values?.length) return json(res, 400, { error: "values is required" });
    if (values.length > MAX_TRANSFORM_ROWS) {
      return json(res, 400, { error: `That is ${values.length} rows; the limit is ${MAX_TRANSFORM_ROWS}.` });
    }
    if (!instruction) return json(res, 400, { error: "instruction is required" });

    const status = engine.status();
    if (status.state !== "ready") return json(res, 503, { error: status.error || "No model is loaded." });
    if (status.busy) return json(res, 409, { error: "The model is busy." });

    res.writeHead(200, {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-store",
      Connection: "keep-alive",
      "X-Accel-Buffering": "no",
    });
    res.write(sseFrame("start", { total: values.length }));

    const controller = new AbortController();
    // A closed pane means nobody is waiting for the answer; stop burning GPU on it.
    const onClose = () => controller.abort();
    req.on("close", onClose);

    try {
      const { results, failed, lossy } = await engine.transform({
        values,
        instruction,
        fields,
        signal: controller.signal,
        onProgress: (p) => res.write(sseFrame("progress", p)),
      });
      res.write(sseFrame("done", { results, failed, lossy: lossy ?? [] }));
    } catch (err) {
      res.write(sseFrame("error", {
        message: err.code === "ABORTED" ? "Transform cancelled" : err.message || "Transform failed",
      }));
    } finally {
      req.off("close", onClose);
      res.end();
    }
  });

  routes.set("POST /api/transform/abort", async (_req, res) => {
    engine.abort();
    json(res, 200, { ok: true });
  });
}
