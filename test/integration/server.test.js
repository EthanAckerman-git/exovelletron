import { describe, it, expect, beforeAll, afterAll } from "vitest";
import { request as httpsRequest } from "node:https";
import { mkdtemp, rm, mkdir, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { EventEmitter } from "node:events";

import { createPaths } from "../../core/config.js";
import { generateCerts, certFiles } from "../../core/setup/certs.js";
import { createAppServer } from "../../core/server/server.js";
import { readFile } from "node:fs/promises";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");

/** Minimal https client that accepts our throwaway self-signed certificate. */
function call(port, urlPath, { method = "GET", headers = {}, body } = {}) {
  return new Promise((resolve, reject) => {
    const req = httpsRequest(
      { host: "127.0.0.1", port, path: urlPath, method, headers, rejectUnauthorized: false },
      (res) => {
        const chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () =>
          resolve({ status: res.statusCode, headers: res.headers, text: Buffer.concat(chunks).toString("utf8") }));
      },
    );
    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}

/** Stand-in engine: deterministic, no model, streams two chunks and one action. */
class StubEngine extends EventEmitter {
  modelId = "qwen3.5-4b";
  isBusy = false;
  loads = [];
  unloads = 0;
  status() { return { state: "ready", modelId: this.modelId, error: null, contextTokens: 8192, busy: false }; }
  async memoryInfo() { return { total: 19e9, free: 12e9, used: 7e9 }; }
  async load(id, modelPath) { this.loads.push({ id, modelPath }); this.modelId = id; }
  async unload() { this.unloads += 1; this.modelId = null; }
  async resetConversation() {}
  abort() {}
  async chat({ onToken }) {
    onToken("Hello ");
    onToken("world");
    const action = { id: "a1", type: "write_formula", sheet: null, address: "A1:A2", formula: "=1", summary: "Fill 2 cells" };
    this.emit("action", action);
    this.emit("search", { query: "usps address format" });
    this.emit("fetch", { url: "https://pe.usps.com/businessmail101" });
    return { text: "Hello world", actions: [action], stats: { seconds: 0.1, tokens: 2, tokensPerSecond: 20 } };
  }
}

/**
 * Three real catalog ids decorated the way ModelStore.list() decorates them, with
 * fit levels chosen so the tier maths has one comfortable pick, one stretch, and
 * one that must never be recommended.
 */
const stubModels = {
  activeDownloadId: null,
  installed: true,
  removed: [],
  async list() {
    return [
      { id: "qwen3.5-2b", bytes: 1.3e9, intelligence: 1, installed: true, fit: { level: "great", label: "Plenty of room" } },
      { id: "qwen3.5-4b", bytes: 2.9e9, intelligence: 2, installed: true, fit: { level: "ok", label: "Fits" } },
      { id: "qwen3.5-9b", bytes: 6.0e9, intelligence: 3, installed: false, fit: { level: "too-big", label: "Too large" } },
    ];
  },
  async isInstalled() { return this.installed; },
  pathFor(id) { return `/dev/null/${id}.gguf`; },
  async remove(id) { this.removed.push(id); },
  download() { return Promise.resolve(); },
  cancel() { return false; },
  on() {}, off() {},
};

let server;
let port;
let token;
let tmp;
let serverEngine;
let prefs;
const serverPrefs = () => prefs;

beforeAll(async () => {
  tmp = await mkdtemp(path.join(tmpdir(), "eal-test-"));
  const paths = createPaths(tmp);
  await mkdir(paths.certsDir, { recursive: true });
  await generateCerts(paths);
  const f = certFiles(paths);

  // The server serves dist/; make sure the pages it needs exist for this run.
  await mkdir(path.join(ROOT, "dist", "app"), { recursive: true });
  for (const [file, contents] of [
    ["landing.html", "<!doctype html><title>landing</title>"],
    ["taskpane.html", '<!doctype html><script>window.__EAL_TOKEN__="__EAL_SESSION_TOKEN__";</script>'],
    ["commands.html", "<!doctype html><title>commands</title>"],
  ]) {
    const target = path.join(ROOT, "dist", file);
    try { await readFile(target); } catch { await writeFile(target, contents, "utf8"); }
  }

  prefs = { webSearch: false, modelId: "qwen3.5-4b" };
  serverEngine = new StubEngine();
  server = createAppServer({
    credentials: { key: await readFile(f.key), cert: await readFile(f.cert) },
    port: 0,
    engine: serverEngine,
    models: stubModels,
    appInfo: { version: "test" },
    settings: { get: () => prefs, set: async (patch) => (prefs = { ...prefs, ...patch }) },
  });
  port = await server.listen();
  token = server.token;
}, 60_000);

afterAll(async () => {
  await server?.close();
  await rm(tmp, { recursive: true, force: true });
});

const auth = () => ({ "x-eal-token": token, "Content-Type": "application/json" });

describe("local server", () => {
  it("binds loopback only", () => {
    expect(server.server.address().address).toBe("127.0.0.1");
  });

  it("serves the landing page", async () => {
    const res = await call(port, "/");
    expect(res.status).toBe(200);
    expect(res.headers["content-type"]).toMatch(/text\/html/);
  });

  it("injects the session token into the task pane and forbids caching it", async () => {
    const res = await call(port, "/taskpane.html");
    expect(res.status).toBe(200);
    expect(res.text).toContain(token);
    expect(res.text).not.toContain("__EAL_SESSION_TOKEN__");
    expect(res.headers["cache-control"]).toMatch(/no-store/);
  });

  it("sets a content security policy that pins connect-src to this origin", async () => {
    const res = await call(port, "/taskpane.html");
    expect(res.headers["content-security-policy"]).toContain(`connect-src 'self' https://localhost:${port}`);
  });

  it("rejects API calls with no token", async () => {
    const res = await call(port, "/api/status");
    expect(res.status).toBe(401);
  });

  it("rejects API calls with a wrong token", async () => {
    const res = await call(port, "/api/status", { headers: { "x-eal-token": "nope" } });
    expect(res.status).toBe(401);
  });

  it("rejects a foreign origin even with a valid token", async () => {
    const res = await call(port, "/api/status", { headers: { ...auth(), origin: "https://evil.example" } });
    expect(res.status).toBe(403);
  });

  it("reports status when authorized", async () => {
    const res = await call(port, "/api/status", { headers: auth() });
    expect(res.status).toBe(200);
    const body = JSON.parse(res.text);
    expect(body.ok).toBe(true);
    expect(body.engine.state).toBe("ready");
    expect(body.machine.totalRamGb).toBeGreaterThan(0);
  });

  it("streams a chat turn as SSE with tokens, an action, and a done frame", async () => {
    const res = await call(port, "/api/chat", {
      method: "POST",
      headers: auth(),
      body: JSON.stringify({ message: "hi" }),
    });
    expect(res.status).toBe(200);
    expect(res.headers["content-type"]).toMatch(/text\/event-stream/);
    expect(res.text).toContain("event: token");
    expect(res.text).toContain("event: action");
    expect(res.text).toContain("event: done");
    expect(res.text).toContain("Hello ");
    // Web-tool receipts ride the same stream so the pane can show them.
    expect(res.text).toContain("event: search");
    expect(res.text).toContain("event: fetch");
    expect(res.text).toContain("pe.usps.com");

    const done = res.text.split("event: done\ndata: ")[1].split("\n")[0];
    expect(JSON.parse(done).actions).toHaveLength(1);
  });

  it("rejects an empty message", async () => {
    const res = await call(port, "/api/chat", { method: "POST", headers: auth(), body: JSON.stringify({ message: "  " }) });
    expect(res.status).toBe(400);
  });

  it("reports web search off by default and lets the pane flip it", async () => {
    let status = JSON.parse((await call(port, "/api/status", { headers: auth() })).text);
    expect(status.webSearch).toBe(false);

    const on = await call(port, "/api/settings", {
      method: "POST", headers: auth(), body: JSON.stringify({ webSearch: true }),
    });
    expect(JSON.parse(on.text)).toMatchObject({ ok: true, webSearch: true });

    status = JSON.parse((await call(port, "/api/status", { headers: auth() })).text);
    expect(status.webSearch).toBe(true);

    await call(port, "/api/settings", { method: "POST", headers: auth(), body: JSON.stringify({ webSearch: false }) });
  });

  it("rejects a non-boolean web search setting", async () => {
    const res = await call(port, "/api/settings", {
      method: "POST", headers: auth(), body: JSON.stringify({ webSearch: "yes" }),
    });
    expect(res.status).toBe(400);
  });

  it("lists models with fit, tiers, and the machine they were graded against", async () => {
    const res = await call(port, "/api/models", { headers: auth() });
    expect(res.status).toBe(200);
    const body = JSON.parse(res.text);
    expect(body.models).toHaveLength(3);
    for (const m of body.models) expect(m.fit.level).toBeTruthy();
    // Highest intelligence that fits is Max; the comfortable one below it is Balanced;
    // nothing smaller with plenty of room remains for Fast. Too-big is never picked.
    expect(body.tiers).toEqual({ fast: null, balanced: "qwen3.5-2b", max: "qwen3.5-4b" });
    expect(body.machine.availableBytes).toBe(19e9);
    expect(body.activeId).toBe("qwen3.5-4b");
  });

  it("selecting a model loads it and persists the choice", async () => {
    const res = await call(port, "/api/models/select", {
      method: "POST", headers: auth(), body: JSON.stringify({ id: "qwen3.5-2b" }),
    });
    expect(res.status).toBe(200);
    expect(JSON.parse(res.text).activeId).toBe("qwen3.5-2b");
    expect(serverEngine.loads.at(-1)).toEqual({ id: "qwen3.5-2b", modelPath: "/dev/null/qwen3.5-2b.gguf" });
    // The regression this pins: a pane-side switch used to vanish on restart.
    expect(serverPrefs().modelId).toBe("qwen3.5-2b");
  });

  it("refuses to switch models mid-reply, and does not persist the attempt", async () => {
    const before = JSON.parse((await call(port, "/api/models", { headers: auth() })).text).activeId;
    serverEngine.isBusy = true;
    const res = await call(port, "/api/models/select", {
      method: "POST", headers: auth(), body: JSON.stringify({ id: "qwen3.5-4b" }),
    });
    serverEngine.isBusy = false;
    expect(res.status).toBe(409);
    expect(JSON.parse(res.text).error).toMatch(/Finish the current message/);
    expect(JSON.parse((await call(port, "/api/models", { headers: auth() })).text).activeId).toBe(before);
  });

  it("refuses to select a model that is not downloaded", async () => {
    stubModels.installed = false;
    const res = await call(port, "/api/models/select", {
      method: "POST", headers: auth(), body: JSON.stringify({ id: "qwen3.5-9b" }),
    });
    stubModels.installed = true;
    expect(res.status).toBe(409);
    expect(JSON.parse(res.text).error).toMatch(/not downloaded/);
  });

  it("rejects selecting or downloading an unknown model", async () => {
    for (const route of ["/api/models/select", "/api/models/download"]) {
      const res = await call(port, route, {
        method: "POST", headers: auth(), body: JSON.stringify({ id: "gpt-99" }),
      });
      expect(res.status, route).toBe(400);
    }
  });

  it("refuses a second concurrent download", async () => {
    stubModels.activeDownloadId = "qwen3.5-4b";
    const res = await call(port, "/api/models/download", {
      method: "POST", headers: auth(), body: JSON.stringify({ id: "qwen3.5-2b" }),
    });
    stubModels.activeDownloadId = null;
    expect(res.status).toBe(409);
  });

  it("removing the active model unloads it and resets the persisted choice", async () => {
    await call(port, "/api/models/select", {
      method: "POST", headers: auth(), body: JSON.stringify({ id: "qwen3.5-2b" }),
    });
    const unloadsBefore = serverEngine.unloads;
    const res = await call(port, "/api/models/remove", {
      method: "POST", headers: auth(), body: JSON.stringify({ id: "qwen3.5-2b" }),
    });
    expect(res.status).toBe(200);
    expect(serverEngine.unloads).toBe(unloadsBefore + 1);
    expect(stubModels.removed).toContain("qwen3.5-2b");
    expect(serverPrefs().modelId).toBe("qwen3.5-4b"); // back to the default, not dangling
  });

  it("404s unknown API endpoints", async () => {
    const res = await call(port, "/api/nope", { headers: auth() });
    expect(res.status).toBe(404);
  });

  it("refuses path traversal out of the served roots", async () => {
    for (const attack of ["/vendor/../../../../etc/passwd", "/app/../../package.json", "/assets/%2e%2e/%2e%2e/package.json"]) {
      const res = await call(port, attack);
      expect([403, 404], `${attack} returned ${res.status}`).toContain(res.status);
      expect(res.text).not.toContain("exovelletron");
    }
  });

  it("rejects non-GET on static routes", async () => {
    const res = await call(port, "/taskpane.html", { method: "DELETE" });
    expect(res.status).toBe(405);
  });
});
