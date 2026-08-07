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
  status() { return { state: "ready", modelId: this.modelId, error: null, contextTokens: 8192, busy: false }; }
  async resetConversation() {}
  abort() {}
  async chat({ onToken }) {
    onToken("Hello ");
    onToken("world");
    const action = { id: "a1", type: "write_formula", sheet: null, address: "A1:A2", formula: "=1", summary: "Fill 2 cells" };
    this.emit("action", action);
    return { text: "Hello world", actions: [action], stats: { seconds: 0.1, tokens: 2, tokensPerSecond: 20 } };
  }
}

const stubModels = {
  activeDownloadId: null,
  async list() { return []; },
  async isInstalled() { return true; },
  pathFor() { return "/dev/null"; },
  on() {}, off() {},
};

let server;
let port;
let token;
let tmp;

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

  server = createAppServer({
    credentials: { key: await readFile(f.key), cert: await readFile(f.cert) },
    port: 0,
    engine: new StubEngine(),
    models: stubModels,
    appInfo: { version: "test" },
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

    const done = res.text.split("event: done\ndata: ")[1].split("\n")[0];
    expect(JSON.parse(done).actions).toHaveLength(1);
  });

  it("rejects an empty message", async () => {
    const res = await call(port, "/api/chat", { method: "POST", headers: auth(), body: JSON.stringify({ message: "  " }) });
    expect(res.status).toBe(400);
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
