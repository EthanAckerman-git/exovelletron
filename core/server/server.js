/**
 * The local HTTPS server: serves the task pane and the API from one origin.
 *
 * Dependencies (engine, model store, config) are injected rather than imported so the
 * whole surface can be exercised in tests against lightweight stubs.
 */
import { createServer } from "node:https";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { serveStatic } from "./static.js";
import { mintToken, authorizeApiRequest, securityHeaders, TOKEN_HEADER } from "./security.js";
import { registerStatusRoutes } from "./routes/status.js";
import { registerChatRoutes } from "./routes/chat.js";
import { registerModelRoutes } from "./routes/models.js";

const DIST = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "dist");

export const TOKEN_PLACEHOLDER = "__EAL_SESSION_TOKEN__";

export function json(res, status, body) {
  const payload = JSON.stringify(body);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(payload),
    "Cache-Control": "no-store",
  });
  res.end(payload);
}

export async function readJsonBody(req, { limit = 4_000_000 } = {}) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > limit) throw new Error("Request body too large");
    chunks.push(chunk);
  }
  if (!chunks.length) return {};
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    throw new Error("Request body is not valid JSON");
  }
}

/**
 * @param {object} deps
 * @param {{key: Buffer, cert: Buffer}} deps.credentials
 * @param {number} deps.port
 * @param {object} deps.engine     inference engine (see core/llm/engine.js)
 * @param {object} deps.models     model store (see core/models/store.js)
 * @param {object} [deps.appInfo]  { version }
 */
export function createAppServer({ credentials, port, engine, models, appInfo = {} }) {
  const token = mintToken();
  /** @type {Map<string, (req, res, ctx) => Promise<boolean|void>>} */
  const routes = new Map();

  const ctx = { engine, models, token, port, appInfo, json, readJsonBody };

  registerStatusRoutes(routes, ctx);
  registerChatRoutes(routes, ctx);
  registerModelRoutes(routes, ctx);

  async function serveHtml(file, res) {
    let html = await readFile(path.join(DIST, file), "utf8");
    html = html.replaceAll(TOKEN_PLACEHOLDER, token);
    const body = Buffer.from(html, "utf8");
    res.writeHead(200, {
      "Content-Type": "text/html; charset=utf-8",
      "Content-Length": body.length,
      // The token is embedded, so this must never be cached to disk by the webview.
      "Cache-Control": "no-store, no-cache, must-revalidate",
      ...securityHeaders(port),
    });
    res.end(body);
  }

  async function handle(req, res) {
    const url = new URL(req.url, `https://localhost:${port}`);
    const pathname = url.pathname;

    if (req.method === "OPTIONS") {
      res.writeHead(204, {
        "Access-Control-Allow-Origin": `https://localhost:${port}`,
        "Access-Control-Allow-Headers": `Content-Type, ${TOKEN_HEADER}`,
        "Access-Control-Allow-Methods": "GET, POST, DELETE, OPTIONS",
      }).end();
      return;
    }

    if (pathname.startsWith("/api/")) {
      const auth = authorizeApiRequest(req, { token, port });
      if (!auth.ok) return json(res, auth.status, { error: auth.message });

      const key = `${req.method} ${pathname}`;
      const handler = routes.get(key);
      if (!handler) return json(res, 404, { error: `No such endpoint: ${key}` });
      try {
        await handler(req, res, { ...ctx, url });
      } catch (err) {
        if (!res.headersSent) json(res, 500, { error: err.message || "Internal error" });
        else res.end();
      }
      return;
    }

    if (req.method !== "GET" && req.method !== "HEAD") {
      return json(res, 405, { error: "Method not allowed" });
    }

    if (pathname === "/" || pathname === "/index.html") return serveHtml("landing.html", res);
    if (pathname === "/taskpane.html") return serveHtml("taskpane.html", res);
    if (pathname === "/commands.html") return serveHtml("commands.html", res);

    // Long-lived immutable assets: the Office runtime and our icons.
    if (pathname.startsWith("/vendor/")) {
      if (await serveStatic(path.join(DIST, "vendor"), pathname.slice("/vendor".length), res, {
        cache: "public, max-age=31536000, immutable",
      })) return;
    }
    if (pathname.startsWith("/assets/")) {
      if (await serveStatic(path.join(DIST, "assets"), pathname.slice("/assets".length), res, {
        cache: "public, max-age=86400",
      })) return;
    }
    if (pathname.startsWith("/app/")) {
      if (await serveStatic(path.join(DIST, "app"), pathname.slice("/app".length), res)) return;
    }

    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" }).end("Not found");
  }

  const server = createServer({ key: credentials.key, cert: credentials.cert }, (req, res) => {
    handle(req, res).catch((err) => {
      if (!res.headersSent) res.writeHead(500, { "Content-Type": "text/plain" }).end(String(err.message));
      else res.end();
    });
  });

  // Streaming a long generation must not be killed by an idle timeout.
  server.headersTimeout = 0;
  server.requestTimeout = 0;
  server.timeout = 0;

  return {
    token,
    port,
    server,
    /** @returns {Promise<number>} the bound port */
    listen() {
      return new Promise((resolve, reject) => {
        const onError = (err) => {
          server.removeListener("listening", onListening);
          reject(
            err.code === "EADDRINUSE"
              ? Object.assign(new Error(`Port ${port} is already in use.`), { code: "EADDRINUSE", port })
              : err,
          );
        };
        const onListening = () => {
          server.removeListener("error", onError);
          resolve(server.address().port);
        };
        server.once("error", onError);
        server.once("listening", onListening);
        // 127.0.0.1 only: never reachable from the network.
        server.listen(port, "127.0.0.1");
      });
    },
    close() {
      return new Promise((resolve) => {
        server.closeAllConnections?.();
        server.close(() => resolve());
      });
    },
  };
}
