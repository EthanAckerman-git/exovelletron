/**
 * Static file serving with traversal protection.
 *
 * Everything is read from a fixed root; any resolved path that escapes that root is
 * refused before it touches the filesystem.
 */
import { createReadStream } from "node:fs";
import { stat } from "node:fs/promises";
import path from "node:path";

const TYPES = new Map(Object.entries({
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".woff2": "font/woff2",
  ".map": "application/json; charset=utf-8",
}));

export const contentTypeFor = (file) => TYPES.get(path.extname(file).toLowerCase()) ?? "application/octet-stream";

/**
 * Resolve `urlPath` inside `root`, or null when it escapes.
 * @param {string} root absolute directory
 * @param {string} urlPath url path, may be percent-encoded
 */
export function resolveWithinRoot(root, urlPath) {
  let decoded;
  try {
    decoded = decodeURIComponent(urlPath.split("?")[0]);
  } catch {
    return null;
  }
  if (decoded.includes("\0")) return null;

  const normalizedRoot = path.resolve(root);
  const resolved = path.resolve(normalizedRoot, "." + path.posix.normalize("/" + decoded));

  if (resolved !== normalizedRoot && !resolved.startsWith(normalizedRoot + path.sep)) return null;
  return resolved;
}

/**
 * Stream a file from `root`. Returns false when there is nothing to serve, so the
 * caller can fall through to a 404.
 */
export async function serveStatic(root, urlPath, res, { cache = "no-store" } = {}) {
  const file = resolveWithinRoot(root, urlPath);
  if (!file) {
    res.writeHead(403).end("Forbidden");
    return true;
  }
  let info;
  try {
    info = await stat(file);
  } catch {
    return false;
  }
  if (!info.isFile()) return false;

  res.writeHead(200, {
    "Content-Type": contentTypeFor(file),
    "Content-Length": info.size,
    "Cache-Control": cache,
    "X-Content-Type-Options": "nosniff",
  });
  createReadStream(file).pipe(res);
  return true;
}
