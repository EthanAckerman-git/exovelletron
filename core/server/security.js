/**
 * Request authorisation for the local API.
 *
 * The server only ever binds to 127.0.0.1, so it is unreachable from the network. That
 * still leaves other software on this Mac — and any website the user visits, since
 * https://localhost is a valid cross-origin target. Two independent checks close that:
 *
 *  1. A random per-run session token, injected into the task pane HTML at serve time
 *     and required on every /api/ call. A page that cannot read our HTML cannot guess
 *     it, and it dies with the process.
 *  2. An origin allowlist, so a browser tab on some other site cannot drive the API
 *     even if a token ever leaked.
 */
import { randomBytes, timingSafeEqual } from "node:crypto";

export const TOKEN_HEADER = "x-eal-token";

export const mintToken = () => randomBytes(32).toString("base64url");

/** Constant-time compare that tolerates length mismatch without throwing. */
export function tokensMatch(a, b) {
  if (typeof a !== "string" || typeof b !== "string") return false;
  const ba = Buffer.from(a);
  const bb = Buffer.from(b);
  if (ba.length !== bb.length) return false;
  return timingSafeEqual(ba, bb);
}

/**
 * Origins allowed to call the API: our own origin (the task pane, same-origin) and
 * Excel's task pane host, which on some builds sends its own origin.
 */
export function isAllowedOrigin(origin, port) {
  if (!origin || origin === "null") return true; // same-origin requests often omit it
  const allowed = [
    `https://localhost:${port}`,
    `https://127.0.0.1:${port}`,
  ];
  return allowed.includes(origin);
}

/**
 * @returns {{ok: true} | {ok: false, status: number, message: string}}
 */
export function authorizeApiRequest(req, { token, port }) {
  const origin = req.headers.origin;
  if (!isAllowedOrigin(origin, port)) {
    return { ok: false, status: 403, message: "Origin not allowed" };
  }
  const supplied = req.headers[TOKEN_HEADER];
  if (!tokensMatch(Array.isArray(supplied) ? supplied[0] : supplied, token)) {
    return { ok: false, status: 401, message: "Missing or invalid session token" };
  }
  return { ok: true };
}

/** Headers applied to every response. */
export function securityHeaders(port) {
  return {
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "no-referrer",
    // Everything the pane loads is served by us; connect-src is limited to our origin
    // so the pane structurally cannot reach the internet even if code tried to.
    "Content-Security-Policy": [
      "default-src 'self'",
      "script-src 'self' 'unsafe-inline' 'unsafe-eval'",
      "style-src 'self' 'unsafe-inline'",
      "img-src 'self' data:",
      "font-src 'self' data:",
      `connect-src 'self' https://localhost:${port} https://127.0.0.1:${port}`,
      "frame-ancestors 'self' https://*.officeapps.live.com https://*.microsoft.com",
    ].join("; "),
  };
}
