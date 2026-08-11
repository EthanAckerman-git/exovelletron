/**
 * Page reading for the model — the second half of the opt-in web toggle.
 *
 * When the user pastes a link, or a search result plainly holds the answer, the model
 * can open that one page and read its text. What leaves the Mac is the URL being
 * fetched; what comes back is text, clamped hard so a hostile or enormous page cannot
 * flood the context window.
 *
 * The URL guard is a hostname-pattern check, not resolve-then-verify. That is a
 * deliberate fit to this threat model: the only service on this machine is our own
 * loopback-bound, token-authenticated server, and the model cannot learn the token.
 * The patterns close off loopback, RFC-1918 ranges, link-local, and mDNS names;
 * redirects are followed manually so every hop is re-checked against the same rules.
 */
import { decodeResultUrl } from "./search.js";

export const MAX_RESPONSE_BYTES = 2 * 1024 * 1024;
export const TIMEOUT_MS = 15_000;
export const MAX_REDIRECTS = 3;
export const DEFAULT_MAX_CHARS = 7_000;

const USER_AGENT =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15";

const READABLE_TYPES = ["text/html", "application/xhtml+xml", "text/plain"];

/** Hostnames and ranges that must never be reachable through the model. */
const BLOCKED_HOST_PATTERNS = [
  /^localhost$/i,
  /\.localhost$/i,
  /\.local$/i,
  /^0\.0\.0\.0$/,
  /^127\./,
  /^10\./,
  /^192\.168\./,
  /^169\.254\./,
  /^172\.(1[6-9]|2\d|3[01])\./,
];

/**
 * @param {string} rawUrl
 * @returns {{ok:true, url:URL} | {ok:false, reason:string}}
 */
export function isAllowedUrl(rawUrl) {
  let url;
  try {
    url = new URL(String(rawUrl ?? "").trim());
  } catch {
    return { ok: false, reason: "it is not a valid web address" };
  }
  if (url.protocol !== "http:" && url.protocol !== "https:") {
    return { ok: false, reason: `only http and https pages can be opened, not ${url.protocol.replace(":", "")}` };
  }
  // Any IPv6 literal is refused outright — legitimate pages on raw IPv6 addresses are
  // vanishingly rare, and allowing them would mean re-implementing the ranges above.
  if (url.hostname.includes("[") || url.hostname.includes(":")) {
    return { ok: false, reason: "raw IPv6 addresses are not allowed" };
  }
  if (BLOCKED_HOST_PATTERNS.some((re) => re.test(url.hostname))) {
    return { ok: false, reason: "local and private network addresses are not allowed" };
  }
  return { ok: true, url };
}

/** Minimal entity decoding, including the numeric forms real pages actually use. */
const decodeEntities = (s) => String(s)
  .replace(/&#x([0-9a-f]+);/gi, (_, hex) => safeCodePoint(parseInt(hex, 16)))
  .replace(/&#(\d+);/g, (_, dec) => safeCodePoint(parseInt(dec, 10)))
  .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
  .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&apos;/g, "'")
  .replace(/&nbsp;/g, " ").replace(/&mdash;/g, "—").replace(/&ndash;/g, "–")
  .replace(/&rsquo;/g, "'").replace(/&lsquo;/g, "'")
  .replace(/&rdquo;/g, "”").replace(/&ldquo;/g, "“");

const safeCodePoint = (n) => {
  try { return Number.isFinite(n) && n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : ""; } catch { return ""; }
};

/** Subtrees that are chrome or code, never content. Removed before text extraction. */
const DROP_SUBTREES = /<(script|style|noscript|template|svg|iframe|object|embed|nav|header|footer|form|aside)\b[\s\S]*?<\/\1>/gi;

/**
 * Turn an HTML page into readable text: strip the chrome, keep the paragraph and
 * heading structure, decode entities, and collapse the whitespace storm that remains.
 */
export function htmlToText(html) {
  let s = String(html ?? "");

  const title = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(s)?.[1];

  // Dropping subtrees repeats because these elements nest (a header inside a nav);
  // the cap keeps a pathological page from looping us.
  for (let i = 0; i < 5 && DROP_SUBTREES.test(s); i++) s = s.replace(DROP_SUBTREES, " ");
  s = s.replace(/<!--[\s\S]*?-->/g, " ");

  s = s
    .replace(/<(?:br|hr)\s*\/?>/gi, "\n")
    .replace(/<li\b[^>]*>/gi, "\n- ")
    .replace(/<\/(?:p|div|h[1-6]|li|tr|table|ul|ol|blockquote|section|article|dt|dd)>/gi, "\n")
    .replace(/<\/(?:td|th)>/gi, "  ")
    .replace(/<[^>]*>/g, " ");

  s = decodeEntities(s)
    .split("\n")
    .map((line) => line.replace(/[ \t ]+/g, " ").trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return title ? `${decodeEntities(title).replace(/\s+/g, " ").trim()}\n\n${s}` : s;
}

/** Honest truncation: the model is told exactly how much of the page it is seeing. */
export function clampText(text, maxChars) {
  if (text.length <= maxChars) return text;
  return `${text.slice(0, maxChars)}\n\n(page truncated — showing the first ${maxChars.toLocaleString("en-US")} of ${text.length.toLocaleString("en-US")} characters)`;
}

/** Read a body stream up to the byte cap, then stop rather than buffering the rest. */
async function readCapped(res) {
  if (!res.body?.getReader) {
    const text = await res.text();
    return text.slice(0, MAX_RESPONSE_BYTES);
  }
  const reader = res.body.getReader();
  const decoder = new TextDecoder("utf-8", { fatal: false });
  let out = "";
  let bytes = 0;
  while (bytes < MAX_RESPONSE_BYTES) {
    const { done, value } = await reader.read();
    if (done) break;
    bytes += value.byteLength;
    out += decoder.decode(value, { stream: true });
  }
  try { await reader.cancel(); } catch { /* already done */ }
  return out + decoder.decode();
}

/**
 * Open one web page and return its text, or a model-readable explanation of why not.
 * Never throws — a failed fetch should end in the model telling the user, not in a
 * crashed turn.
 *
 * @param {string} rawUrl
 * @param {{fetchImpl?:typeof fetch, maxChars?:number}} [opts]  injectable for tests
 */
export async function fetchPage(rawUrl, { fetchImpl = fetch, maxChars = DEFAULT_MAX_CHARS } = {}) {
  const given = String(rawUrl ?? "").trim().slice(0, 2000);
  if (!given) return "No URL was given.";

  // Search results carry DuckDuckGo redirect URLs; unwrap before judging the target.
  let verdict = isAllowedUrl(decodeResultUrl(given));
  if (!verdict.ok) return `That URL cannot be opened — ${verdict.reason}. Tell the user.`;

  try {
    let url = verdict.url;
    let res = null;

    // Follow redirects by hand so every hop faces the same allow-list.
    for (let hop = 0; hop <= MAX_REDIRECTS; hop++) {
      res = await fetchImpl(url.href, {
        headers: { "User-Agent": USER_AGENT, Accept: "text/html,application/xhtml+xml,text/plain" },
        redirect: "manual",
        signal: AbortSignal.timeout(TIMEOUT_MS),
      });
      if (res.status < 300 || res.status >= 400) break;
      const location = res.headers.get("location");
      if (!location) break;
      if (hop === MAX_REDIRECTS) return "That page redirected too many times. Tell the user.";
      verdict = isAllowedUrl(new URL(location, url).href);
      if (!verdict.ok) return `That URL cannot be opened — it redirects somewhere that ${verdict.reason}. Tell the user.`;
      url = verdict.url;
    }

    if (!res.ok) return `Opening the page failed (HTTP ${res.status}). Tell the user and answer from what you know.`;

    const contentType = (res.headers.get("content-type") ?? "").split(";")[0].trim().toLowerCase();
    if (contentType && !READABLE_TYPES.includes(contentType)) {
      return `That page is not readable text (it is ${contentType}). Tell the user.`;
    }

    const body = await readCapped(res);
    const text = contentType === "text/plain" ? body.trim() : htmlToText(body);
    if (!text) return "The page loaded but contained no readable text. Tell the user.";

    return `Contents of ${url.href}:\n\n${clampText(text, maxChars)}`;
  } catch (err) {
    const offline = /fetch failed|ENOTFOUND|ECONNREFUSED|network|timeout|abort/i.test(String(err?.message ?? err));
    return offline
      ? "The page could not be opened — this Mac appears to be offline. Tell the user and answer from what you know."
      : `Opening the page failed: ${err.message}. Tell the user and answer from what you know.`;
  }
}
