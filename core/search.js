/**
 * Web search for the model — strictly opt-in.
 *
 * The app's core promise is that nothing leaves the Mac, so search is OFF by default
 * and the model only receives this tool after the user flips the toggle in the pane.
 * When enabled, exactly one thing is transmitted: the search query, sent to
 * DuckDuckGo's HTML endpoint. DuckDuckGo because it needs no API key or account —
 * the app stays install-and-go — and does not profile users. Only the result list
 * (titles, URLs, snippets) comes back; no pages are fetched.
 *
 * The parser is pure and exported so it can be tested against a saved fixture.
 */

export const SEARCH_ENDPOINT = "https://html.duckduckgo.com/html/";
export const MAX_RESULTS = 5;
const TIMEOUT_MS = 12_000;

/** Minimal entity decoding for the handful DuckDuckGo actually emits. */
const decodeEntities = (s) => String(s)
  .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
  .replace(/&quot;/g, '"').replace(/&#x27;/g, "'").replace(/&#39;/g, "'")
  .replace(/&nbsp;/g, " ");

const stripTags = (s) => decodeEntities(String(s).replace(/<[^>]*>/g, "")).replace(/\s+/g, " ").trim();

/** Result links are redirect URLs like //duckduckgo.com/l/?uddg=<encoded>&rut=… */
export function decodeResultUrl(href) {
  const raw = decodeEntities(href);
  const m = /[?&]uddg=([^&]+)/.exec(raw);
  if (m) {
    try { return decodeURIComponent(m[1]); } catch { /* fall through */ }
  }
  if (raw.startsWith("//")) return `https:${raw}`;
  return raw;
}

/**
 * Pull results out of DuckDuckGo's HTML page.
 * @returns {Array<{title:string,url:string,snippet:string}>}
 */
export function parseSearchHtml(html, limit = MAX_RESULTS) {
  const results = [];
  const links = /<a[^>]*class="[^"]*result__a[^"]*"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g;
  const snippets = [...String(html).matchAll(
    /class="[^"]*result__snippet[^"]*"[^>]*>([\s\S]*?)<\/(?:a|div|td)>/g,
  )].map((m) => stripTags(m[1]));

  let match;
  while ((match = links.exec(String(html))) !== null && results.length < limit) {
    const url = decodeResultUrl(match[1]);
    const title = stripTags(match[2]);
    if (!title || !/^https?:/i.test(url)) continue;
    results.push({ title, url, snippet: snippets[results.length] ?? "" });
  }
  return results;
}

/** One compact text block the model can actually read. */
export function formatSearchResults(query, results) {
  if (!results.length) return `No web results found for "${query}".`;
  return results
    .map((r, i) => `${i + 1}. ${r.title}\n   ${r.url}${r.snippet ? `\n   ${r.snippet}` : ""}`)
    .join("\n");
}

/**
 * Run one search and return the results as model-readable text. Failures come back
 * as text too — the model should tell the user search is down, not crash the turn.
 * @param {string} query
 * @param {{fetchImpl?:typeof fetch}} [opts]  injectable for tests
 */
export async function searchWeb(query, { fetchImpl = fetch } = {}) {
  const q = String(query ?? "").trim().slice(0, 400);
  if (!q) return "The search query was empty.";

  try {
    const res = await fetchImpl(`${SEARCH_ENDPOINT}?q=${encodeURIComponent(q)}`, {
      headers: {
        // The endpoint serves a captcha page to clients with no User-Agent.
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
        Accept: "text/html",
      },
      signal: AbortSignal.timeout(TIMEOUT_MS),
    });
    if (!res.ok) return `Web search failed (HTTP ${res.status}). Tell the user and answer from what you know.`;
    return formatSearchResults(q, parseSearchHtml(await res.text()));
  } catch (err) {
    const offline = /fetch failed|ENOTFOUND|ECONNREFUSED|network|timeout|abort/i.test(String(err?.message ?? err));
    return offline
      ? "Web search is unavailable — this Mac appears to be offline. Tell the user and answer from what you know."
      : `Web search failed: ${err.message}. Tell the user and answer from what you know.`;
  }
}
