import { describe, it, expect, vi } from "vitest";
import {
  parseSearchHtml,
  decodeResultUrl,
  formatSearchResults,
  searchWeb,
} from "../../core/search.js";

/** A trimmed slice of what html.duckduckgo.com actually returns. */
const FIXTURE = `
<div class="result results_links web-result">
  <h2 class="result__title">
    <a rel="nofollow" class="result__a" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fwww.usps.com%2F&amp;rut=abc123">USPS - United States <b>Postal</b> Service</a>
  </h2>
  <a class="result__snippet" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fwww.usps.com%2F">Official site for tracking &amp; ZIP lookup.</a>
</div>
<div class="result results_links web-result">
  <h2 class="result__title">
    <a rel="nofollow" class="result__a" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fpe.usps.com%2Ftext%2Fpub28%2Fwelcome.htm&amp;rut=def">Publication 28 - Postal Addressing Standards</a>
  </h2>
  <a class="result__snippet" href="#">Standards for <b>addresses</b>.</a>
</div>`;

describe("decodeResultUrl", () => {
  it("unwraps the uddg redirect", () => {
    expect(decodeResultUrl("//duckduckgo.com/l/?uddg=https%3A%2F%2Fwww.usps.com%2F&amp;rut=abc"))
      .toBe("https://www.usps.com/");
  });

  it("upgrades protocol-relative links and passes plain URLs through", () => {
    expect(decodeResultUrl("//example.com/page")).toBe("https://example.com/page");
    expect(decodeResultUrl("https://example.com/x")).toBe("https://example.com/x");
  });
});

describe("parseSearchHtml", () => {
  it("extracts titles, decoded urls, and de-tagged snippets in order", () => {
    const results = parseSearchHtml(FIXTURE);
    expect(results).toHaveLength(2);
    expect(results[0]).toEqual({
      title: "USPS - United States Postal Service",
      url: "https://www.usps.com/",
      snippet: "Official site for tracking & ZIP lookup.",
    });
    expect(results[1].url).toBe("https://pe.usps.com/text/pub28/welcome.htm");
  });

  it("respects the result limit", () => {
    expect(parseSearchHtml(FIXTURE, 1)).toHaveLength(1);
  });

  it("returns nothing for a page with no results", () => {
    expect(parseSearchHtml("<html><body>No results.</body></html>")).toEqual([]);
  });
});

describe("formatSearchResults", () => {
  it("numbers results with their urls and snippets", () => {
    const text = formatSearchResults("usps", parseSearchHtml(FIXTURE));
    expect(text).toContain("1. USPS - United States Postal Service");
    expect(text).toContain("https://www.usps.com/");
    expect(text).toContain("2. Publication 28");
  });

  it("says plainly when nothing was found", () => {
    expect(formatSearchResults("xyzzy", [])).toMatch(/No web results/);
  });
});

describe("searchWeb", () => {
  it("returns formatted results on success", async () => {
    const fetchImpl = vi.fn(async () => ({ ok: true, text: async () => FIXTURE }));
    const out = await searchWeb("usps address format", { fetchImpl });
    expect(out).toContain("1. USPS");
    expect(fetchImpl.mock.calls[0][0]).toContain("q=usps%20address%20format");
  });

  // The handler's return value goes straight to the model, so failures must come
  // back as instructions the model can act on, never as a thrown error.
  it("turns an offline failure into text the model can relay", async () => {
    const fetchImpl = vi.fn(async () => { throw new Error("fetch failed: ENOTFOUND html.duckduckgo.com"); });
    const out = await searchWeb("anything", { fetchImpl });
    expect(out).toMatch(/offline/i);
  });

  it("reports a non-200 without throwing", async () => {
    const fetchImpl = vi.fn(async () => ({ ok: false, status: 503 }));
    expect(await searchWeb("anything", { fetchImpl })).toMatch(/HTTP 503/);
  });

  it("refuses an empty query without a network call", async () => {
    const fetchImpl = vi.fn();
    expect(await searchWeb("   ", { fetchImpl })).toMatch(/empty/);
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});
