import { describe, it, expect, vi } from "vitest";
import {
  isAllowedUrl, htmlToText, clampText, fetchPage,
  MAX_RESPONSE_BYTES, DEFAULT_MAX_CHARS,
} from "../../core/fetchpage.js";

const page = (body, { status = 200, type = "text/html", headers = {} } = {}) => ({
  ok: status >= 200 && status < 300,
  status,
  headers: { get: (k) => ({ "content-type": type, ...headers })[k.toLowerCase()] ?? null },
  text: async () => body,
});

describe("isAllowedUrl", () => {
  it("allows ordinary public pages", () => {
    expect(isAllowedUrl("https://pe.usps.com/businessmail101?ViewName=DeliveryAddress").ok).toBe(true);
    expect(isAllowedUrl("http://example.com/a/b").ok).toBe(true);
  });

  it("blocks every local and private address shape", () => {
    const blocked = [
      "http://localhost/", "https://localhost:39217/api/status", "http://foo.localhost/",
      "http://printer.local/", "http://127.0.0.1/", "http://127.9.9.9:8080/",
      "http://0.0.0.0/", "http://10.0.0.5/", "http://192.168.1.1/admin",
      "http://169.254.1.1/", "http://172.16.0.1/", "http://172.31.255.255/",
      "http://[::1]/", "http://[fe80::1]/",
    ];
    for (const url of blocked) {
      expect(isAllowedUrl(url).ok, url).toBe(false);
    }
  });

  it("does not over-block the public 172.x space", () => {
    expect(isAllowedUrl("http://172.15.0.1/").ok).toBe(true);
    expect(isAllowedUrl("http://172.32.0.1/").ok).toBe(true);
  });

  it("blocks non-web protocols and garbage", () => {
    expect(isAllowedUrl("ftp://example.com/file").ok).toBe(false);
    expect(isAllowedUrl("file:///etc/passwd").ok).toBe(false);
    expect(isAllowedUrl("javascript:alert(1)").ok).toBe(false);
    expect(isAllowedUrl("not a url").ok).toBe(false);
    expect(isAllowedUrl("").ok).toBe(false);
  });
});

describe("htmlToText", () => {
  it("keeps content and structure, drops chrome and code", () => {
    const html = `<html><head><title>USPS Addressing</title><script>evil()</script>
      <style>.x{}</style></head><body>
      <nav><a href="/">Home</a></nav>
      <h1>Delivery Address</h1>
      <p>Print or type the address in &quot;all caps&quot;.</p>
      <ul><li>Use two&#8211;letter state abbreviations</li><li>Omit punctuation</li></ul>
      <footer>© USPS</footer></body></html>`;
    const text = htmlToText(html);
    expect(text).toContain("USPS Addressing");
    expect(text).toContain("Delivery Address");
    expect(text).toContain('Print or type the address in "all caps".');
    expect(text).toContain("- Use two–letter state abbreviations");
    expect(text).not.toContain("evil");
    expect(text).not.toContain("Home");
    expect(text).not.toContain("©");
  });

  it("survives nested chrome elements", () => {
    const text = htmlToText("<nav><header><p>menu</p></header></nav><p>Real content</p>");
    expect(text).toBe("Real content");
  });

  it("decodes numeric entities and collapses whitespace floods", () => {
    const text = htmlToText("<p>A&#65;&#x42;   \n\n\n\n  B</p><p></p><p></p><p>C</p>");
    expect(text).toContain("AAB");
    expect(text).not.toMatch(/\n{3}/);
  });
});

describe("clampText", () => {
  it("passes short text through untouched", () => {
    expect(clampText("hello", 100)).toBe("hello");
  });

  it("truncates honestly, saying exactly how much was shown", () => {
    const clamped = clampText("x".repeat(10_000), 1_000);
    expect(clamped).toContain("x".repeat(1_000));
    expect(clamped).toContain("(page truncated — showing the first 1,000 of 10,000 characters)");
  });
});

describe("fetchPage", () => {
  it("fetches, extracts, and prefixes the final URL", async () => {
    const fetchImpl = vi.fn(async () => page("<title>T</title><p>Hello page</p>"));
    const result = await fetchPage("https://example.com/doc", { fetchImpl });
    expect(result).toContain("Contents of https://example.com/doc");
    expect(result).toContain("Hello page");
    expect(fetchImpl.mock.calls[0][1].redirect).toBe("manual");
  });

  it("unwraps DuckDuckGo redirect URLs before judging them", async () => {
    const fetchImpl = vi.fn(async () => page("<p>ok</p>"));
    const wrapped = "https://duckduckgo.com/l/?uddg=" + encodeURIComponent("https://example.com/real");
    await fetchPage(wrapped, { fetchImpl });
    expect(fetchImpl.mock.calls[0][0]).toBe("https://example.com/real");

    const evil = "https://duckduckgo.com/l/?uddg=" + encodeURIComponent("http://127.0.0.1/steal");
    expect(await fetchPage(evil, { fetchImpl })).toMatch(/cannot be opened/);
  });

  it("re-checks every redirect hop against the allow-list", async () => {
    const fetchImpl = vi.fn(async (url) => {
      if (url === "https://example.com/start") {
        return page("", { status: 302, headers: { location: "http://192.168.1.1/internal" } });
      }
      return page("<p>should never load</p>");
    });
    const result = await fetchPage("https://example.com/start", { fetchImpl });
    expect(result).toMatch(/redirects somewhere/);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it("follows a legitimate redirect and gives up after too many", async () => {
    let calls = 0;
    const hopper = vi.fn(async () => {
      calls += 1;
      return page("", { status: 301, headers: { location: `https://example.com/${calls}` } });
    });
    expect(await fetchPage("https://example.com/0", { fetchImpl: hopper })).toMatch(/redirected too many times/);

    const once = vi.fn(async (url) => url.endsWith("/final")
      ? page("<p>made it</p>")
      : page("", { status: 302, headers: { location: "https://example.com/final" } }));
    expect(await fetchPage("https://example.com/start", { fetchImpl: once })).toContain("made it");
  });

  it("returns readable failure text instead of throwing", async () => {
    expect(await fetchPage("http://localhost/x", { fetchImpl: vi.fn() })).toMatch(/cannot be opened/);
    expect(await fetchPage("", { fetchImpl: vi.fn() })).toBe("No URL was given.");
    expect(await fetchPage("https://e.com/", { fetchImpl: vi.fn(async () => page("nope", { status: 404 })) }))
      .toMatch(/HTTP 404/);
    expect(await fetchPage("https://e.com/", { fetchImpl: vi.fn(async () => { throw new Error("fetch failed"); }) }))
      .toMatch(/appears to be offline/);
    expect(await fetchPage("https://e.com/a.pdf", { fetchImpl: vi.fn(async () => page("%PDF", { type: "application/pdf" })) }))
      .toMatch(/not readable text/);
  });

  it("reads plain text without HTML processing", async () => {
    const result = await fetchPage("https://e.com/robots.txt", {
      fetchImpl: vi.fn(async () => page("User-agent: *\nDisallow:", { type: "text/plain" })),
    });
    expect(result).toContain("User-agent: *");
  });

  it("clamps giant pages to the char budget with the honest note", async () => {
    const huge = `<p>${"word ".repeat(5_000)}</p>`;
    const result = await fetchPage("https://e.com/big", {
      fetchImpl: vi.fn(async () => page(huge)),
      maxChars: 500,
    });
    expect(result).toContain("(page truncated — showing the first 500 of");
    expect(result.length).toBeLessThan(700);
  });

  it("has sane budget constants", () => {
    expect(MAX_RESPONSE_BYTES).toBe(2 * 1024 * 1024);
    expect(DEFAULT_MAX_CHARS).toBeGreaterThan(1_000);
  });
});
