import { describe, it, expect } from "vitest";
import path from "node:path";
import { normalizeConfig, DEFAULTS, createPaths } from "../../core/config.js";
import { buildManifest } from "../../core/setup/manifest.js";
import { resolveWithinRoot, contentTypeFor } from "../../core/server/static.js";
import { tokensMatch, isAllowedOrigin, authorizeApiRequest, mintToken, securityHeaders, TOKEN_HEADER } from "../../core/server/security.js";
import { CATALOG, getModel, downloadUrl, formatBytes, fitForMachine, DEFAULT_MODEL_ID } from "../../core/models/catalog.js";
import { renderSheetContext, buildUserTurn, estimateTokens } from "../../core/llm/prompt.js";
import { isBenignBackendNoise } from "../../core/llm/engine.js";

describe("config", () => {
  it("falls back to defaults for junk input", () => {
    expect(normalizeConfig(null)).toEqual(DEFAULTS);
    expect(normalizeConfig("nope")).toEqual(DEFAULTS);
  });

  it("ignores keys of the wrong type and unknown keys", () => {
    const c = normalizeConfig({ port: "8080", nonsense: 1, temperature: 0.9 });
    expect(c.port).toBe(DEFAULTS.port);
    expect(c).not.toHaveProperty("nonsense");
    expect(c.temperature).toBe(0.9);
  });

  it("clamps out-of-range values", () => {
    expect(normalizeConfig({ port: 80 }).port).toBe(DEFAULTS.port);
    expect(normalizeConfig({ port: 99999 }).port).toBe(DEFAULTS.port);
    expect(normalizeConfig({ temperature: 5 }).temperature).toBe(DEFAULTS.temperature);
    expect(normalizeConfig({ maxTokens: 1 }).maxTokens).toBe(DEFAULTS.maxTokens);
    expect(normalizeConfig({ contextTokens: 10 }).contextTokens).toBe(DEFAULTS.contextTokens);
  });

  it("rejects NaN", () => {
    expect(normalizeConfig({ temperature: NaN }).temperature).toBe(DEFAULTS.temperature);
  });

  it("derives every path from the supplied home directory", () => {
    const p = createPaths("/tmp/fakehome");
    expect(p.dataDir).toBe("/tmp/fakehome/Library/Application Support/Excel AI Local");
    expect(p.wefDir).toContain("com.microsoft.Excel");
    for (const key of ["certsDir", "modelsDir", "configFile", "logFile"]) {
      expect(p[key].startsWith("/tmp/fakehome")).toBe(true);
    }
  });
});

describe("manifest", () => {
  it("embeds the port in every URL", () => {
    const xml = buildManifest(39217);
    expect(xml).toContain("https://localhost:39217/taskpane.html");
    expect(xml).toContain("https://localhost:39217/commands.html");
    expect(xml).not.toMatch(/https:\/\/localhost:(?!39217)\d+/);
  });

  it("declares the Workbook host and write permission", () => {
    const xml = buildManifest(39217);
    expect(xml).toContain('<Host Name="Workbook"/>');
    expect(xml).toContain("<Permissions>ReadWriteDocument</Permissions>");
  });

  it("references only icons the server actually serves", () => {
    const xml = buildManifest(39217);
    for (const size of [16, 32, 80, 64, 128]) {
      if (xml.includes(`icon-${size}.png`)) expect(xml).toContain(`/assets/icon-${size}.png`);
    }
    expect(xml).toContain("icon-16.png");
    expect(xml).toContain("icon-32.png");
    expect(xml).toContain("icon-80.png");
  });

  it("is well-formed enough to have balanced root tags", () => {
    const xml = buildManifest(39217);
    expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8"?>')).toBe(true);
    expect((xml.match(/<OfficeApp/g) || []).length).toBe(1);
    expect((xml.match(/<\/OfficeApp>/g) || []).length).toBe(1);
    expect((xml.match(/<VersionOverrides/g) || []).length).toBe(1);
  });

  it.each([0, 80, 70000, 1.5, NaN, "39217"])("rejects invalid port %o", (port) => {
    expect(() => buildManifest(port)).toThrow(/Invalid port/);
  });
});

describe("static file safety", () => {
  const root = "/srv/dist";

  it("resolves normal paths inside the root", () => {
    expect(resolveWithinRoot(root, "/app/main.js")).toBe(path.join(root, "app/main.js"));
  });

  it.each([
    "/../../../etc/passwd",
    "/..%2f..%2fetc/passwd",
    "/app/../../secrets.txt",
    "/%2e%2e/%2e%2e/etc/passwd",
  ])("refuses traversal via %o", (attack) => {
    const resolved = resolveWithinRoot(root, attack);
    expect(resolved === null || resolved.startsWith(root)).toBe(true);
  });

  it("refuses null bytes and bad encoding", () => {
    expect(resolveWithinRoot(root, "/a\0b")).toBeNull();
    expect(resolveWithinRoot(root, "/%zz")).toBeNull();
  });

  it("maps content types", () => {
    expect(contentTypeFor("a.html")).toMatch(/text\/html/);
    expect(contentTypeFor("a.js")).toMatch(/javascript/);
    expect(contentTypeFor("a.png")).toBe("image/png");
    expect(contentTypeFor("a.unknown")).toBe("application/octet-stream");
  });
});

describe("api security", () => {
  it("compares tokens without leaking length", () => {
    const t = mintToken();
    expect(tokensMatch(t, t)).toBe(true);
    expect(tokensMatch(t, t.slice(0, -1))).toBe(false);
    expect(tokensMatch(t, undefined)).toBe(false);
    expect(tokensMatch(undefined, undefined)).toBe(false);
  });

  it("mints distinct high-entropy tokens", () => {
    const a = mintToken();
    expect(a).not.toBe(mintToken());
    expect(a.length).toBeGreaterThanOrEqual(40);
  });

  it("allows our own origin and no origin, rejects others", () => {
    expect(isAllowedOrigin(undefined, 39217)).toBe(true);
    expect(isAllowedOrigin("https://localhost:39217", 39217)).toBe(true);
    expect(isAllowedOrigin("https://evil.example", 39217)).toBe(false);
    expect(isAllowedOrigin("https://localhost:1234", 39217)).toBe(false);
  });

  it("authorizes only with a matching token", () => {
    const token = mintToken();
    const ok = authorizeApiRequest({ headers: { [TOKEN_HEADER]: token } }, { token, port: 39217 });
    expect(ok).toEqual({ ok: true });

    expect(authorizeApiRequest({ headers: {} }, { token, port: 39217 }))
      .toMatchObject({ ok: false, status: 401 });
    expect(authorizeApiRequest({ headers: { [TOKEN_HEADER]: "nope" } }, { token, port: 39217 }))
      .toMatchObject({ ok: false, status: 401 });
    expect(authorizeApiRequest({ headers: { origin: "https://evil.example", [TOKEN_HEADER]: token } }, { token, port: 39217 }))
      .toMatchObject({ ok: false, status: 403 });
  });

  it("pins connect-src to our own origin so the pane cannot phone home", () => {
    const csp = securityHeaders(39217)["Content-Security-Policy"];
    expect(csp).toContain("default-src 'self'");

    const directives = Object.fromEntries(
      csp.split(";").map((d) => {
        const [name, ...rest] = d.trim().split(/\s+/);
        return [name, rest];
      }),
    );
    // No wildcard may appear anywhere the pane could fetch or load code from.
    // frame-ancestors is exempt: Office legitimately hosts the pane from its own domains.
    for (const directive of ["default-src", "script-src", "connect-src", "img-src", "style-src", "font-src"]) {
      expect(directives[directive].join(" ")).not.toContain("*");
    }
    expect(directives["connect-src"]).toEqual([
      "'self'", "https://localhost:39217", "https://127.0.0.1:39217",
    ]);
  });
});

describe("model catalog", () => {
  it("has a single default that exists", () => {
    const defaults = CATALOG.filter((m) => m.default);
    expect(defaults).toHaveLength(1);
    expect(defaults[0].id).toBe(DEFAULT_MODEL_ID);
    expect(getModel(DEFAULT_MODEL_ID)).toBeTruthy();
  });

  it("every entry is fully specified", () => {
    for (const m of CATALOG) {
      expect(m.sha256).toMatch(/^[0-9a-f]{64}$/);
      expect(m.bytes).toBeGreaterThan(1e8);
      expect(m.file.endsWith(".gguf")).toBe(true);
      expect(m.repo).toMatch(/^[\w.-]+\/[\w.-]+$/);
      expect(m.strengths.length).toBeGreaterThan(0);
      expect(m.recommendedRamGb).toBeGreaterThanOrEqual(m.minRamGb);
    }
  });

  it("has unique ids and filenames", () => {
    expect(new Set(CATALOG.map((m) => m.id)).size).toBe(CATALOG.length);
    expect(new Set(CATALOG.map((m) => m.file)).size).toBe(CATALOG.length);
  });

  it("builds a huggingface resolve url", () => {
    expect(downloadUrl(getModel("qwen3.5-4b")))
      .toBe("https://huggingface.co/unsloth/Qwen3.5-4B-GGUF/resolve/main/Qwen3.5-4B-UD-Q4_K_XL.gguf?download=true");
  });

  it("formats byte sizes", () => {
    expect(formatBytes(0)).toBe("0 B");
    expect(formatBytes(999)).toBe("999 B");
    expect(formatBytes(2_912_109_728)).toBe("2.7 GB");
    expect(formatBytes(-5)).toBe("0 B");
  });

  it("grades fit against available RAM", () => {
    const m = getModel("qwen3.5-9b");
    expect(fitForMachine(m, 32).level).toBe("great");
    expect(fitForMachine(m, 16).level).toBe("ok");
    expect(fitForMachine(m, 8).level).toBe("tight");
  });
});

describe("prompt rendering", () => {
  const ctx = {
    workbookName: "Q3.xlsx",
    sheetName: "Sales",
    sheetNames: ["Sales", "Notes"],
    usedRange: { address: "A1:E501", rowCount: 501, columnCount: 5 },
    selection: { address: "E2:E501", rowCount: 500, columnCount: 1, isSingleCell: false },
    headers: ["Region", "Rep", "Units"],
    columnLetters: ["A", "B", "C"],
    sample: { startRow: 2, columnLetters: ["A", "B", "C"], rows: [["North", "Ada", 120], ["South", "Ben", 80]] },
  };

  it("includes the structural facts the model needs", () => {
    const out = renderSheetContext(ctx, 6000);
    expect(out).toContain("Active sheet: Sales");
    expect(out).toContain("501 rows x 5 columns");
    expect(out).toContain("Current selection: E2:E501");
    expect(out).toContain("A=Region");
  });

  it("says the rows are a sample, not the whole sheet", () => {
    expect(renderSheetContext(ctx, 6000)).toMatch(/NOT the full sheet/);
  });

  it("respects a tight token budget", () => {
    const big = {
      ...ctx,
      sample: {
        startRow: 2,
        columnLetters: ["A", "B", "C"],
        rows: Array.from({ length: 500 }, (_, i) => [`row${i}`, "x".repeat(50), i]),
      },
    };
    const budget = 500;
    expect(estimateTokens(renderSheetContext(big, budget))).toBeLessThanOrEqual(budget * 1.4);
  });

  it("handles a missing worksheet gracefully", () => {
    expect(renderSheetContext(null)).toMatch(/No workbook data/);
  });

  it("wraps the question with worksheet context", () => {
    const turn = buildUserTurn("What is this?", ctx, 6000);
    expect(turn).toContain("<worksheet>");
    expect(turn.trimEnd().endsWith("What is this?")).toBe(true);
  });

  it("passes the question through when there is no context", () => {
    expect(buildUserTurn("Hello", null, 6000)).toBe("Hello");
  });
});

describe("backend log filtering", () => {
  it("suppresses the benign Metal tensor-API fallback", () => {
    expect(isBenignBackendNoise("ggml_metal_library_init_from_source: error compiling source")).toBe(true);
    expect(isBenignBackendNoise("ggml_metal_device_init: - the tensor API is not supported in this environment - disabling")).toBe(true);
  });

  it("lets real errors through", () => {
    expect(isBenignBackendNoise("failed to allocate buffer")).toBe(false);
    expect(isBenignBackendNoise(undefined)).toBe(false);
  });
});
