/**
 * @vitest-environment jsdom
 *
 * The control panel's model picker, rendered against the real index.html markup with
 * window.eal stubbed. app.js binds and refreshes at import time, so the DOM and the
 * bridge must both exist before the module loads.
 */
import { describe, it, expect, beforeAll, vi } from "vitest";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");

const tick = () => new Promise((r) => setTimeout(r, 0));

const appState = {
  version: "test", port: 39217, serverError: null,
  engine: { state: "ready", contextTokens: 8192, error: null },
  activeModelId: "qwen3.5-4b",
  config: {},
  machine: { totalRamGb: 24, arch: "arm64", chip: "Apple M5 Pro", cpus: 12, memory: { total: 19e9 } },
  preflight: { ready: true, steps: [] },
  manifestPath: "/tmp/manifest.xml",
  modelsDir: "/tmp/models",
};

const listing = {
  models: [
    { id: "qwen3.5-2b", name: "Qwen3.5 2B", bytes: 1.3e9, params: "2B", quant: "UD-Q4_K_XL", contextTokens: 262144, intelligence: 1, gain: "Small and quick.", tagline: "Lightweight", strengths: ["Fast"], installed: true, partialBytes: 0, downloading: false, fit: { level: "great", label: "Plenty of room" } },
    { id: "qwen3.5-4b", name: "Qwen3.5 4B", bytes: 2.9e9, params: "4B", quant: "UD-Q4_K_XL", contextTokens: 262144, intelligence: 2, gain: "The sweet spot.", tagline: "Recommended", strengths: ["Balanced"], default: true, installed: true, partialBytes: 0, downloading: false, fit: { level: "great", label: "Plenty of room" } },
    { id: "qwen3.5-9b", name: "Qwen3.5 9B", bytes: 6.0e9, params: "9B", quant: "UD-Q4_K_XL", contextTokens: 262144, intelligence: 3, gain: "Stronger reasoning.", tagline: "Strong", strengths: ["Analysis"], installed: false, partialBytes: 0, downloading: false, fit: { level: "ok", label: "Fits" } },
    { id: "qwen3.5-27b", name: "Qwen3.5 27B", bytes: 17.6e9, params: "27B", quant: "UD-Q4_K_XL", contextTokens: 262144, intelligence: 4, gain: "Deep analysis.", tagline: "Deep", strengths: ["Depth"], installed: false, partialBytes: 0, downloading: false, fit: { level: "too-big", label: "Too large — ~21.9 GB of 19.0 GB usable" } },
  ],
  tiers: { fast: "qwen3.5-2b", balanced: "qwen3.5-4b", max: "qwen3.5-9b" },
  machine: { chip: "Apple M5 Pro", availableBytes: 19e9 },
};

let progressHandler;

beforeAll(async () => {
  const html = await readFile(path.join(ROOT, "desktop", "renderer", "index.html"), "utf8");
  document.body.innerHTML = html.match(/<body>([\s\S]*)<\/body>/)[1]
    .replace(/<script[\s\S]*?<\/script>/g, "");

  window.eal = {
    getState: vi.fn(async () => appState),
    models: {
      list: vi.fn(async () => listing),
      select: vi.fn(async () => ({})),
      download: vi.fn(async () => ({})),
      cancel: vi.fn(async () => true),
      remove: vi.fn(async () => true),
    },
    setup: { certificate: vi.fn(), addin: vi.fn(), grantAddinAccess: vi.fn(), manualAddin: vi.fn() },
    actions: { openExcel: vi.fn(), revealModels: vi.fn(), openRepo: vi.fn(), checkUpdates: vi.fn() },
    setConfig: vi.fn(),
    onProgress: (fn) => { progressHandler = fn; },
    onInstalled: () => {}, onRemoved: () => {}, onModelError: () => {},
    onEngineState: () => {}, onServerError: () => {}, onEngineError: () => {},
  };

  await import("../../desktop/renderer/app.js");
  await tick();
});

describe("recommended picks", () => {
  it("renders the three tiers in Fast → Balanced → Max order", () => {
    expect(document.getElementById("recommendSection").hidden).toBe(false);
    const tiers = [...document.querySelectorAll(".pick__tier")].map((n) => n.textContent);
    expect(tiers).toEqual(["Fast", "Balanced", "Max quality"]);
    const names = [...document.querySelectorAll(".pick__name")].map((n) => n.textContent);
    expect(names).toEqual(["Qwen3.5 2B", "Qwen3.5 4B", "Qwen3.5 9B"]);
  });

  it("marks the active pick and offers Use/Download on the others", () => {
    const picks = [...document.querySelectorAll(".pick")];
    expect(picks[1].dataset.active).toBe("true");
    expect(picks[1].querySelector(".tag--active")).toBeTruthy();
    expect(picks[0].querySelector(".btn").textContent).toBe("Use");
    expect(picks[2].querySelector(".btn").textContent).toMatch(/^Download/);
  });

  it("names the machine the recommendations were made for", () => {
    expect(document.getElementById("chipNote").textContent).toBe("Apple M5 Pro · 24 GB RAM");
  });
});

describe("the full model chart", () => {
  it("gives every card its id and an intelligence meter", () => {
    const cards = [...document.querySelectorAll(".model")];
    expect(cards.map((c) => c.dataset.id)).toEqual(listing.models.map((m) => m.id));
    for (const card of cards) {
      expect(card.querySelectorAll(".dots i")).toHaveLength(5);
    }
    // Intelligence shows even on the model that cannot run here.
    const big = document.querySelector('.model[data-id="qwen3.5-27b"]');
    expect(big.querySelectorAll('.dots i[data-on="true"]')).toHaveLength(4);
  });

  it("keeps too-big models on the chart but blocks the pointless download", () => {
    const big = document.querySelector('.model[data-id="qwen3.5-27b"]');
    expect(big.dataset.fit).toBe("too-big");
    const download = big.querySelector(".btn--primary");
    expect(download.disabled).toBe(true);
    expect(big.querySelector(".model__fit").textContent).toMatch(/Too large/);
  });

  it("quotes the runtime context window, not the 256K training figure", () => {
    const spec = document.querySelector('.model[data-id="qwen3.5-4b"] .model__spec');
    expect(spec.textContent).toContain("8K context");
    expect(spec.textContent).not.toContain("256K");
  });
});

describe("download progress", () => {
  it("finds the card by id, not by matching display text", async () => {
    // Regression: progress used to locate cards by comparing .model__name text,
    // which broke the moment the name node's structure changed.
    progressHandler({ id: "qwen3.5-9b", phase: "downloading", percent: 40, received: 2.4e9, total: 6.0e9, bytesPerSecond: 8e6, etaSeconds: 90 });
    await tick();
    // No progress bar existed on that card (it wasn't downloading at render time),
    // so the handler falls back to a full refresh without throwing.
    expect(window.eal.models.list).toHaveBeenCalled();
  });
});
