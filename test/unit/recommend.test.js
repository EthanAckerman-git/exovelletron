import { describe, it, expect, vi } from "vitest";
import { CATALOG, fitForMachine } from "../../core/models/catalog.js";
import { recommendTiers, groupForDisplay, TIER_LABELS } from "../../core/models/recommend.js";
import { parseChipString, detectChip } from "../../core/models/machine.js";

/**
 * Decorate the real catalog the way ModelStore.list() does, for a Mac with the given
 * installed RAM. Apple Silicon exposes roughly 74% of installed RAM as the GPU's
 * usable working set (measured on this codebase's dev machine: ~19 GB of 24 GB),
 * so the profiles use that ratio.
 */
const macWith = (ramGb) => {
  const availableBytes = ramGb * 1024 ** 3 * 0.737;
  return CATALOG.map((m) => ({ ...m, fit: fitForMachine(m, { availableBytes, contextTokens: 8192 }) }));
};

describe("recommendTiers", () => {
  // The ladder every profile expects, pinned explicitly so a catalog or estimate
  // change that silently reshuffles recommendations fails loudly here.
  const expectations = [
    [8,  { fast: null,         balanced: "qwen3.5-2b", max: "qwen3.5-4b" }],
    [16, { fast: "qwen3.5-2b", balanced: "qwen3.5-4b", max: "qwen3.5-9b" }],
    [24, { fast: "qwen3.5-2b", balanced: "qwen3.5-4b", max: "qwen3.5-9b" }],
    [32, { fast: "qwen3.5-4b", balanced: "qwen3.5-9b", max: "qwen3.5-27b" }],
    [48, { fast: "qwen3.5-9b", balanced: "qwen3.5-27b", max: "qwen3.5-35b-a3b" }],
    [64, { fast: "qwen3.5-9b", balanced: "qwen3.5-27b", max: "qwen3.5-35b-a3b" }],
  ];

  for (const [ramGb, expected] of expectations) {
    it(`recommends the right ladder on a ${ramGb} GB Mac`, () => {
      expect(recommendTiers(macWith(ramGb))).toEqual(expected);
    });
  }

  it("never recommends a model that does not fit", () => {
    for (const ramGb of [8, 16, 24, 32, 48, 64, 128]) {
      const models = macWith(ramGb);
      const tiers = recommendTiers(models);
      for (const id of Object.values(tiers)) {
        if (!id) continue;
        const fit = models.find((m) => m.id === id).fit.level;
        expect(["great", "ok", "tight"]).toContain(fit);
      }
    }
  });

  it("never puts the same model in two tiers", () => {
    for (const ramGb of [4, 8, 16, 24, 32, 48, 64]) {
      const picked = Object.values(recommendTiers(macWith(ramGb))).filter(Boolean);
      expect(new Set(picked).size).toBe(picked.length);
    }
  });

  it("returns all nulls rather than lying when nothing fits", () => {
    const none = CATALOG.map((m) => ({ ...m, fit: { level: "too-big" } }));
    expect(recommendTiers(none)).toEqual({ fast: null, balanced: null, max: null });
    expect(recommendTiers([])).toEqual({ fast: null, balanced: null, max: null });
  });

  it("keeps Balanced comfortable — a tight fit may only be Max", () => {
    for (const ramGb of [8, 16, 24, 32, 48, 64]) {
      const models = macWith(ramGb);
      const tiers = recommendTiers(models);
      if (!tiers.balanced) continue;
      const level = models.find((m) => m.id === tiers.balanced).fit.level;
      expect(["great", "ok"]).toContain(level);
    }
  });
});

describe("groupForDisplay", () => {
  it("orders recommended cards Fast → Balanced → Max and keeps the full chart", () => {
    const models = macWith(48);
    const { recommended, all } = groupForDisplay(models, recommendTiers(models));
    expect(recommended.map((r) => r.tier)).toEqual(["fast", "balanced", "max"]);
    expect(recommended.map((r) => r.model.id))
      .toEqual(["qwen3.5-9b", "qwen3.5-27b", "qwen3.5-35b-a3b"]);
    // The chart below always shows every model, even ones that are too big.
    expect(all.map((m) => m.id)).toEqual(CATALOG.map((m) => m.id));
  });

  it("omits empty tiers instead of rendering blanks", () => {
    const models = macWith(8);
    const { recommended } = groupForDisplay(models, recommendTiers(models));
    expect(recommended.map((r) => r.tier)).toEqual(["balanced", "max"]);
  });

  it("labels every tier for display", () => {
    expect(TIER_LABELS.map((t) => t.key)).toEqual(["fast", "balanced", "max"]);
    for (const t of TIER_LABELS) expect(t.label.length).toBeGreaterThan(2);
  });
});

describe("chip detection", () => {
  it("recognises Apple Silicon marketing names", () => {
    expect(parseChipString("Apple M1")).toEqual({ chip: "Apple M1", apple: true });
    expect(parseChipString("Apple M2 Pro")).toEqual({ chip: "Apple M2 Pro", apple: true });
    expect(parseChipString("Apple M3 Max\n")).toEqual({ chip: "Apple M3 Max", apple: true });
    expect(parseChipString("Apple M5 Pro")).toEqual({ chip: "Apple M5 Pro", apple: true });
  });

  it("collapses Intel brand strings to just Intel", () => {
    expect(parseChipString("Intel(R) Core(TM) i7-9750H CPU @ 2.60GHz"))
      .toEqual({ chip: "Intel", apple: false });
  });

  it("returns null for garbage without guessing", () => {
    expect(parseChipString("")).toEqual({ chip: null, apple: false });
    expect(parseChipString(null)).toEqual({ chip: null, apple: false });
    expect(parseChipString("Snapdragon X Elite")).toEqual({ chip: null, apple: false });
  });

  it("never throws when sysctl is unavailable", async () => {
    const result = await detectChip({ execImpl: vi.fn(async () => { throw new Error("ENOENT"); }) });
    expect(result.chip).toBeNull();
  });

  it("reads the brand string through sysctl", async () => {
    const execImpl = vi.fn(async () => ({ stdout: "Apple M5 Pro\n" }));
    expect(await detectChip({ execImpl })).toEqual({ chip: "Apple M5 Pro", apple: true });
    expect(execImpl).toHaveBeenCalledWith("sysctl", ["-n", "machdep.cpu.brand_string"]);
  });
});
