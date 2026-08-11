/**
 * Turn a decorated model list (catalog entries + fit labels) into the three
 * recommendations the picker shows: Fast, Balanced, and Max quality.
 *
 * The tiers form a ladder, not a ranking of absolutes: Max is the smartest model this
 * machine can actually hold, Balanced the smartest that fits comfortably below it, and
 * Fast the smartest that still has plenty of room to spare. On a small Mac the ladder
 * loses rungs from the bottom up — a tier that would have to lie returns null instead.
 *
 * Pure functions on already-decorated data, so every machine profile is unit-testable
 * without hardware.
 */

const FITS_AT_ALL = new Set(["great", "ok", "tight"]);
const FITS_COMFORTABLY = new Set(["great", "ok"]);

const byIntelligenceThenSize = (a, b) =>
  (b.intelligence ?? 0) - (a.intelligence ?? 0) || b.bytes - a.bytes;

/**
 * @param {Array<{id:string,bytes:number,intelligence?:number,fit?:{level:string}}>} models
 * @returns {{fast: string|null, balanced: string|null, max: string|null}}
 */
export function recommendTiers(models) {
  const eligible = (models ?? [])
    .filter((m) => FITS_AT_ALL.has(m.fit?.level))
    .sort(byIntelligenceThenSize);

  const max = eligible[0] ?? null;

  const balanced = eligible.find(
    (m) => max && m.id !== max.id && FITS_COMFORTABLY.has(m.fit.level) && m.bytes < max.bytes,
  ) ?? null;

  const ceiling = balanced ?? max;
  const fast = eligible.find(
    (m) => ceiling && m.id !== ceiling.id && m.fit.level === "great" && m.bytes < ceiling.bytes,
  ) ?? null;

  return { fast: fast?.id ?? null, balanced: balanced?.id ?? null, max: max?.id ?? null };
}

/** Human labels for the tiers, in display order. */
export const TIER_LABELS = Object.freeze([
  { key: "fast", label: "Fast", note: "Quick replies, light footprint" },
  { key: "balanced", label: "Balanced", note: "Quality and speed in proportion" },
  { key: "max", label: "Max quality", note: "The smartest this Mac can run" },
]);

/**
 * Shape the list for rendering: the recommended cards (in Fast → Max order, empty
 * tiers omitted) and the complete chart of every model.
 */
export function groupForDisplay(models, tiers) {
  const byId = new Map((models ?? []).map((m) => [m.id, m]));
  const recommended = TIER_LABELS
    .map(({ key, label, note }) => {
      const model = byId.get(tiers?.[key] ?? "");
      return model ? { tier: key, label, note, model } : null;
    })
    .filter(Boolean);
  return { recommended, all: models ?? [] };
}
