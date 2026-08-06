/**
 * The models offered in the app, with byte sizes and SHA-256 digests captured from the
 * Hugging Face LFS metadata so downloads can be verified rather than trusted.
 *
 * All three are Qwen3.5 in Unsloth's "UD" dynamic quantisation, which calibrates
 * per-layer bit widths and holds up noticeably better than a flat Q4_K_M at
 * effectively the same file size.
 */

export const DEFAULT_MODEL_ID = "qwen3.5-4b";

/** @typedef {{id:string,name:string,repo:string,file:string,bytes:number,sha256:string,
 *   params:string,quant:string,contextTokens:number,minRamGb:number,recommendedRamGb:number,
 *   tagline:string,strengths:string[],default?:boolean}} ModelSpec */

/** @type {ModelSpec[]} */
export const CATALOG = Object.freeze([
  {
    id: "qwen3.5-2b",
    name: "Qwen3.5 2B",
    repo: "unsloth/Qwen3.5-2B-GGUF",
    file: "Qwen3.5-2B-UD-Q4_K_XL.gguf",
    bytes: 1339752704,
    sha256: "0af96165ea615bea39a04118d63f0b6d35908aea850ee4a51aa6151d851b8b35",
    params: "2B",
    quant: "UD-Q4_K_XL",
    contextTokens: 262144,
    minRamGb: 4,
    recommendedRamGb: 8,
    tagline: "Lightweight",
    strengths: ["Runs on almost any Mac", "Fastest replies", "Smallest download"],
  },
  {
    id: "qwen3.5-4b",
    name: "Qwen3.5 4B",
    repo: "unsloth/Qwen3.5-4B-GGUF",
    file: "Qwen3.5-4B-UD-Q4_K_XL.gguf",
    bytes: 2912109728,
    sha256: "b252c5610a42ca82d20fe2a12813e9d069eed89292907e26c783eeb0bc961bc7",
    params: "4B",
    quant: "UD-Q4_K_XL",
    contextTokens: 262144,
    minRamGb: 8,
    recommendedRamGb: 16,
    tagline: "Recommended",
    strengths: ["Best balance of speed and accuracy", "Reliable formulas", "Strong structured output"],
    default: true,
  },
  {
    id: "qwen3.5-9b",
    name: "Qwen3.5 9B",
    repo: "unsloth/Qwen3.5-9B-GGUF",
    file: "Qwen3.5-9B-UD-Q4_K_XL.gguf",
    bytes: 5966095584,
    sha256: "6f5d30666c2d8ae16a306e616d95341dcf3cc46810df84d7e6f5a7d1e4c1b293",
    params: "9B",
    quant: "UD-Q4_K_XL",
    contextTokens: 262144,
    minRamGb: 16,
    recommendedRamGb: 24,
    tagline: "Maximum quality",
    strengths: ["Best at multi-step analysis", "Handles complex formulas", "Needs 16 GB+ of RAM"],
  },
]);

export const getModel = (id) => CATALOG.find((m) => m.id === id) ?? null;

export const downloadUrl = (model) =>
  `https://huggingface.co/${model.repo}/resolve/main/${encodeURIComponent(model.file)}?download=true`;

export function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB", "TB"];
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** i;
  return `${value >= 10 || i === 0 ? Math.round(value) : value.toFixed(1)} ${units[i]}`;
}

/**
 * Fit assessment for the machine we are running on, used to steer the picker.
 * @param {ModelSpec} model
 * @param {number} totalRamGb
 */
export function fitForMachine(model, totalRamGb) {
  if (totalRamGb >= model.recommendedRamGb) return { level: "great", label: "Runs great on this Mac" };
  if (totalRamGb >= model.minRamGb) return { level: "ok", label: "Runs on this Mac" };
  return { level: "tight", label: `Needs ${model.minRamGb} GB of RAM` };
}
