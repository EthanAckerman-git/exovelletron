/**
 * Filesystem layout, defaults, and persisted user preferences.
 *
 * Every path is derived from a single `home` value so tests can point the whole
 * app at a temp directory instead of touching the real user's Library.
 */
import { readFile, writeFile, mkdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

export const APP_NAME = "Excel AI Local";
export const ADDIN_ID = "b6f4a2c1-9d3e-4a7b-8c15-2e6f0a9d4b73";

/**
 * Fixed by default because the add-in manifest hardcodes the URL — Excel reads the
 * manifest at launch, so the port cannot drift between runs. If the port is taken we
 * pick another and rewrite the manifest rather than failing.
 */
export const DEFAULT_PORT = 39217;

export const DEFAULTS = Object.freeze({
  port: DEFAULT_PORT,
  modelId: "qwen3.5-4b",
  contextTokens: 8192,
  sheetContextTokens: 6000,
  temperature: 0.3,
  maxTokens: 1536,
  launchAtLogin: false,
  setupCompleted: false,
});

export function createPaths(home = os.homedir()) {
  const dataDir = path.join(home, "Library", "Application Support", APP_NAME);
  return {
    home,
    dataDir,
    certsDir: path.join(dataDir, "certs"),
    modelsDir: path.join(dataDir, "models"),
    configFile: path.join(dataDir, "config.json"),
    logFile: path.join(dataDir, "app.log"),
    // Excel for Mac reads sideloaded add-in manifests from inside its sandbox container.
    wefDir: path.join(home, "Library", "Containers", "com.microsoft.Excel", "Data", "Documents", "wef"),
  };
}

export const paths = createPaths();

/** Merge persisted prefs over defaults, dropping unknown and malformed keys. */
export function normalizeConfig(raw) {
  const out = { ...DEFAULTS };
  if (!raw || typeof raw !== "object") return out;
  for (const [key, fallback] of Object.entries(DEFAULTS)) {
    const value = raw[key];
    if (value === undefined || value === null) continue;
    if (typeof value !== typeof fallback) continue;
    if (typeof value === "number" && !Number.isFinite(value)) continue;
    out[key] = value;
  }
  if (out.port < 1024 || out.port > 65535) out.port = DEFAULT_PORT;
  if (out.temperature < 0 || out.temperature > 2) out.temperature = DEFAULTS.temperature;
  if (out.maxTokens < 64 || out.maxTokens > 32768) out.maxTokens = DEFAULTS.maxTokens;
  if (out.contextTokens < 2048 || out.contextTokens > 262144) out.contextTokens = DEFAULTS.contextTokens;
  return out;
}

export async function loadConfig(p = paths) {
  try {
    return normalizeConfig(JSON.parse(await readFile(p.configFile, "utf8")));
  } catch {
    return { ...DEFAULTS };
  }
}

export async function saveConfig(patch, p = paths) {
  const next = normalizeConfig({ ...(await loadConfig(p)), ...patch });
  await mkdir(p.dataDir, { recursive: true });
  await writeFile(p.configFile, JSON.stringify(next, null, 2), "utf8");
  return next;
}

export async function ensureDirs(p = paths) {
  await Promise.all([
    mkdir(p.dataDir, { recursive: true }),
    mkdir(p.certsDir, { recursive: true }),
    mkdir(p.modelsDir, { recursive: true }),
  ]);
}
