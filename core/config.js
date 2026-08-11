/**
 * Filesystem layout, defaults, and persisted user preferences.
 *
 * Every path is derived from a single `home` value so tests can point the whole
 * app at a temp directory instead of touching the real user's Library.
 */
import { readFile, writeFile, mkdir, rename, rm, readdir, stat } from "node:fs/promises";
import { existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { DEFAULT_MODEL_ID } from "./models/catalog.js";

export const APP_NAME = "Exovelletron";
export const ADDIN_ID = "b6f4a2c1-9d3e-4a7b-8c15-2e6f0a9d4b73";

/**
 * Fixed by default because the add-in manifest hardcodes the URL — Excel reads the
 * manifest at launch, so the port cannot drift between runs. If the port is taken we
 * pick another and rewrite the manifest rather than failing.
 */
export const DEFAULT_PORT = 39217;

export const DEFAULTS = Object.freeze({
  port: DEFAULT_PORT,
  modelId: DEFAULT_MODEL_ID,
  contextTokens: 8192,
  sheetContextTokens: 6000,
  temperature: 0.3,
  maxTokens: 1536,
  launchAtLogin: false,
  setupCompleted: false,
  // Off by default: the app's promise is that nothing leaves this Mac. When the user
  // turns it on, only search queries are transmitted — see core/search.js.
  webSearch: false,
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
  await migrateLegacyDataDir(p);
  await Promise.all([
    mkdir(p.dataDir, { recursive: true }),
    mkdir(p.certsDir, { recursive: true }),
    mkdir(p.modelsDir, { recursive: true }),
  ]);
}

/** Application Support folders used by earlier names of this app. */
export const LEGACY_APP_NAMES = ["Excel AI Local"];

/** True for a directory containing nothing; a file is never "empty" for this purpose. */
async function isEmptyDir(target) {
  try {
    if (!(await stat(target)).isDirectory()) return false;
    return (await readdir(target)).length === 0;
  } catch {
    return false;
  }
}

/** The things worth carrying across a rename. Everything else is regenerated. */
const MIGRATABLE = ["models", "conversations", "config.json"];

/**
 * Move data written under a previous app name into the current one.
 *
 * Without this, renaming the app silently orphans a multi-gigabyte model download and
 * every saved conversation, and the user is asked to fetch it all again for no reason.
 *
 * Migrating item by item rather than renaming the whole folder is deliberate: Electron
 * creates its own userData directory at the new path before any of our code runs, so the
 * destination always exists by the time we get here and a directory rename would fail.
 * Nothing is overwritten — an item already present at the destination wins.
 *
 * @returns {Promise<string[]>} the items that were moved
 */
export async function migrateLegacyDataDir(p = paths) {
  const supportDir = path.dirname(p.dataDir);
  const moved = [];

  for (const legacy of LEGACY_APP_NAMES) {
    const from = path.join(supportDir, legacy);
    if (from === p.dataDir || !existsSync(from)) continue;

    await mkdir(p.dataDir, { recursive: true });
    for (const item of MIGRATABLE) {
      const source = path.join(from, item);
      const target = path.join(p.dataDir, item);
      if (!existsSync(source)) continue;

      // An empty destination directory is not real data — a previous launch will have
      // created `models/` before anything was downloaded into it. Treating that as
      // "already migrated" would strand the legacy copy forever.
      if (existsSync(target)) {
        if (!(await isEmptyDir(target))) continue;
        await rm(target, { recursive: true, force: true });
      }

      try {
        await rename(source, target);
        moved.push(item);
      } catch {
        // Cross-device or permission failure just means that item starts fresh.
      }
    }
  }
  return moved;
}
