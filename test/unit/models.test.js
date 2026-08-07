import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { mkdtemp, rm, mkdir, writeFile, readdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { ModelStore } from "../../core/models/store.js";
import { getModel, DEFAULT_MODEL_ID } from "../../core/models/catalog.js";
import { isInstalled, modelFilePath, partFilePath, deleteModel, downloadModel } from "../../core/models/downloader.js";
import { createPaths, migrateLegacyDataDir, LEGACY_APP_NAMES, APP_NAME } from "../../core/config.js";

let dir;
const model = getModel(DEFAULT_MODEL_ID);

beforeEach(async () => {
  dir = await mkdtemp(path.join(tmpdir(), "eal-models-"));
});
afterEach(async () => {
  await rm(dir, { recursive: true, force: true });
});

/** Stand in for a completed download without moving gigabytes around. */
async function fakeInstall(modelsDir, spec = model) {
  await mkdir(modelsDir, { recursive: true });
  await writeFile(modelFilePath(modelsDir, spec), Buffer.alloc(spec.bytes > 1e6 ? 8 : spec.bytes));
}

describe("model removal and re-download", () => {
  it("reports a model as installed only at the exact expected size", async () => {
    const modelsDir = path.join(dir, "models");
    await fakeInstall(modelsDir);
    // A truncated file is not an install; it would crash the engine on load.
    expect(await isInstalled(modelsDir, model)).toBe(false);
  });

  it("deletes both the model and any partial download", async () => {
    const modelsDir = path.join(dir, "models");
    await mkdir(modelsDir, { recursive: true });
    await writeFile(modelFilePath(modelsDir, model), Buffer.alloc(16));
    await writeFile(partFilePath(modelsDir, model), Buffer.alloc(16));

    await deleteModel(modelsDir, model);

    expect(existsSync(modelFilePath(modelsDir, model))).toBe(false);
    expect(existsSync(partFilePath(modelsDir, model))).toBe(false);
    expect(await readdir(modelsDir)).toEqual([]);
  });

  it("leaves the folder ready for a clean re-download after removal", async () => {
    const modelsDir = path.join(dir, "models");
    const store = new ModelStore(modelsDir);
    await mkdir(modelsDir, { recursive: true });
    await writeFile(modelFilePath(modelsDir, model), Buffer.alloc(16));

    await store.remove(model.id);
    expect(await store.isInstalled(model.id)).toBe(false);

    const listed = (await store.list()).find((m) => m.id === model.id);
    expect(listed.installed).toBe(false);
    expect(listed.partialBytes).toBe(0);
    expect(listed.downloading).toBe(false);
  });

  it("refuses to remove a model that is not in the catalog", async () => {
    const store = new ModelStore(path.join(dir, "models"));
    await expect(store.remove("not-a-model")).rejects.toThrow(/Unknown model/);
  });

  it("re-downloads from scratch when no partial file remains", async () => {
    const modelsDir = path.join(dir, "models");
    const tiny = { ...model, bytes: 4, sha256: "9f86d081884c7d659a2feaa0c55ad015a3bf4f1b2b0b822cd15d6c15b0f00a08" };

    let rangeHeaderSeen = null;
    const fetchImpl = async (_url, options) => {
      rangeHeaderSeen = options.headers.Range ?? null;
      return {
        ok: true,
        status: 200,
        body: (async function* () { yield Buffer.from("test"); })(),
      };
    };

    const file = await downloadModel({ modelsDir, model: tiny, fetchImpl });
    expect(existsSync(file)).toBe(true);
    // A fresh download must not ask the server to resume from an offset.
    expect(rangeHeaderSeen).toBeNull();
  });

  it("resumes from a partial file rather than starting over", async () => {
    const modelsDir = path.join(dir, "models");
    const tiny = { ...model, bytes: 4, sha256: "9f86d081884c7d659a2feaa0c55ad015a3bf4f1b2b0b822cd15d6c15b0f00a08" };
    await mkdir(modelsDir, { recursive: true });
    await writeFile(partFilePath(modelsDir, tiny), Buffer.from("te"));

    let rangeHeaderSeen = null;
    const fetchImpl = async (_url, options) => {
      rangeHeaderSeen = options.headers.Range ?? null;
      return {
        ok: true,
        status: 206,
        body: (async function* () { yield Buffer.from("st"); })(),
      };
    };

    await downloadModel({ modelsDir, model: tiny, fetchImpl });
    expect(rangeHeaderSeen).toBe("bytes=2-");
  });

  it("discards a download whose digest does not match", async () => {
    const modelsDir = path.join(dir, "models");
    const tiny = { ...model, bytes: 4, sha256: "0".repeat(64) };
    const fetchImpl = async () => ({
      ok: true,
      status: 200,
      body: (async function* () { yield Buffer.from("test"); })(),
    });

    await expect(downloadModel({ modelsDir, model: tiny, fetchImpl })).rejects.toThrow(/integrity check/);
    expect(existsSync(modelFilePath(modelsDir, tiny))).toBe(false);
    expect(existsSync(partFilePath(modelsDir, tiny))).toBe(false);
  });
});

describe("data directory migration", () => {
  // Regression: renaming the app pointed Application Support at a new folder, which
  // orphaned a 2.9 GB model download and every saved conversation.
  it("moves models and conversations from a previous app name", async () => {
    const paths = createPaths(dir);
    const legacy = path.join(path.dirname(paths.dataDir), LEGACY_APP_NAMES[0]);
    await mkdir(path.join(legacy, "models"), { recursive: true });
    await mkdir(path.join(legacy, "conversations"), { recursive: true });
    await writeFile(path.join(legacy, "models", "weights.gguf"), "x");
    await writeFile(path.join(legacy, "config.json"), "{}");

    const moved = await migrateLegacyDataDir(paths);

    expect(moved).toEqual(expect.arrayContaining(["models", "conversations", "config.json"]));
    expect(existsSync(path.join(paths.modelsDir, "weights.gguf"))).toBe(true);
    expect(existsSync(paths.configFile)).toBe(true);
  });

  // Electron creates its userData folder at the new path before any of our code runs,
  // so the destination always exists by migration time. An all-or-nothing rename would
  // see that and skip, silently orphaning a 2.9 GB download.
  it("still migrates when the destination directory already exists", async () => {
    const paths = createPaths(dir);
    await mkdir(paths.dataDir, { recursive: true });
    await writeFile(path.join(paths.dataDir, "Cookies"), "electron");

    const legacy = path.join(path.dirname(paths.dataDir), LEGACY_APP_NAMES[0]);
    await mkdir(path.join(legacy, "models"), { recursive: true });
    await writeFile(path.join(legacy, "models", "weights.gguf"), "x");

    expect(await migrateLegacyDataDir(paths)).toContain("models");
    expect(existsSync(path.join(paths.modelsDir, "weights.gguf"))).toBe(true);
  });

  // A previous launch creates models/ before anything is downloaded into it; treating
  // that empty folder as real data would strand the legacy copy forever.
  it("migrates past an empty destination folder", async () => {
    const paths = createPaths(dir);
    await mkdir(paths.modelsDir, { recursive: true });

    const legacy = path.join(path.dirname(paths.dataDir), LEGACY_APP_NAMES[0]);
    await mkdir(path.join(legacy, "models"), { recursive: true });
    await writeFile(path.join(legacy, "models", "weights.gguf"), "x");

    expect(await migrateLegacyDataDir(paths)).toContain("models");
    expect(existsSync(path.join(paths.modelsDir, "weights.gguf"))).toBe(true);
  });

  it("never overwrites data that already exists under the new name", async () => {
    const paths = createPaths(dir);
    await mkdir(paths.modelsDir, { recursive: true });
    await writeFile(path.join(paths.modelsDir, "current.gguf"), "keep");

    const legacy = path.join(path.dirname(paths.dataDir), LEGACY_APP_NAMES[0]);
    await mkdir(path.join(legacy, "models"), { recursive: true });
    await writeFile(path.join(legacy, "models", "old.gguf"), "old");

    expect(await migrateLegacyDataDir(paths)).toEqual([]);
    expect(existsSync(path.join(paths.modelsDir, "current.gguf"))).toBe(true);
    expect(existsSync(path.join(paths.modelsDir, "old.gguf"))).toBe(false);
  });

  it("does nothing when there is no legacy folder", async () => {
    expect(await migrateLegacyDataDir(createPaths(dir))).toEqual([]);
  });

  it("does not list the current name as a legacy one", () => {
    expect(LEGACY_APP_NAMES).not.toContain(APP_NAME);
  });
});
