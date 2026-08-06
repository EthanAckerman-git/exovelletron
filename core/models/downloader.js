/**
 * Resumable, verified model downloads.
 *
 * A 3 GB download over a hotel wifi will get interrupted, so this writes to a `.part`
 * file, resumes with a Range request, verifies SHA-256 against the digest baked into
 * the catalog, and only then moves the file into place. A partial or corrupted file is
 * never presented to the engine as installed.
 */
import { createWriteStream, createReadStream } from "node:fs";
import { stat, rename, rm, mkdir } from "node:fs/promises";
import { createHash } from "node:crypto";
import { pipeline } from "node:stream/promises";
import path from "node:path";
import { downloadUrl } from "./catalog.js";

export const modelFilePath = (modelsDir, model) => path.join(modelsDir, model.file);
export const partFilePath = (modelsDir, model) => path.join(modelsDir, `${model.file}.part`);

const sizeOf = async (file) => {
  try {
    return (await stat(file)).size;
  } catch {
    return 0;
  }
};

/** Streaming SHA-256 of a file on disk. */
export async function hashFile(file, onProgress) {
  const hash = createHash("sha256");
  let read = 0;
  const stream = createReadStream(file, { highWaterMark: 4 * 1024 * 1024 });
  for await (const chunk of stream) {
    hash.update(chunk);
    read += chunk.length;
    onProgress?.(read);
  }
  return hash.digest("hex");
}

/** True when the model is present and its digest matches the catalog. */
export async function isInstalled(modelsDir, model) {
  const file = modelFilePath(modelsDir, model);
  return (await sizeOf(file)) === model.bytes;
}

/**
 * Smooths instantaneous byte rates into a usable speed/ETA readout.
 */
class RateEstimator {
  #samples = [];
  #windowMs = 8000;

  update(bytes) {
    const now = Date.now();
    this.#samples.push([now, bytes]);
    while (this.#samples.length > 2 && now - this.#samples[0][0] > this.#windowMs) this.#samples.shift();
  }

  /** @returns {number} bytes per second, 0 until there is enough signal */
  get bytesPerSecond() {
    if (this.#samples.length < 2) return 0;
    const [t0, b0] = this.#samples[0];
    const [t1, b1] = this.#samples[this.#samples.length - 1];
    const dt = (t1 - t0) / 1000;
    return dt > 0.25 ? Math.max(0, (b1 - b0) / dt) : 0;
  }
}

/**
 * Download `model` into `modelsDir`, resuming any prior partial file.
 *
 * @param {object} opts
 * @param {string} opts.modelsDir
 * @param {import('./catalog.js').ModelSpec} opts.model
 * @param {AbortSignal} [opts.signal]
 * @param {(p:{phase:string,received:number,total:number,percent:number,bytesPerSecond:number,etaSeconds:number|null}) => void} [opts.onProgress]
 * @param {typeof globalThis.fetch} [opts.fetchImpl]  injectable for tests
 */
export async function downloadModel({ modelsDir, model, signal, onProgress, fetchImpl = fetch }) {
  await mkdir(modelsDir, { recursive: true });
  const finalPath = modelFilePath(modelsDir, model);
  const partPath = partFilePath(modelsDir, model);

  if (await isInstalled(modelsDir, model)) {
    onProgress?.({ phase: "done", received: model.bytes, total: model.bytes, percent: 100, bytesPerSecond: 0, etaSeconds: 0 });
    return finalPath;
  }

  let received = await sizeOf(partPath);
  // A part file bigger than the target means it is from a different build — start over.
  if (received > model.bytes) {
    await rm(partPath, { force: true });
    received = 0;
  }

  const rate = new RateEstimator();
  const emit = (phase) => {
    rate.update(received);
    const bps = rate.bytesPerSecond;
    const remaining = model.bytes - received;
    onProgress?.({
      phase,
      received,
      total: model.bytes,
      percent: model.bytes ? (received / model.bytes) * 100 : 0,
      bytesPerSecond: bps,
      etaSeconds: bps > 0 ? Math.round(remaining / bps) : null,
    });
  };

  if (received < model.bytes) {
    emit("downloading");

    const headers = { "User-Agent": "ExcelAILocal/1.0" };
    if (received > 0) headers.Range = `bytes=${received}-`;

    const res = await fetchImpl(downloadUrl(model), { headers, signal, redirect: "follow" });

    if (!res.ok || !res.body) {
      throw new Error(
        res.status === 416
          ? "Download state was inconsistent. Remove the model and try again."
          : `Download failed: HTTP ${res.status} ${res.statusText || ""}`.trim(),
      );
    }
    // A server that ignores Range replies 200 and restarts the stream from zero.
    if (received > 0 && res.status !== 206) {
      await rm(partPath, { force: true });
      received = 0;
    }

    const out = createWriteStream(partPath, { flags: received > 0 ? "a" : "w" });
    let lastEmit = 0;
    try {
      await pipeline(
        (async function* () {
          for await (const chunk of res.body) {
            received += chunk.length;
            const now = Date.now();
            if (now - lastEmit > 200) {
              lastEmit = now;
              emit("downloading");
            }
            yield chunk;
          }
        })(),
        out,
        { signal },
      );
    } catch (err) {
      if (signal?.aborted) {
        // Keep the .part file so the next attempt resumes instead of restarting.
        const e = new Error("Download cancelled");
        e.code = "ABORTED";
        throw e;
      }
      throw err;
    }
  }

  const finalSize = await sizeOf(partPath);
  if (finalSize !== model.bytes) {
    await rm(partPath, { force: true });
    throw new Error(`Download incomplete: expected ${model.bytes} bytes, got ${finalSize}. Please try again.`);
  }

  onProgress?.({ phase: "verifying", received: model.bytes, total: model.bytes, percent: 100, bytesPerSecond: 0, etaSeconds: null });
  const digest = await hashFile(partPath);
  if (digest !== model.sha256) {
    await rm(partPath, { force: true });
    throw new Error("Downloaded file failed its integrity check and was discarded. Please try again.");
  }

  await rename(partPath, finalPath);
  onProgress?.({ phase: "done", received: model.bytes, total: model.bytes, percent: 100, bytesPerSecond: 0, etaSeconds: 0 });
  return finalPath;
}

export async function deleteModel(modelsDir, model) {
  await rm(modelFilePath(modelsDir, model), { force: true });
  await rm(partFilePath(modelsDir, model), { force: true });
}
