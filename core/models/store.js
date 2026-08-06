/**
 * Tracks which models are on disk and owns the one in-flight download.
 *
 * Downloads are deliberately serialised: two concurrent multi-gigabyte pulls just make
 * each other slower and confuse the progress UI.
 */
import { EventEmitter } from "node:events";
import { stat } from "node:fs/promises";
import os from "node:os";
import { CATALOG, getModel, fitForMachine } from "./catalog.js";
import { downloadModel, deleteModel, isInstalled, modelFilePath, partFilePath } from "./downloader.js";

const totalRamGb = () => Math.round(os.totalmem() / 1024 ** 3);

export class ModelStore extends EventEmitter {
  #modelsDir;
  #active = null; // { id, controller, progress }

  constructor(modelsDir) {
    super();
    this.#modelsDir = modelsDir;
  }

  get modelsDir() {
    return this.#modelsDir;
  }

  get activeDownloadId() {
    return this.#active?.id ?? null;
  }

  pathFor(id) {
    const model = getModel(id);
    return model ? modelFilePath(this.#modelsDir, model) : null;
  }

  /** Catalog entries decorated with on-disk state and machine fit. */
  async list() {
    const ram = totalRamGb();
    return Promise.all(
      CATALOG.map(async (model) => {
        const installed = await isInstalled(this.#modelsDir, model);
        let partialBytes = 0;
        if (!installed) {
          try {
            partialBytes = (await stat(partFilePath(this.#modelsDir, model))).size;
          } catch {
            partialBytes = 0;
          }
        }
        const active = this.#active?.id === model.id ? this.#active.progress : null;
        return {
          ...model,
          installed,
          partialBytes,
          downloading: this.#active?.id === model.id,
          progress: active,
          fit: fitForMachine(model, ram),
        };
      }),
    );
  }

  async isInstalled(id) {
    const model = getModel(id);
    return model ? isInstalled(this.#modelsDir, model) : false;
  }

  /**
   * Begin downloading `id`. Resolves when the download finishes; progress arrives via
   * the `progress` event so callers can stream it to the UI.
   */
  async download(id) {
    const model = getModel(id);
    if (!model) throw new Error(`Unknown model: ${id}`);
    if (this.#active) throw new Error(`Already downloading ${this.#active.id}. Cancel it first.`);
    if (await isInstalled(this.#modelsDir, model)) return modelFilePath(this.#modelsDir, model);

    const controller = new AbortController();
    this.#active = { id, controller, progress: null };

    try {
      const file = await downloadModel({
        modelsDir: this.#modelsDir,
        model,
        signal: controller.signal,
        onProgress: (p) => {
          if (this.#active) this.#active.progress = p;
          this.emit("progress", { id, ...p });
        },
      });
      this.emit("installed", { id });
      return file;
    } catch (err) {
      this.emit("error", { id, message: err.message, cancelled: err.code === "ABORTED" });
      throw err;
    } finally {
      this.#active = null;
    }
  }

  cancel(id) {
    if (this.#active && (!id || this.#active.id === id)) {
      this.#active.controller.abort();
      return true;
    }
    return false;
  }

  async remove(id) {
    const model = getModel(id);
    if (!model) throw new Error(`Unknown model: ${id}`);
    if (this.#active?.id === id) this.cancel(id);
    await deleteModel(this.#modelsDir, model);
    this.emit("removed", { id });
  }
}
