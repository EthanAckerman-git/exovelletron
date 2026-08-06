/**
 * The checks behind the setup wizard.
 *
 * Each step reports whether it is done and, when it isn't, a sentence the user can act
 * on. The wizard renders this list directly, so the wording here is what people read.
 */
import { existsSync } from "node:fs";
import { mkdir, writeFile, rm } from "node:fs/promises";
import path from "node:path";
import { certsExist, isCaTrusted } from "./certs.js";
import { isManifestInstalled } from "./manifest.js";
import { isInstalled } from "../models/downloader.js";
import { getModel } from "../models/catalog.js";

export const STEP = {
  CERTIFICATE: "certificate",
  ADDIN: "addin",
  MODEL: "model",
};

/**
 * Can we actually write into Excel's sideload folder?
 *
 * Probing costs one file create/delete and is worth it: without Full Disk Access every
 * write fails with EPERM and macOS raises no prompt, so the wizard has to detect the
 * condition itself rather than waiting for the user to hit an error.
 *
 * @returns {Promise<"ok"|"blocked">}
 */
export async function checkAddinAccess(paths) {
  const probe = path.join(paths.wefDir, ".eal-write-probe");
  try {
    await mkdir(paths.wefDir, { recursive: true });
    await writeFile(probe, "", "utf8");
    await rm(probe, { force: true });
    return "ok";
  } catch (err) {
    if (err?.code === "EPERM" || err?.code === "EACCES" || err?.code === "EROFS") return "blocked";
    return "blocked";
  }
}

/**
 * @param {object} opts
 * @param {object} opts.paths
 * @param {number} opts.port
 * @param {string} opts.modelId
 * @returns {Promise<{ready:boolean, steps:Array}>}
 */
export async function preflight({ paths, port, modelId }) {
  const model = getModel(modelId);

  const [certs, trusted, manifest, modelReady] = await Promise.all([
    certsExist(paths),
    isCaTrusted(paths),
    isManifestInstalled(port, paths),
    model ? isInstalled(paths.modelsDir, model) : Promise.resolve(false),
  ]);

  // Only probe when the manifest isn't already in place — if it is, access clearly works.
  const access = manifest ? "ok" : await checkAddinAccess(paths);

  const steps = [
    {
      id: STEP.CERTIFICATE,
      title: "Local certificate",
      detail: certs && trusted
        ? "Excel can load the pane securely from this Mac."
        : "Excel only loads add-ins over HTTPS. This creates a certificate for localhost and trusts it.",
      done: certs && trusted,
      needsPassword: true,
      action: "Create and trust",
    },
    {
      id: STEP.ADDIN,
      title: "Excel add-in",
      detail: manifest
        ? "Installed. Look for Excel AI Local on the Home tab."
        : access === "blocked"
          ? "macOS keeps Excel's add-in folder sealed off from other apps, so this last step needs one drag from you. It takes about ten seconds and you only do it once."
          : "Registers the add-in with Excel so it appears on the Home tab.",
      done: manifest,
      needsPassword: false,
      blocked: !manifest && access === "blocked",
      action: "Install",
      // Excel only reads manifests at launch.
      note: manifest ? null : "Excel must be restarted after this step.",
    },
    {
      id: STEP.MODEL,
      title: model ? `${model.name} model` : "Language model",
      detail: modelReady
        ? "Downloaded and verified."
        : model
          ? `A one-time ${(model.bytes / 1e9).toFixed(1)} GB download. Everything after this works offline.`
          : "Choose a model to download.",
      done: modelReady,
      needsPassword: false,
      action: "Download",
    },
  ];

  return { ready: steps.every((s) => s.done), steps };
}

/** True when Excel's sandbox container exists, i.e. Excel has been run at least once. */
export const excelHasBeenLaunched = (paths) =>
  existsSync(paths.wefDir) ||
  existsSync(paths.wefDir.replace(/\/wef$/, ""));
