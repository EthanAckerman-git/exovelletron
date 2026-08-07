/**
 * Electron main process: owns the window, the tray, the HTTPS server, and the engine.
 *
 * The renderer never touches Node APIs — everything goes through the narrow IPC surface
 * declared in preload.js.
 */
import { app, BrowserWindow, Tray, Menu, shell, ipcMain, nativeImage, dialog } from "electron";
import path from "node:path";
import { existsSync } from "node:fs";
import os from "node:os";
import { fileURLToPath } from "node:url";

import { paths, ensureDirs, loadConfig, saveConfig, DEFAULT_PORT } from "../core/config.js";
import { ensureCertFiles, ensureTrustedCerts } from "../core/setup/certs.js";
import { installManifest, isManifestInstalled, manifestPath, exportManifest, installManifestAtChosenDir } from "../core/setup/manifest.js";
import { preflight } from "../core/setup/preflight.js";
import { createAppServer } from "../core/server/server.js";
import { ModelStore } from "../core/models/store.js";
import { Engine } from "../core/llm/engine.js";
import { getModel, DEFAULT_MODEL_ID } from "../core/models/catalog.js";

const DIR = path.dirname(fileURLToPath(import.meta.url));

let win = null;
let tray = null;
let server = null;
let engine = null;
let models = null;
let config = null;
let activePort = DEFAULT_PORT;
let serverError = null;

/* ------------------------------------------------------------------------ window */

function createWindow() {
  win = new BrowserWindow({
    width: 760,
    height: 660,
    minWidth: 620,
    minHeight: 520,
    show: false,
    title: "Excel AI Local",
    titleBarStyle: "hiddenInset",
    backgroundColor: "#141318",
    webPreferences: {
      preload: path.join(DIR, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  win.loadFile(path.join(DIR, "renderer", "index.html"));
  win.once("ready-to-show", () => win.show());
  win.on("closed", () => { win = null; });

  // Keep navigation inside the app; open any real link in the user's browser.
  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: "deny" };
  });
}

const showWindow = () => {
  if (win) {
    win.show();
    win.focus();
  } else {
    createWindow();
  }
};

/* -------------------------------------------------------------------------- tray */

function createTray() {
  // A template image lets macOS invert it correctly for light and dark menu bars.
  const icon = nativeImage.createFromPath(path.join(DIR, "..", "assets", "trayTemplate.png"));
  icon.setTemplateImage(true);
  tray = new Tray(icon.isEmpty() ? nativeImage.createEmpty() : icon);
  tray.setToolTip("Excel AI Local");
  refreshTrayMenu();
  tray.on("click", showWindow);
}

function refreshTrayMenu() {
  if (!tray) return;
  const status = engine?.status();
  const label = serverError
    ? "Engine not running"
    : status?.state === "ready"
      ? `Ready · ${getModel(status.modelId)?.name ?? "model loaded"}`
      : status?.state === "loading"
        ? "Loading model…"
        : status?.state === "generating"
          ? "Answering…"
          : "No model loaded";

  tray.setContextMenu(Menu.buildFromTemplate([
    { label, enabled: false },
    { type: "separator" },
    { label: "Open Excel AI Local", click: showWindow },
    { label: "Open Microsoft Excel", click: () => shell.openPath("/Applications/Microsoft Excel.app") },
    { type: "separator" },
    { label: "Quit", click: () => app.quit() },
  ]));
}

/* ------------------------------------------------------------------------ server */

async function startServer() {
  serverError = null;
  const credentials = await ensureCertFiles(paths);

  // The manifest hard-codes the port, so prefer the configured one and only move if
  // something else already holds it — then rewrite the manifest to match.
  const candidates = [config.port, DEFAULT_PORT, 39218, 39219, 39220].filter((p, i, a) => a.indexOf(p) === i);

  for (const port of candidates) {
    const instance = createAppServer({
      credentials,
      port,
      engine,
      models,
      appInfo: { version: app.getVersion() },
    });
    try {
      await instance.listen();
      server = instance;
      activePort = port;
      if (config.port !== port) config = await saveConfig({ port }, paths);

      // Refreshing the manifest is best-effort. macOS may deny access to Excel's
      // container until the user grants Full Disk Access, and that must not stop the
      // engine from running — the wizard surfaces it as its own step.
      if (!(await isManifestInstalled(port, paths))) {
        await installManifest(port, paths).catch((err) => {
          if (err.code !== "TCC_DENIED") throw err;
        });
      }
      return;
    } catch (err) {
      if (err.code !== "EADDRINUSE") throw err;
    }
  }
  throw new Error("Could not find a free port for the local server.");
}

async function startEngine() {
  const id = config.modelId ?? DEFAULT_MODEL_ID;
  if (!(await models.isInstalled(id))) return;
  engine.updateConfig({
    contextTokens: config.contextTokens,
    sheetContextTokens: config.sheetContextTokens,
    temperature: config.temperature,
    maxTokens: config.maxTokens,
  });
  await engine.load(id, models.pathFor(id));
}

/* --------------------------------------------------------------------------- IPC */

const send = (channel, payload) => win?.webContents.send(channel, payload);

function registerIpc() {
  ipcMain.handle("app:state", async () => ({
    version: app.getVersion(),
    port: activePort,
    serverError,
    engine: engine.status(),
    activeModelId: engine.modelId,
    config,
    machine: { totalRamGb: Math.round(os.totalmem() / 1024 ** 3), arch: os.arch(), cpus: os.cpus().length },
    preflight: await preflight({ paths, port: activePort, modelId: config.modelId ?? DEFAULT_MODEL_ID }),
    manifestPath: manifestPath(paths),
    modelsDir: paths.modelsDir,
  }));

  ipcMain.handle("models:list", () => models.list());

  ipcMain.handle("models:download", (_e, id) => {
    models.download(id).then(
      async () => {
        // First model in: start it so the pane becomes usable without another click.
        if (!engine.modelId) {
          await saveConfig({ modelId: id }, paths);
          config = await loadConfig(paths);
          await startEngine().catch(() => {});
        }
      },
      () => { /* reported through the error event */ },
    );
    return { started: true };
  });

  ipcMain.handle("models:cancel", (_e, id) => models.cancel(id));

  ipcMain.handle("models:remove", async (_e, id) => {
    if (engine.modelId === id) await engine.unload();
    await models.remove(id);
    return true;
  });

  ipcMain.handle("models:select", async (_e, id) => {
    if (!(await models.isInstalled(id))) throw new Error("That model is not downloaded yet.");
    config = await saveConfig({ modelId: id }, paths);
    await engine.load(id, models.pathFor(id));
    return engine.status();
  });

  ipcMain.handle("setup:certificate", async () => {
    await ensureTrustedCerts(paths);
    // The running server holds the old credentials; restart it on the new pair.
    if (server) { await server.close(); server = null; }
    await startServer();
    return true;
  });

  ipcMain.handle("setup:addin", async () => {
    await installManifest(activePort, paths);
    return true;
  });

  /**
   * Install the add-in via a folder the user picks in a native open panel.
   *
   * On macOS 26+ an app-container Data directory is sealed against every other process.
   * Measured on this machine: the shell, a signed app, scripted Finder, and Terminal
   * *with* Full Disk Access are all refused with EPERM, and no permission prompt is
   * offered — so Full Disk Access is deliberately never suggested to the user.
   *
   * Picking the folder in an open panel is different: it is explicit consent, so macOS
   * issues a sandbox extension and the write succeeds. Verified end to end — EPERM
   * before the pick, success after.
   */
  ipcMain.handle("setup:grantAddinAccess", async () => {
    const startAt = existsSync(paths.wefDir) ? paths.wefDir : path.dirname(paths.wefDir);
    const result = await dialog.showOpenDialog(win, {
      title: "Select Excel's add-in folder",
      message: "Select the “wef” folder so Excel AI Local can install the add-in",
      defaultPath: startAt,
      properties: ["openDirectory", "createDirectory"],
      buttonLabel: "Grant Access",
    });

    if (result.canceled || !result.filePaths?.length) return { ok: false, cancelled: true };

    const file = await installManifestAtChosenDir(result.filePaths[0], activePort, paths);
    return { ok: true, file };
  });

  /** Last-resort fallback: stage the manifest and reveal it for a manual drag. */
  ipcMain.handle("setup:manualAddin", async () => {
    const exported = await exportManifest(activePort, path.join(paths.dataDir, "addin"), paths);
    shell.showItemInFolder(exported);
    setTimeout(async () => {
      const opened = await shell.openPath(paths.wefDir);
      if (opened) await shell.openPath(path.dirname(paths.wefDir));
    }, 700);
    return { file: exported, destination: paths.wefDir };
  });

  ipcMain.handle("app:openExcel", () => shell.openPath("/Applications/Microsoft Excel.app"));
  ipcMain.handle("app:revealModels", () => shell.openPath(paths.modelsDir));
  ipcMain.handle("app:openLanding", () => shell.openExternal(`https://localhost:${activePort}/`));

  ipcMain.handle("config:set", async (_e, patch) => {
    config = await saveConfig(patch, paths);
    engine.updateConfig({
      temperature: config.temperature,
      maxTokens: config.maxTokens,
      contextTokens: config.contextTokens,
    });
    return config;
  });
}

/* -------------------------------------------------------------------- lifecycle */

async function boot() {
  await ensureDirs(paths);
  config = await loadConfig(paths);

  models = new ModelStore(paths.modelsDir);
  engine = new Engine({
    config: {
      contextTokens: config.contextTokens,
      sheetContextTokens: config.sheetContextTokens,
      temperature: config.temperature,
      maxTokens: config.maxTokens,
    },
  });

  models.on("progress", (p) => send("models:progress", p));
  models.on("installed", (p) => { send("models:installed", p); refreshTrayMenu(); });
  models.on("removed", (p) => send("models:removed", p));
  models.on("error", (p) => send("models:error", p));
  engine.on("state", (s) => { send("engine:state", s); refreshTrayMenu(); });

  registerIpc();
  createWindow();
  createTray();

  try {
    await startServer();
  } catch (err) {
    serverError = err.message;
    send("app:serverError", err.message);
    dialog.showErrorBox("Excel AI Local could not start", err.message);
    // Note: manifest-permission failures never reach here — they are surfaced as a
    // setup step, not as a fatal startup error.
  }

  refreshTrayMenu();
  startEngine().catch((err) => send("app:engineError", err.message));
}

app.whenReady().then(boot);

// A menu-bar utility should keep running when its window is closed.
app.on("window-all-closed", () => { /* stay resident */ });
app.on("activate", showWindow);

app.on("before-quit", async () => {
  try { await engine?.unload(); } catch { /* shutting down */ }
  try { await server?.close(); } catch { /* shutting down */ }
});
