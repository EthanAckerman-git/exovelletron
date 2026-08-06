/**
 * The only bridge between the renderer and Node. Everything is an explicit, named
 * channel — the renderer gets no filesystem, no shell, and no ipcRenderer handle.
 */
const { contextBridge, ipcRenderer } = require("electron");

/** Wrap a push channel so callers get an unsubscribe function rather than a raw listener. */
const on = (channel) => (handler) => {
  const wrapped = (_event, payload) => handler(payload);
  ipcRenderer.on(channel, wrapped);
  return () => ipcRenderer.removeListener(channel, wrapped);
};

contextBridge.exposeInMainWorld("eal", {
  getState: () => ipcRenderer.invoke("app:state"),

  models: {
    list: () => ipcRenderer.invoke("models:list"),
    download: (id) => ipcRenderer.invoke("models:download", id),
    cancel: (id) => ipcRenderer.invoke("models:cancel", id),
    remove: (id) => ipcRenderer.invoke("models:remove", id),
    select: (id) => ipcRenderer.invoke("models:select", id),
  },

  setup: {
    certificate: () => ipcRenderer.invoke("setup:certificate"),
    addin: () => ipcRenderer.invoke("setup:addin"),
    manualAddin: () => ipcRenderer.invoke("setup:manualAddin"),
  },

  actions: {
    openExcel: () => ipcRenderer.invoke("app:openExcel"),
    revealModels: () => ipcRenderer.invoke("app:revealModels"),
    openLanding: () => ipcRenderer.invoke("app:openLanding"),
    openPrivacySettings: () => ipcRenderer.invoke("app:openPrivacySettings"),
  },

  setConfig: (patch) => ipcRenderer.invoke("config:set", patch),

  onProgress: on("models:progress"),
  onInstalled: on("models:installed"),
  onRemoved: on("models:removed"),
  onModelError: on("models:error"),
  onEngineState: on("engine:state"),
  onServerError: on("app:serverError"),
  onEngineError: on("app:engineError"),
});
