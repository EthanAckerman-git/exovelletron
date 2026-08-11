/**
 * Control panel renderer. Talks to the main process only through window.eal.
 */

const $ = (id) => document.getElementById(id);
const el = (tag, className, text) => {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
};

const CHECK = '<svg viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2.5 6.2l2.4 2.4L9.5 3.6"/></svg>';

const state = {
  modelsById: new Map(),
  tiers: { fast: null, balanced: null, max: null },
  machine: { chip: null, availableBytes: 0 },
  progress: new Map(),
  busySteps: new Set(),
  stepErrors: new Map(),
};

/** Display order and copy for the recommendation tiers. */
const TIERS = [
  { key: "fast", label: "Fast", note: "Quick replies, light footprint" },
  { key: "balanced", label: "Balanced", note: "Quality and speed in proportion" },
  { key: "max", label: "Max quality", note: "The smartest this Mac can run" },
];

const gb = (n) => `${(n / 1e9).toFixed(n < 1e10 ? 2 : 1)} GB`;
const rate = (n) => (n > 0 ? `${(n / 1e6).toFixed(1)} MB/s` : "—");
const eta = (s) => (s == null ? "—" : s > 60 ? `${Math.floor(s / 60)}m ${s % 60}s` : `${s}s`);

let toastTimer = null;
function toast(message) {
  const node = $("toast");
  node.textContent = message;
  node.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { node.hidden = true; }, 3800);
}

/* --------------------------------------------------------------------- render */

function renderStatus(app) {
  const engineState = app.serverError ? "error" : app.engine.state;
  $("dot").dataset.state = engineState;

  const model = state.modelsById.get(app.activeModelId);
  $("statusText").textContent = app.serverError
    ? app.serverError
    : {
        ready: model ? `Ready · ${model.name}` : "Ready",
        loading: "Loading model…",
        generating: "Answering a question…",
        error: app.engine.error || "Engine error",
        idle: "No model loaded",
      }[engineState] ?? engineState;

  $("endpoint").textContent = app.serverError ? "offline" : `127.0.0.1:${app.port}`;
  $("version").textContent = `v${app.version}`;
  $("footPath").textContent = app.modelsDir;
  const machineNote = [app.machine.chip, `${app.machine.totalRamGb} GB RAM`].filter(Boolean).join(" · ");
  $("chipNote").textContent = machineNote;
  // The chart's own header repeats the machine only when the recommended section —
  // which normally carries it — has nothing to show.
  $("ramNote").textContent = $("recommendSection").hidden ? machineNote : "";
}

/** Five dots, `n` of them lit. The chart's shorthand for relative intelligence. */
function dotsMeter(n) {
  const meter = el("span", "dots");
  meter.title = `Intelligence ${n} of 5 within this family`;
  for (let i = 1; i <= 5; i++) {
    const dot = el("i");
    if (i <= n) dot.dataset.on = "true";
    meter.appendChild(dot);
  }
  return meter;
}

/** Switch the active model, with the button showing what is happening. */
async function useModel(model, btn) {
  btn.disabled = true;
  btn.textContent = "Loading…";
  try {
    await window.eal.models.select(model.id);
    toast(`${model.name} is now active.`);
  } catch (err) {
    toast(err.message.replace(/^Error invoking remote method '[^']+':\s*/, ""));
  }
  await refresh();
}

async function startDownload(model, btn) {
  btn.disabled = true;
  await window.eal.models.download(model.id);
  await refresh();
}

function renderSetup(app) {
  const { steps, ready } = app.preflight;
  $("setupSection").hidden = ready;
  $("readySection").hidden = !ready;
  $("setupCount").textContent = ready ? "" : `${steps.filter((s) => s.done).length} of ${steps.length}`;
  if (ready) return;

  const list = $("steps");
  list.textContent = "";

  for (const step of steps) {
    const li = el("li", "step");
    li.dataset.done = String(step.done);
    const busy = state.busySteps.has(step.id);
    if (busy) li.dataset.busy = "true";

    const mark = el("div", "step__mark");
    if (step.done) mark.innerHTML = CHECK;
    li.appendChild(mark);

    const body = el("div", "step__body");
    body.appendChild(el("p", "step__title", step.title));
    body.appendChild(el("p", "step__detail", step.detail));
    if (step.note && !step.done) body.appendChild(el("p", "step__note", step.note));
    const error = state.stepErrors.get(step.id);
    if (error) body.appendChild(el("p", "step__error", error));

    // The model step is driven from the recommendation cards below, so no button here.
    if (!step.done && step.id === "model") {
      body.appendChild(el("p", "step__note", "Pick a recommended model below — Balanced is the safe choice."));
      li.appendChild(body);
    } else if (!step.done && step.blocked) {
      // macOS seals Excel's container against other processes, but honours a folder the
      // user picks in an open panel. One click, one confirm, done.
      const how = el("ol", "step__how");
      how.appendChild(el("li", null, "A folder picker opens, already pointing at Excel's “wef” folder."));
      how.appendChild(el("li", null, "Click Grant Access. The add-in installs itself."));
      how.appendChild(el("li", null, "Quit and reopen Excel."));
      body.appendChild(how);

      const row = el("div", "step__buttons");
      const grant = el("button", "btn btn--primary", busy ? "Installing…" : "Grant access and install");
      grant.disabled = busy;
      grant.addEventListener("click", () => runStep(step, "addin-grant"));

      const manual = el("button", "linkish", "Install by hand instead");
      manual.addEventListener("click", async () => {
        await window.eal.setup.manualAddin();
        toast("Drag the highlighted file into the wef folder.");
      });
      row.append(grant, manual);
      body.appendChild(row);
      li.appendChild(body);
    } else if (!step.done) {
      const btn = el("button", "btn btn--primary", busy ? "Working…" : step.action);
      btn.disabled = busy;
      btn.addEventListener("click", () => runStep(step));
      li.append(body, btn);
    } else {
      li.appendChild(body);
    }
    list.appendChild(li);
  }
}

function renderPicks(app) {
  const section = $("recommendSection");
  const container = $("picks");
  container.textContent = "";

  const picks = TIERS
    .map((tier) => ({ ...tier, model: state.modelsById.get(state.tiers?.[tier.key]) }))
    .filter((p) => p.model);
  section.hidden = picks.length === 0;

  for (const { label, note, model } of picks) {
    const card = el("div", "pick");
    card.dataset.active = String(app.activeModelId === model.id);

    const head = el("div", "pick__head");
    head.appendChild(el("span", "pick__tier", label));
    head.appendChild(dotsMeter(model.intelligence));
    card.appendChild(head);

    card.appendChild(el("h3", "pick__name", model.name));
    card.appendChild(el("p", "pick__gain", model.gain));

    const foot = el("div", "pick__foot");
    if (app.activeModelId === model.id) {
      foot.appendChild(el("span", "tag tag--active", "Active"));
    } else if (model.downloading) {
      const btn = el("button", "btn btn--quiet", "Downloading…");
      btn.disabled = true;
      foot.appendChild(btn);
    } else if (model.installed) {
      const use = el("button", "btn btn--primary", "Use");
      use.addEventListener("click", () => useModel(model, use));
      foot.appendChild(use);
    } else {
      const download = el("button", "btn btn--primary", `Download ${gb(model.bytes)}`);
      download.addEventListener("click", () => startDownload(model, download));
      foot.appendChild(download);
    }
    foot.appendChild(el("span", "pick__note", note));
    card.appendChild(foot);

    container.appendChild(card);
  }
}

function renderModels(app) {
  const container = $("models");
  container.textContent = "";

  for (const model of state.modelsById.values()) {
    const card = el("div", "model");
    card.dataset.id = model.id;
    card.dataset.active = String(app.activeModelId === model.id);
    card.dataset.fit = model.fit.level;

    const top = el("div", "model__top");
    top.appendChild(el("h3", "model__name", model.name));
    if (app.activeModelId === model.id) top.appendChild(el("span", "tag tag--active", "Active"));
    else if (model.default) top.appendChild(el("span", "tag tag--rec", model.tagline));
    else top.appendChild(el("span", "tag", model.tagline));
    top.appendChild(dotsMeter(model.intelligence));
    card.appendChild(top);

    const spec = el("p", "model__spec");
    // The window the engine actually runs, not the model's 256K training maximum —
    // that number set expectations the app deliberately does not meet.
    const runtimeK = Math.round((app.engine.contextTokens ?? 8192) / 1024);
    spec.innerHTML = `<b>${gb(model.bytes)}</b> · ${model.params} · ${model.quant} · ${runtimeK}K context`;
    spec.title = `${model.name} supports up to ${Math.round(model.contextTokens / 1024)}K of context; the app runs a ${runtimeK}K window for speed and memory.`;
    card.appendChild(spec);

    card.appendChild(el("p", "model__why", model.strengths.join(" · ")));

    const progress = state.progress.get(model.id);
    if (progress) {
      const bar = el("div", `progress${progress.phase === "verifying" ? " progress--verify" : ""}`);
      const fill = el("i");
      fill.style.width = `${Math.max(2, progress.percent)}%`;
      bar.appendChild(fill);
      card.appendChild(bar);

      const meta = el("div", "model__progress");
      meta.appendChild(el("span", null, progress.phase === "verifying"
        ? "Verifying integrity…"
        : `${progress.percent.toFixed(0)}% · ${gb(progress.received)} of ${gb(progress.total)}`));
      meta.appendChild(el("span", null, progress.phase === "verifying"
        ? ""
        : `${rate(progress.bytesPerSecond)} · ${eta(progress.etaSeconds)}`));
      card.appendChild(meta);
    }

    const row = el("div", "model__row");

    if (model.downloading) {
      const cancel = el("button", "btn btn--quiet", "Cancel");
      cancel.addEventListener("click", async () => {
        await window.eal.models.cancel(model.id);
        state.progress.delete(model.id);
        toast(`Cancelled ${model.name}. Progress is kept for next time.`);
        await refresh();
      });
      row.appendChild(cancel);
    } else if (!model.installed) {
      const label = model.partialBytes > 0 ? `Resume (${gb(model.partialBytes)} done)` : `Download ${gb(model.bytes)}`;
      const download = el("button", "btn btn--primary", label);
      if (model.fit.level === "too-big") {
        // Visible so the chart stays complete, disabled so nobody downloads 17 GB
        // of model this Mac cannot hold.
        download.disabled = true;
        download.title = "This model needs more memory than this Mac can offer it.";
      }
      download.addEventListener("click", () => startDownload(model, download));
      row.appendChild(download);
    } else if (app.activeModelId !== model.id) {
      const use = el("button", "btn btn--primary", "Use this model");
      use.addEventListener("click", () => useModel(model, use));
      row.appendChild(use);
    }

    if (model.installed && !model.downloading) {
      const remove = el("button", "btn btn--danger", "Delete");
      remove.addEventListener("click", async () => {
        remove.disabled = true;
        try {
          await window.eal.models.remove(model.id);
          toast(`${model.name} deleted.`);
        } catch (err) {
          toast(err.message);
        }
        await refresh();
      });
      row.appendChild(remove);
    }

    const fit = el("span", "model__fit", model.fit.label);
    fit.dataset.level = model.fit.level;
    row.appendChild(fit);

    card.appendChild(row);
    container.appendChild(card);
  }
}

/* ---------------------------------------------------------------------- steps */

/**
 * @param {object} step
 * @param {string} [action]  what to run; defaults to the step's own id. Kept separate so
 *                           a step can offer more than one route to done without its
 *                           busy state and error slot drifting onto a different key.
 */
async function runStep(step, action = step.id) {
  state.busySteps.add(step.id);
  state.stepErrors.delete(step.id);
  await refresh();
  try {
    if (action === "certificate") {
      await window.eal.setup.certificate();
      toast("Certificate created and trusted.");
    } else if (action === "addin-grant") {
      const result = await window.eal.setup.grantAddinAccess();
      if (result.cancelled) toast("Cancelled. The add-in was not installed.");
      else toast("Add-in installed. Quit and reopen Excel to see it.");
    } else if (action === "addin") {
      await window.eal.setup.addin();
      toast("Add-in installed. Restart Excel to see it.");
    }
  } catch (err) {
    state.stepErrors.set(step.id, err.message.replace(/^Error invoking remote method '[^']+':\s*/, ""));
  } finally {
    state.busySteps.delete(step.id);
    await refresh();
  }
}

/* --------------------------------------------------------------------- wiring */

async function refresh() {
  const [app, listing] = await Promise.all([window.eal.getState(), window.eal.models.list()]);
  state.modelsById = new Map(listing.models.map((m) => [m.id, m]));
  state.tiers = listing.tiers;
  state.machine = listing.machine;
  renderStatus(app);
  renderSetup(app);
  renderPicks(app);
  renderModels(app);
}

function bind() {
  $("openExcel").addEventListener("click", () => window.eal.actions.openExcel());
  $("revealModels").addEventListener("click", () => window.eal.actions.revealModels());
  $("openRepo").addEventListener("click", () => window.eal.actions.openRepo());
  $("checkUpdates").addEventListener("click", async () => {
    const btn = $("checkUpdates");
    btn.disabled = true;
    btn.textContent = "Checking…";
    try {
      const info = await window.eal.actions.checkUpdates();
      if (info.status === "update") toast(`Version ${info.latest} is available — see the tray menu to download.`);
      else if (info.status === "current") toast("You're up to date.");
      else toast("Couldn't check for updates — offline, or the GitHub releases page isn't public yet.");
    } finally {
      btn.disabled = false;
      btn.textContent = "Check for updates";
    }
  });

  window.eal.onProgress((p) => {
    state.progress.set(p.id, p);
    // Repainting the whole list on every progress tick would be wasteful; update in place.
    const card = document.querySelector(`.model[data-id="${p.id}"]`);
    if (!card) return;
    const fill = card.querySelector(".progress i");
    const meta = card.querySelector(".model__progress");
    if (!fill || !meta) { refresh(); return; }
    fill.style.width = `${Math.max(2, p.percent)}%`;
    card.querySelector(".progress").className = `progress${p.phase === "verifying" ? " progress--verify" : ""}`;
    meta.children[0].textContent = p.phase === "verifying"
      ? "Verifying integrity…"
      : `${p.percent.toFixed(0)}% · ${gb(p.received)} of ${gb(p.total)}`;
    meta.children[1].textContent = p.phase === "verifying" ? "" : `${rate(p.bytesPerSecond)} · ${eta(p.etaSeconds)}`;
  });

  window.eal.onInstalled(async ({ id }) => {
    state.progress.delete(id);
    toast(`${state.modelsById.get(id)?.name ?? "Model"} downloaded and verified.`);
    await refresh();
  });

  window.eal.onRemoved(() => refresh());

  window.eal.onModelError(async ({ id, message, cancelled }) => {
    state.progress.delete(id);
    if (!cancelled) toast(message);
    await refresh();
  });

  window.eal.onEngineState(() => refresh());
  window.eal.onServerError((message) => toast(message));
  window.eal.onEngineError((message) => toast(message));
}

bind();
refresh().catch((err) => toast(err.message));
