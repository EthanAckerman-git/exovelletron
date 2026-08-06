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

const state = { modelsById: new Map(), progress: new Map(), busySteps: new Set(), stepErrors: new Map() };

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
  $("ramNote").textContent = `${app.machine.totalRamGb} GB RAM · ${app.machine.arch}`;
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

    // The model step is driven from the Models list below, so it gets no button here.
    if (!step.done && step.id === "model") {
      body.appendChild(el("p", "step__note", "Choose a model below to download."));
      li.appendChild(body);
    } else if (!step.done && step.blocked) {
      // macOS seals Excel's container against every other process, so this step is a
      // hand-off: two Finder windows and one drag.
      const how = el("ol", "step__how");
      how.appendChild(el("li", null, "Two Finder windows open: the manifest, and Excel's add-ins folder."));
      how.appendChild(el("li", null, "Drag excel-ai-local.manifest.xml into the Excel folder."));
      how.appendChild(el("li", null, "Quit and reopen Excel, then click Check again."));
      body.appendChild(how);

      const row = el("div", "step__buttons");
      const open = el("button", "btn btn--primary", "Open both folders");
      open.addEventListener("click", async () => {
        await window.eal.setup.manualAddin();
        toast("Drag the highlighted file into the wef folder.");
      });
      const recheck = el("button", "btn btn--quiet", busy ? "Checking…" : "Check again");
      recheck.disabled = busy;
      recheck.addEventListener("click", () => runStep(step));
      row.append(open, recheck);
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

function renderModels(app) {
  const container = $("models");
  container.textContent = "";

  for (const model of state.modelsById.values()) {
    const card = el("div", "model");
    card.dataset.active = String(app.activeModelId === model.id);

    const top = el("div", "model__top");
    top.appendChild(el("h3", "model__name", model.name));
    if (app.activeModelId === model.id) top.appendChild(el("span", "tag tag--active", "Active"));
    else if (model.default) top.appendChild(el("span", "tag tag--rec", model.tagline));
    else top.appendChild(el("span", "tag", model.tagline));
    card.appendChild(top);

    const spec = el("p", "model__spec");
    spec.innerHTML = `<b>${gb(model.bytes)}</b> · ${model.params} · ${model.quant} · ${(model.contextTokens / 1024).toFixed(0)}K context`;
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
      download.addEventListener("click", async () => {
        download.disabled = true;
        await window.eal.models.download(model.id);
        await refresh();
      });
      row.appendChild(download);
    } else if (app.activeModelId !== model.id) {
      const use = el("button", "btn btn--primary", "Use this model");
      use.addEventListener("click", async () => {
        use.disabled = true;
        use.textContent = "Loading…";
        try {
          await window.eal.models.select(model.id);
          toast(`${model.name} is now active.`);
        } catch (err) {
          toast(err.message);
        }
        await refresh();
      });
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

async function runStep(step) {
  state.busySteps.add(step.id);
  state.stepErrors.delete(step.id);
  await refresh();
  try {
    if (step.id === "certificate") {
      await window.eal.setup.certificate();
      toast("Certificate created and trusted.");
    } else if (step.id === "addin") {
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
  const [app, models] = await Promise.all([window.eal.getState(), window.eal.models.list()]);
  state.modelsById = new Map(models.map((m) => [m.id, m]));
  renderStatus(app);
  renderSetup(app);
  renderModels(app);
}

function bind() {
  $("openExcel").addEventListener("click", () => window.eal.actions.openExcel());
  $("revealModels").addEventListener("click", () => window.eal.actions.revealModels());

  window.eal.onProgress((p) => {
    state.progress.set(p.id, p);
    // Repainting the whole list on every progress tick would be wasteful; update in place.
    const card = [...document.querySelectorAll(".model")].find(
      (n) => n.querySelector(".model__name")?.textContent === state.modelsById.get(p.id)?.name,
    );
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
