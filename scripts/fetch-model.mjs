/**
 * CLI model fetcher — used by `npm run fetch-model` and during development.
 * Usage: node scripts/fetch-model.mjs [modelId]
 */
import { ModelStore } from "../core/models/store.js";
import { DEFAULT_MODEL_ID, formatBytes } from "../core/models/catalog.js";
import { paths, ensureDirs } from "../core/config.js";

const id = process.argv[2] || DEFAULT_MODEL_ID;

await ensureDirs();
const store = new ModelStore(paths.modelsDir);

let lastLine = 0;
store.on("progress", (p) => {
  const now = Date.now();
  if (p.phase !== "done" && now - lastLine < 1000) return;
  lastLine = now;
  const speed = p.bytesPerSecond ? `${formatBytes(p.bytesPerSecond)}/s` : "—";
  const eta = p.etaSeconds != null ? `${Math.floor(p.etaSeconds / 60)}m ${p.etaSeconds % 60}s` : "—";
  console.log(`${p.phase} ${p.percent.toFixed(1)}%  ${formatBytes(p.received)}/${formatBytes(p.total)}  ${speed}  eta ${eta}`);
});

try {
  const file = await store.download(id);
  console.log(`\nInstalled: ${file}`);
} catch (err) {
  console.error(`\nFailed: ${err.message}`);
  process.exit(1);
}
