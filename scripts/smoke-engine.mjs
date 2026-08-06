/**
 * End-to-end smoke test against the real model.
 * Verifies load, streaming, tool calling, and action validation.
 * Usage: node scripts/smoke-engine.mjs
 */
import { Engine } from "../core/llm/engine.js";
import { ModelStore } from "../core/models/store.js";
import { DEFAULT_MODEL_ID } from "../core/models/catalog.js";
import { paths } from "../core/config.js";

const store = new ModelStore(paths.modelsDir);
const modelPath = store.pathFor(DEFAULT_MODEL_ID);
if (!(await store.isInstalled(DEFAULT_MODEL_ID))) {
  console.error("Model not installed. Run: node scripts/fetch-model.mjs");
  process.exit(1);
}

const sheetContext = {
  workbookName: "Q3 Sales.xlsx",
  sheetName: "Sales",
  sheetNames: ["Sales", "Notes"],
  usedRange: { address: "A1:E501", rowCount: 501, columnCount: 5 },
  selection: { address: "E2:E501", rowCount: 500, columnCount: 1, isSingleCell: false },
  headers: ["Region", "Rep", "Units", "Unit Price", "Revenue"],
  columnLetters: ["A", "B", "C", "D", "E"],
  sample: {
    startRow: 2,
    columnLetters: ["A", "B", "C", "D", "E"],
    rows: [
      ["North", "Ada Chen", 120, 19.99, ""],
      ["South", "Ben Ortiz", 80, 24.5, ""],
      ["East", "Cy Patel", 240, 12.0, ""],
      ["West", "Dee Kim", 15, 99.95, ""],
    ],
  },
};

const engine = new Engine({ config: { contextTokens: 8192, temperature: 0.2, maxTokens: 700 } });

console.log("Loading model...");
const t0 = Date.now();
await engine.load(DEFAULT_MODEL_ID, modelPath);
console.log(`Loaded in ${((Date.now() - t0) / 1000).toFixed(1)}s\n`);

const cases = [
  { name: "formula proposal", q: "Fill the Revenue column with units times unit price." },
  { name: "plain answer", q: "In one sentence, what does the Units column represent?" },
  { name: "formatting", q: "Format the Unit Price column as US dollars with 2 decimals." },
];

let failures = 0;

for (const c of cases) {
  console.log(`--- ${c.name} ---`);
  console.log(`Q: ${c.q}`);
  process.stdout.write("A: ");
  await engine.resetConversation();
  const res = await engine.chat({
    message: c.q,
    sheetContext,
    onToken: (t) => process.stdout.write(t),
  });
  console.log(`\n[${res.stats.tokensPerSecond} tok/s, ${res.stats.seconds}s]`);
  if (res.actions.length) {
    for (const a of res.actions) {
      console.log(`  ACTION ${a.type}: ${a.summary}`);
      if (a.formula) console.log(`    formula: ${a.formula}`);
      if (a.numberFormat) console.log(`    numberFormat: ${a.numberFormat}`);
    }
  } else {
    console.log("  (no actions proposed)");
  }
  if (c.name !== "plain answer" && !res.actions.length) {
    console.log("  !! expected an action");
    failures++;
  }
  console.log();
}

await engine.unload();
console.log(failures ? `SMOKE FAILED (${failures})` : "SMOKE PASSED");
process.exit(failures ? 1 : 0);
