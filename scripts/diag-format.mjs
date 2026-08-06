import { Engine } from "../core/llm/engine.js";
import { ModelStore } from "../core/models/store.js";
import { DEFAULT_MODEL_ID } from "../core/models/catalog.js";
import { paths } from "../core/config.js";
import { ACTION_SPECS, validateAction } from "../core/llm/actions.js";

const store = new ModelStore(paths.modelsDir);
const engine = new Engine({ config: { contextTokens: 8192, temperature: 0.2, maxTokens: 700 } });
engine.on("action", (a) => console.log("  [event action]", a.type, a.summary));
await engine.load(DEFAULT_MODEL_ID, store.pathFor(DEFAULT_MODEL_ID));

// Instrument the validator to see exactly what the model sends and why it's rejected.
const origValidate = ACTION_SPECS.format_cells.validate;
ACTION_SPECS.format_cells.validate = function (args) {
  console.log("  [format_cells RAW ARGS]", JSON.stringify(args));
  try { const r = origValidate.call(this, args); console.log("  [accepted]"); return r; }
  catch (e) { console.log("  [REJECTED]", e.message); throw e; }
};

const sheetContext = {
  workbookName: "Q3 Sales.xlsx", sheetName: "Sales", sheetNames: ["Sales"],
  usedRange: { address: "A1:E501", rowCount: 501, columnCount: 5 },
  selection: { address: "D2:D501", rowCount: 500, columnCount: 1, isSingleCell: false },
  headers: ["Region", "Rep", "Units", "Unit Price", "Revenue"],
  columnLetters: ["A","B","C","D","E"],
  sample: { startRow: 2, columnLetters: ["A","B","C","D","E"], rows: [["North","Ada",120,19.99,""],["South","Ben",80,24.5,""]] },
};

let out = "";
const res = await engine.chat({
  message: "Format the Unit Price column as US dollars with 2 decimals.",
  sheetContext,
  onToken: (t) => { out += t; },
});
console.log("\n=== RESPONSE TEXT ===");
console.log(JSON.stringify(res.text));
console.log("=== STREAMED ===", JSON.stringify(out));
console.log("=== STATS ===", JSON.stringify(res.stats));
console.log("=== ACTIONS ===", res.actions.length);
process.exit(0);
