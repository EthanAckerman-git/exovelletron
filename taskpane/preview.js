/**
 * Visual harness for the task pane, built from the real render module.
 *
 * Excel's task pane can only be inspected inside Excel, which makes iterating on the
 * transcript slow. This mounts the actual components against sample data so the layout
 * can be checked in a browser. Development only — never referenced by the manifest.
 */
import { el, buildTurn, buildProposal, renderProse } from "./ui/render.js";

const thread = document.getElementById("thread");

function turn(role, text) {
  const t = buildTurn(role);
  if (role === "assistant") renderProse(t.body, text);
  else t.body.textContent = text;
  thread.appendChild(t.turn);
  return t;
}

const noop = { onApply: async () => ({ undo: async () => {} }), onDismiss: () => {} };

turn("user", "Fill the Revenue column with units times unit price.");
const a1 = turn("assistant", "Revenue is units multiplied by unit price, so a formula keeps it correct as the data changes.");
a1.turn.appendChild(buildProposal({
  id: "1",
  type: "write_formula",
  sheet: "Sales",
  address: "E2:E501",
  formula: "=IFERROR(C2*D2,\"\")",
  summary: "Fill 500 cells in Sales!E2:E501",
}, noop));

turn("user", "Add a status column that flags anything under 20 units.");
const a2 = turn("assistant", "Added a `Status` column using a threshold of 20 units.");
a2.turn.appendChild(buildProposal({
  id: "2",
  type: "write_values",
  sheet: "Sales",
  address: "F2:F6",
  values: [["OK"], ["OK"], ["OK"], ["Low"], ["OK"]],
  summary: "Write 5 values to Sales!F2:F6",
}, noop));

const a3 = turn("assistant", "Formatted the prices as US currency.");
a3.turn.appendChild(buildProposal({
  id: "3",
  type: "format_cells",
  sheet: "Sales",
  address: "D2:D501",
  numberFormat: "$#,##0.00",
  summary: "Format Sales!D2:D501",
}, noop));

const err = buildTurn("error");
err.body.textContent = "The model is still loading.";
thread.appendChild(err.turn);

const meta = el("div", "composer__hint");
meta.appendChild(el("span", "mono", "61.3 tok/s · 2.4s"));
thread.appendChild(meta);
