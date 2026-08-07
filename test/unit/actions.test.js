import { describe, it, expect } from "vitest";
import {
  parseRange, columnToIndex, indexToColumn, validateAction, stripBlankParams,
  MAX_CELLS_PER_ACTION, ACTION_SPECS, ACTION_NAMES, lastRowOf,
} from "../../core/llm/actions.js";

describe("column letter conversion", () => {
  it("round-trips", () => {
    for (const [letters, index] of [["A", 1], ["Z", 26], ["AA", 27], ["AZ", 52], ["BA", 53], ["XFD", 16384]]) {
      expect(columnToIndex(letters)).toBe(index);
      expect(indexToColumn(index)).toBe(letters);
    }
  });
});

describe("parseRange", () => {
  it("parses a single cell", () => {
    const r = parseRange("B2");
    expect(r).toMatchObject({ sheet: null, address: "B2:B2", rows: 1, cols: 1 });
  });

  it("parses a range with a sheet name", () => {
    const r = parseRange("Sales!A1:D20");
    expect(r).toMatchObject({ sheet: "Sales", address: "A1:D20", rows: 20, cols: 4 });
  });

  it("handles quoted sheet names containing spaces", () => {
    expect(parseRange("'Q3 Data'!A1:B2").sheet).toBe("Q3 Data");
  });

  it("strips absolute-reference dollar signs", () => {
    expect(parseRange("$C$5").address).toBe("C5:C5");
  });

  it("normalises reversed corners", () => {
    expect(parseRange("D20:A1").address).toBe("A1:D20");
  });

  it.each(["", "   ", "notacell", "A0", "1A", "A1:", "!A1", "AAAA1"])("rejects %o", (bad) => {
    expect(() => parseRange(bad)).toThrow();
  });

  it("rejects rows beyond the worksheet", () => {
    expect(() => parseRange("A1:A2000000")).toThrow(/outside the worksheet/);
  });
});

describe("stripBlankParams", () => {
  // Regression: the model emits every schema property, so blank optionals arrived as "".
  // Rejecting them sent it into a retry loop that consumed the whole token budget.
  it("drops blank and nullish top-level values", () => {
    expect(stripBlankParams({ a: "", b: "  ", c: null, d: undefined, e: "x", f: 0, g: false }))
      .toEqual({ e: "x", f: 0, g: false });
  });

  it("leaves nested empty strings alone", () => {
    const values = [["a", ""], ["", "b"]];
    expect(stripBlankParams({ values }).values).toEqual(values);
  });
});

describe("validateAction: write_formula", () => {
  it("accepts a formula and summarises the fill", () => {
    const r = validateAction("write_formula", { address: "E2:E501", formula: "=C2*D2", sheet: "Sales" });
    expect(r.ok).toBe(true);
    expect(r.action).toMatchObject({ type: "write_formula", sheet: "Sales", address: "E2:E501", formula: "=C2*D2" });
    expect(r.action.summary).toMatch(/Fill 500 cells/);
  });

  it("requires a leading =", () => {
    expect(validateAction("write_formula", { address: "A1", formula: "C2*D2" }))
      .toMatchObject({ ok: false, error: expect.stringMatching(/must start with/) });
  });

  it("refuses to fill more than the cell budget", () => {
    const r = validateAction("write_formula", { address: "A1:A60000", formula: "=1" });
    expect(r).toMatchObject({ ok: false, error: expect.stringContaining(String(MAX_CELLS_PER_ACTION)) });
  });

  // Regression: asked to fill Revenue for six rows of data, the model proposed E2:E1000,
  // which would have written formulas over 994 empty rows.
  it("trims a fill that runs past the last row of data", () => {
    const r = validateAction("write_formula", { address: "E2:E1000", formula: "=C2*D2" }, { lastRow: 6 });
    expect(r.ok).toBe(true);
    expect(r.action.address).toBe("E2:E6");
    expect(r.action.clamped).toBe(true);
    expect(r.action.summary).toMatch(/Fill 5 cells/);
    expect(r.action.summary).toMatch(/trimmed to the rows that contain data/);
  });

  it("leaves a fill that already fits the data alone", () => {
    const r = validateAction("write_formula", { address: "E2:E6", formula: "=C2*D2" }, { lastRow: 6 });
    expect(r.action.address).toBe("E2:E6");
    expect(r.action.clamped).toBe(false);
    expect(r.action.summary).not.toMatch(/trimmed/);
  });

  it("does not trim when the whole range sits below the data", () => {
    // Writing into fresh rows past the end is a legitimate request, not padding.
    const r = validateAction("write_formula", { address: "E20:E30", formula: "=1" }, { lastRow: 6 });
    expect(r.action.address).toBe("E20:E30");
    expect(r.action.clamped).toBe(false);
  });

  it("ignores the bound when the used range is unknown", () => {
    const r = validateAction("write_formula", { address: "E2:E1000", formula: "=1" }, {});
    expect(r.action.address).toBe("E2:E1000");
  });
});

describe("validateAction: split_column", () => {
  const base = { source: "A2:A21", target: "B2", columns: ["Name", "Street", "City", "State", "ZIP"], instruction: "Pull the fields out of each CSV line." };

  it("builds a grid target from the column list", () => {
    const r = validateAction("split_column", base);
    expect(r.ok).toBe(true);
    expect(r.action).toMatchObject({ source: "A2:A21", address: "B2:F21", rows: 20 });
    expect(r.action.columns).toHaveLength(5);
  });

  it("puts the headers in the row above the data", () => {
    expect(validateAction("split_column", base).action.headerAddress).toBe("B1:F1");
  });

  it("omits headers when there is no room above", () => {
    const r = validateAction("split_column", { ...base, source: "A1:A5", target: "B1" });
    expect(r.action.headerAddress).toBeNull();
  });

  // Writing over the column being read would destroy the input halfway through.
  it("refuses a target that overlaps the source column", () => {
    expect(validateAction("split_column", { ...base, target: "A2" }))
      .toMatchObject({ ok: false, error: expect.stringMatching(/overwrite the source column A/) });
  });

  it("needs at least two output columns", () => {
    expect(validateAction("split_column", { ...base, columns: ["Everything"] }))
      .toMatchObject({ ok: false, error: expect.stringMatching(/at least two/) });
  });

  it("drops blank column names before counting", () => {
    expect(validateAction("split_column", { ...base, columns: ["Name", "  ", "City"] }).action.columns)
      .toEqual(["Name", "City"]);
  });

  it("trims a source that runs past the data", () => {
    const r = validateAction("split_column", { ...base, source: "A2:A1000" }, { lastRow: 21 });
    expect(r.action.source).toBe("A2:A21");
    expect(r.action.clamped).toBe(true);
  });

  it("refuses a job that would write more cells than the budget", () => {
    const r = validateAction("split_column", { ...base, source: "A2:A5000", columns: new Array(20).fill("c").map((c, i) => c + i) });
    expect(r).toMatchObject({ ok: false, error: expect.stringContaining("limit is") });
  });

  it("requires a real instruction", () => {
    expect(validateAction("split_column", { ...base, instruction: "split" }).ok).toBe(false);
  });
});

describe("validateAction: transform_column", () => {
  const base = { source: "B2:B500", target: "C2", instruction: "Standardise this address to USPS format." };

  it("accepts a whole-column rewrite", () => {
    const r = validateAction("split_column", { ...base, columns: ["a", "b"] });
    expect(r.ok).toBe(true);
  });

  it("rewrites in place when source and target match", () => {
    const r = validateAction("transform_column", { source: "B2:B10", target: "B2:B10", instruction: "Uppercase every value." });
    expect(r.action.inPlace).toBe(true);
    expect(r.action.rows).toBe(9);
  });

  it("works on a single column only", () => {
    expect(validateAction("transform_column", { ...base, source: "A2:C9" }).ok).toBe(false);
  });
});

describe("lastRowOf", () => {
  it("reads the final row from a used range", () => {
    expect(lastRowOf({ usedRange: { address: "A1:E501" } })).toBe(501);
    expect(lastRowOf({ usedRange: { address: "Sheet1!B2:D9" } })).toBe(9);
  });

  it("returns null when there is nothing usable", () => {
    expect(lastRowOf(null)).toBeNull();
    expect(lastRowOf({})).toBeNull();
    expect(lastRowOf({ usedRange: { address: "garbage" } })).toBeNull();
  });
});

describe("validateAction: write_values", () => {
  it("accepts a grid matching the range", () => {
    const r = validateAction("write_values", { address: "A1:B2", values: [["a", "b"], ["c", "d"]] });
    expect(r.ok).toBe(true);
    expect(r.action.values).toEqual([["a", "b"], ["c", "d"]]);
  });

  it("rejects a grid whose shape disagrees with the range", () => {
    expect(validateAction("write_values", { address: "A1:B2", values: [["a", "b"]] }))
      .toMatchObject({ ok: false, error: expect.stringMatching(/1x2 but 2x2/) });
  });

  it("rejects ragged rows", () => {
    expect(validateAction("write_values", { address: "A1:B2", values: [["a", "b"], ["c"]] }))
      .toMatchObject({ ok: false, error: expect.stringMatching(/same length/) });
  });

  it("coerces null cells to empty strings", () => {
    const r = validateAction("write_values", { address: "A1:A2", values: [[null], ["x"]] });
    expect(r.action.values).toEqual([[""], ["x"]]);
  });
});

describe("validateAction: format_cells", () => {
  // Regression: `fill: ""` used to be a hard error rather than "not specified".
  it("treats a blank fill as unspecified", () => {
    const r = validateAction("format_cells", { address: "D2:D501", numberFormat: "$#,##0.00", fill: "" });
    expect(r.ok).toBe(true);
    expect(r.action.fill).toBeUndefined();
    expect(r.action.numberFormat).toBe("$#,##0.00");
  });

  it("normalises hex colours", () => {
    expect(validateAction("format_cells", { address: "A1", fill: "fff3cd" }).action.fill).toBe("#FFF3CD");
  });

  it("rejects a malformed colour that was actually supplied", () => {
    expect(validateAction("format_cells", { address: "A1", fill: "reddish" }))
      .toMatchObject({ ok: false, error: expect.stringMatching(/hex colour/) });
  });

  it("requires at least one property to change", () => {
    expect(validateAction("format_cells", { address: "A1" }))
      .toMatchObject({ ok: false, error: expect.stringMatching(/at least one/) });
  });
});

describe("validateAction: sort_range", () => {
  it("computes the column offset within the range", () => {
    const r = validateAction("sort_range", { address: "B2:F100", byColumn: "D", ascending: false });
    expect(r.action).toMatchObject({ offset: 2, ascending: false });
  });

  it("rejects a sort column outside the range", () => {
    expect(validateAction("sort_range", { address: "B2:D100", byColumn: "Z" }))
      .toMatchObject({ ok: false, error: expect.stringMatching(/outside the range/) });
  });

  it("defaults to ascending", () => {
    expect(validateAction("sort_range", { address: "A1:B2", byColumn: "A" }).action.ascending).toBe(true);
  });
});

describe("validateAction: extract_table", () => {
  const base = {
    source: "A2:A40",
    target: "C2",
    columns: ["Name", "Street", "City", "State", "ZIP"],
    instruction: "Each record is one person's mailing address; fields may span several rows.",
  };

  it("accepts a messy range and anchors the output at the target cell", () => {
    const { action } = validateAction("extract_table", base);
    expect(action).toMatchObject({
      type: "extract_table",
      source: "A2:A40",
      address: "C2",
      targetCol: 3,
      targetRow: 2,
      headerAddress: "C1:G1",
      rows: 39,
    });
  });

  it("clamps the source to the rows that actually hold data", () => {
    const { action } = validateAction("extract_table", { ...base, source: "A2:A1000" }, { lastRow: 21 });
    expect(action.source).toBe("A2:A21");
    expect(action.rows).toBe(20);
    expect(action.clamped).toBe(true);
  });

  it("refuses a target that would overwrite the source while it is being read", () => {
    const result = validateAction("extract_table", { ...base, target: "A2" });
    expect(result.ok).toBe(false);
    expect(result.error).toMatch(/overwrite the source/);
  });

  it("writes no header row when the table starts on row 1", () => {
    expect(validateAction("extract_table", { ...base, target: "C1" }).action.headerAddress).toBeNull();
  });

  it("requires at least one named column and a real instruction", () => {
    expect(validateAction("extract_table", { ...base, columns: ["", "  "] }).ok).toBe(false);
    expect(validateAction("extract_table", { ...base, instruction: "do it" }).ok).toBe(false);
  });
});

describe("validateAction: insert_column", () => {
  it("accepts a column letter and header", () => {
    expect(validateAction("insert_column", { before: "d", header: "Margin" }).action)
      .toMatchObject({ before: "D", header: "Margin" });
  });

  it("rejects a non-letter", () => {
    expect(validateAction("insert_column", { before: "3" }).ok).toBe(false);
  });
});

describe("apply coverage", () => {
  // Regression: split_column was added to the tool list but not to the task pane's
  // routing table, so the model proposed it and applying failed with
  // "Cannot apply unsupported action".
  it("every action the model can propose has an apply path", async () => {
    const { readFile } = await import("node:fs/promises");
    const [applySource, mainSource] = await Promise.all([
      readFile(new URL("../../taskpane/sheet/apply.js", import.meta.url), "utf8"),
      readFile(new URL("../../taskpane/main.js", import.meta.url), "utf8"),
    ]);

    for (const name of ACTION_NAMES) {
      const handled = applySource.includes(`case "${name}"`) || mainSource.includes(`"${name}"`);
      expect(handled, `${name} has no apply path`).toBe(true);
    }
  });
});

describe("action specs", () => {
  it("rejects unknown tools", () => {
    expect(validateAction("drop_table", {})).toMatchObject({ ok: false });
  });

  it("every spec declares a description, schema, and validator", () => {
    for (const name of ACTION_NAMES) {
      const spec = ACTION_SPECS[name];
      expect(spec.description.length).toBeGreaterThan(20);
      expect(spec.params.type).toBe("object");
      expect(Array.isArray(spec.params.required)).toBe(true);
      expect(typeof spec.validate).toBe("function");
      for (const key of spec.params.required) {
        expect(spec.params.properties).toHaveProperty(key);
      }
    }
  });

  it("gives every accepted action a stable id and human summary", () => {
    const r = validateAction("write_formula", { address: "A1", formula: "=1" });
    expect(r.action.id).toMatch(/^write_formula-/);
    expect(r.action.summary).toBeTruthy();
  });
});
