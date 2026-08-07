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

describe("validateAction: insert_column", () => {
  it("accepts a column letter and header", () => {
    expect(validateAction("insert_column", { before: "d", header: "Margin" }).action)
      .toMatchObject({ before: "D", header: "Margin" });
  });

  it("rejects a non-letter", () => {
    expect(validateAction("insert_column", { before: "3" }).ok).toBe(false);
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
