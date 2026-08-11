import { describe, it, expect } from "vitest";
import { splitReadAddress, MAX_READ_CELLS } from "../../taskpane/sheet/reader.js";

describe("splitReadAddress", () => {
  it("passes a plain address through with the given sheet", () => {
    expect(splitReadAddress({ sheet: "Sales", address: "A1:B2" })).toEqual({ sheet: "Sales", ref: "A1:B2" });
    expect(splitReadAddress({ sheet: null, address: "A1" })).toEqual({ sheet: null, ref: "A1" });
  });

  it("absorbs a sheet-qualified address, which wins over the sheet argument", () => {
    expect(splitReadAddress({ sheet: null, address: "Sheet2!A1:B2" })).toEqual({ sheet: "Sheet2", ref: "A1:B2" });
    expect(splitReadAddress({ sheet: "Other", address: "Sheet2!A1" })).toEqual({ sheet: "Sheet2", ref: "A1" });
  });

  it("unquotes sheet names the way Excel writes them", () => {
    expect(splitReadAddress({ sheet: null, address: "'My Sheet'!A1:C3" })).toEqual({ sheet: "My Sheet", ref: "A1:C3" });
    expect(splitReadAddress({ sheet: null, address: "'It''s data'!B2" })).toEqual({ sheet: "It's data", ref: "B2" });
  });

  it("keeps garbage identifiable instead of throwing", () => {
    expect(splitReadAddress({ sheet: null, address: "" })).toEqual({ sheet: null, ref: "" });
    expect(splitReadAddress({ sheet: null, address: undefined })).toEqual({ sheet: null, ref: "" });
  });

  it("keeps the bridge cap where context.js set the precedent", () => {
    // Not a tautology: this pins the pane-side cap so a future edit cannot silently
    // let unbounded reads across the Office bridge.
    expect(MAX_READ_CELLS).toBe(2000);
  });
});
