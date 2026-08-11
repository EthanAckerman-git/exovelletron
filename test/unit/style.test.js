import { describe, it, expect } from "vitest";
import { headerStyleFrom, headerReferenceCell, DEFAULT_HEADER_STYLE } from "../../taskpane/sheet/style.js";

describe("headerReferenceCell", () => {
  it("points at the cell left of the header row", () => {
    expect(headerReferenceCell("B1:F1")).toBe("A1");
    expect(headerReferenceCell("D3:E3")).toBe("C3");
    expect(headerReferenceCell("AA1:AB1")).toBe("Z1");
  });

  it("returns null in column A and for garbage", () => {
    expect(headerReferenceCell("A1:C1")).toBeNull();
    expect(headerReferenceCell(null)).toBeNull();
    expect(headerReferenceCell("")).toBeNull();
  });
});

describe("headerStyleFrom", () => {
  // The point of the feature: a sheet with navy headers gets navy continued, not a
  // house style stamped over it.
  it("continues an existing header's look", () => {
    expect(headerStyleFrom({ fill: "#1F3864", fontColor: "#FFFFFF", bold: true }))
      .toEqual({ fill: "#1F3864", fontColor: "#FFFFFF", bold: true });
  });

  it("keeps a non-bold reference non-bold only when it says so explicitly", () => {
    expect(headerStyleFrom({ fill: "#333333", fontColor: "#EEEEEE", bold: false }).bold).toBe(false);
    expect(headerStyleFrom({ fill: "#333333" }).bold).toBe(true);
  });

  it("falls back to the coloured default when the neighbour is unstyled", () => {
    expect(headerStyleFrom({ fill: "#FFFFFF", fontColor: "#000000", bold: false })).toEqual(DEFAULT_HEADER_STYLE);
    expect(headerStyleFrom(null)).toEqual(DEFAULT_HEADER_STYLE);
    expect(headerStyleFrom({})).toEqual(DEFAULT_HEADER_STYLE);
  });

  it("has a default that is actually coloured, with readable text", () => {
    expect(DEFAULT_HEADER_STYLE.fill).not.toMatch(/^#?FFFFFF$/i);
    expect(DEFAULT_HEADER_STYLE.bold).toBe(true);
    expect(DEFAULT_HEADER_STYLE.fontColor).toBe("#FFFFFF");
  });
});

describe("Excel's ⓘ badge clearance", () => {
  // Excel draws an un-removable sideloaded-add-in badge over the pane's top-right.
  // Regression: the status strip reserved space for it but the history panel's head
  // did not, so the badge sat exactly on the panel's close button.
  it("reserves the badge corner on every top-of-pane header", async () => {
    const { readFile } = await import("node:fs/promises");
    const css = await readFile(new URL("../../taskpane/styles.css", import.meta.url), "utf8");

    const shared = /\.status,\s*\.history__head\s*\{\s*padding-right:\s*var\(--badge-clearance\)/.exec(css);
    expect(shared, "shared badge-clearance rule").toBeTruthy();
    expect(css).toContain("--badge-clearance: 44px");

    // Neither header may re-set padding with the shorthand, which would silently
    // override the shared right-padding reservation.
    for (const header of ["status", "history__head"]) {
      const block = new RegExp(`\\.${header}\\s*\\{[^}]*\\}`, "g");
      for (const m of css.match(block) ?? []) {
        expect(m, `.${header} must not use the padding shorthand`).not.toMatch(/[{;]\s*padding:/);
      }
    }
  });
});
