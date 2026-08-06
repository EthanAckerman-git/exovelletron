import { describe, it, expect } from "vitest";
import { offsetFormula } from "../../taskpane/ui/formula.js";

describe("offsetFormula", () => {
  it("shifts relative row references", () => {
    expect(offsetFormula("=C2*D2", 1)).toBe("=C3*D3");
    expect(offsetFormula("=C2*D2", 5)).toBe("=C7*D7");
  });

  it("leaves absolute rows alone", () => {
    expect(offsetFormula("=C$2*D2", 1)).toBe("=C$2*D3");
    expect(offsetFormula("=$C$2*D2", 3)).toBe("=$C$2*D5");
  });

  it("shifts columns when asked", () => {
    expect(offsetFormula("=A1", 0, 1)).toBe("=B1");
    expect(offsetFormula("=Z1", 0, 1)).toBe("=AA1");
    expect(offsetFormula("=$A1", 0, 1)).toBe("=$A1");
  });

  it("does not mistake function names ending in digits for references", () => {
    expect(offsetFormula("=LOG10(A2)", 1)).toBe("=LOG10(A3)");
    expect(offsetFormula("=SUM(A2:A10)", 1)).toBe("=SUM(A3:A11)");
  });

  it("leaves quoted string literals untouched", () => {
    expect(offsetFormula('=IF(A2>0,"A2 is big","")', 1)).toBe('=IF(A3>0,"A2 is big","")');
    expect(offsetFormula('=CONCAT(A2,"row B2")', 2)).toBe('=CONCAT(A4,"row B2")');
  });

  it("handles escaped quotes inside literals", () => {
    expect(offsetFormula('=IF(A2,"say ""B2""","")', 1)).toBe('=IF(A3,"say ""B2""","")');
  });

  it("returns #REF! when a shift falls off the sheet", () => {
    expect(offsetFormula("=A1", -1)).toBe("=#REF!");
  });

  it("is a no-op for zero delta or non-strings", () => {
    expect(offsetFormula("=A1", 0)).toBe("=A1");
    expect(offsetFormula(undefined, 1)).toBe(undefined);
  });

  it("handles a realistic multi-reference formula", () => {
    expect(offsetFormula('=IFERROR(VLOOKUP(A2,$F$2:$G$99,2,FALSE),"not found")', 3))
      .toBe('=IFERROR(VLOOKUP(A5,$F$2:$G$99,2,FALSE),"not found")');
  });
});
