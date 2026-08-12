import { describe, it, expect } from "vitest";
import { gridsMatch, UNDO_CHANGED_MESSAGE } from "../../taskpane/sheet/apply.js";

/**
 * The comparison every undo runs before restoring: the cells must still hold exactly
 * what the apply wrote. String coercion mirrors how Excel hands formulas back — a
 * numeric literal reads as a number, but the written value may have been a string.
 */
describe("gridsMatch", () => {
  it("accepts an unchanged grid, across number/string representation", () => {
    expect(gridsMatch([["=B2*C2"], ["=B3*C3"]], [["=B2*C2"], ["=B3*C3"]])).toBe(true);
    expect(gridsMatch([[796, "x"]], ["796", "x"].map((v) => v) && [[796, "x"]])).toBe(true);
    expect(gridsMatch([["796"]], [[796]])).toBe(true);
    expect(gridsMatch([[""]], [[""]])).toBe(true);
  });

  it("rejects any edit, resize, or clearing", () => {
    expect(gridsMatch([["a"]], [["b"]])).toBe(false);
    expect(gridsMatch([["a"], ["b"]], [["a"]])).toBe(false);
    expect(gridsMatch([["a", "b"]], [["a"]])).toBe(false);
    expect(gridsMatch([[""]], [["typed later"]])).toBe(false);
    expect(gridsMatch(null, [["a"]])).toBe(false);
  });

  it("keeps the refusal message honest about what happened", () => {
    expect(UNDO_CHANGED_MESSAGE).toMatch(/nothing was touched/);
  });
});
