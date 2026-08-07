import { describe, it, expect, vi } from "vitest";
import {
  batchRows, batchSchema, buildBatchPrompt, parseBatchResult, transformValues,
  DEFAULT_BATCH_SIZE, MAX_TRANSFORM_ROWS, findLossyRows,
} from "../../core/llm/transform.js";

describe("batching", () => {
  it("splits values into fixed-size batches with their start index", () => {
    const batches = batchRows(["a", "b", "c", "d", "e"], 2);
    expect(batches).toEqual([
      { start: 0, rows: ["a", "b"] },
      { start: 2, rows: ["c", "d"] },
      { start: 4, rows: ["e"] },
    ]);
  });

  it("handles an empty input", () => {
    expect(batchRows([], DEFAULT_BATCH_SIZE)).toEqual([]);
  });

  it("pins the schema to an exact result count", () => {
    const schema = batchSchema(7);
    expect(schema.properties.results.minItems).toBe(7);
    expect(schema.properties.results.maxItems).toBe(7);
    expect(schema.required).toContain("results");
  });

  it("numbers the prompt from the real row offset", () => {
    const prompt = buildBatchPrompt("Uppercase it", ["a", "b"], 12);
    expect(prompt).toContain("12. a");
    expect(prompt).toContain("13. b");
    expect(prompt).toContain("exactly 2 results");
    expect(prompt).toContain("Uppercase it");
  });
});

describe("parsing model output", () => {
  it("reads the constrained object form", () => {
    expect(parseBatchResult({ results: ["a", "b"] })).toEqual(["a", "b"]);
  });

  it("accepts a bare array", () => {
    expect(parseBatchResult(["a", "b"])).toEqual(["a", "b"]);
  });

  it("accepts a JSON string and a fenced block", () => {
    expect(parseBatchResult('{"results":["x"]}')).toEqual(["x"]);
    expect(parseBatchResult('```json\n{"results":["y"]}\n```')).toEqual(["y"]);
  });

  it("coerces nulls to empty strings", () => {
    expect(parseBatchResult({ results: ["a", null] })).toEqual(["a", ""]);
  });

  it("returns null for unusable output", () => {
    expect(parseBatchResult("not json at all")).toBeNull();
    expect(parseBatchResult(null)).toBeNull();
    expect(parseBatchResult(42)).toBeNull();
  });
});

describe("transformValues", () => {
  const upper = async (prompt) => {
    const lines = prompt.split("Inputs:\n")[1].split("\n");
    return { results: lines.map((l) => l.replace(/^\d+\.\s?/, "").toUpperCase()) };
  };

  it("transforms every row in order", async () => {
    const { results, failed } = await transformValues({
      values: ["ada", "ben", "cy"],
      instruction: "Uppercase",
      complete: upper,
      batchSize: 2,
    });
    expect(results).toEqual(["ADA", "BEN", "CY"]);
    expect(failed).toEqual([]);
  });

  it("reports progress up to the total", async () => {
    const seen = [];
    await transformValues({
      values: ["a", "b", "c", "d"],
      instruction: "Uppercase",
      complete: upper,
      batchSize: 2,
      onProgress: (p) => seen.push(p.done),
    });
    expect(seen.at(-1)).toBe(4);
    expect(Math.max(...seen)).toBeLessThanOrEqual(4);
  });

  it("never calls the model for blank rows", async () => {
    const complete = vi.fn(upper);
    const { results } = await transformValues({
      values: ["", "  ", ""],
      instruction: "Uppercase",
      complete,
      batchSize: 3,
    });
    expect(complete).not.toHaveBeenCalled();
    expect(results).toEqual(["", "", ""]);
  });

  // The dangerous failure is silent misalignment: row 3's answer landing in row 2 would
  // corrupt data in a way nobody notices. A short batch must be split, never zipped up.
  it("splits a batch when the model returns the wrong number of results", async () => {
    let calls = 0;
    const flaky = async (prompt) => {
      calls++;
      const lines = prompt.split("Inputs:\n")[1].split("\n");
      // Drop a result the first time, behave once the batch is halved.
      if (lines.length === 4) return { results: ["A", "B", "C"] };
      return { results: lines.map((l) => l.replace(/^\d+\.\s?/, "").toUpperCase()) };
    };

    const { results, failed } = await transformValues({
      values: ["a", "b", "c", "d"],
      instruction: "Uppercase",
      complete: flaky,
      batchSize: 4,
    });

    expect(results).toEqual(["A", "B", "C", "D"]);
    expect(failed).toEqual([]);
    expect(calls).toBeGreaterThan(1);
  });

  it("keeps the original value when a single row cannot be transformed", async () => {
    const broken = async () => "unparseable";
    const { results, failed } = await transformValues({
      values: ["keep me"],
      instruction: "Uppercase",
      complete: broken,
      batchSize: 1,
    });
    expect(results).toEqual(["keep me"]);
    expect(failed).toEqual([0]);
  });

  it("isolates one bad row without losing the good ones", async () => {
    const partly = async (prompt) => {
      const lines = prompt.split("Inputs:\n")[1].split("\n");
      const values = lines.map((l) => l.replace(/^\d+\.\s?/, ""));
      if (values.length === 1 && values[0] === "bad") return "garbage";
      if (values.includes("bad")) return { results: [] };
      return { results: values.map((v) => v.toUpperCase()) };
    };

    const { results, failed } = await transformValues({
      values: ["a", "bad", "c", "d"],
      instruction: "Uppercase",
      complete: partly,
      batchSize: 4,
    });

    expect(results).toEqual(["A", "bad", "C", "D"]);
    expect(failed).toEqual([1]);
  });

  it("stops when aborted", async () => {
    const controller = new AbortController();
    const slow = async (prompt) => {
      controller.abort();
      const lines = prompt.split("Inputs:\n")[1].split("\n");
      return { results: lines.map((l) => l.replace(/^\d+\.\s?/, "")) };
    };
    await expect(transformValues({
      values: ["a", "b", "c", "d"],
      instruction: "x",
      complete: slow,
      batchSize: 1,
      signal: controller.signal,
    })).rejects.toThrow(/cancelled/i);
  });

  it("refuses a job beyond the row limit", async () => {
    await expect(transformValues({
      values: new Array(MAX_TRANSFORM_ROWS + 1).fill("x"),
      instruction: "x",
      complete: upper,
    })).rejects.toThrow(/limit is/);
  });
});

describe("findLossyRows", () => {
  // The dangerous split failure is silent: too few output columns and the leftover text
  // is simply discarded, leaving a sheet that looks perfectly fine.
  it("flags a row whose content was dropped", () => {
    const values = ['"RYAN M HUS 326 HILLCREST ST","EL DORADO","KANSAS","67042","UNITED STATES"'];
    const results = [["RYAN M HUS", "EL DORADO", "KANSAS", "67042", "UNITED STATES"]];
    expect(findLossyRows(values, results)).toEqual([0]);
  });

  it("passes a split that keeps everything", () => {
    const values = ['"RYAN M HUS 326 HILLCREST ST","EL DORADO","KANSAS","67042","UNITED STATES"'];
    const results = [["RYAN M HUS", "326 HILLCREST ST", "EL DORADO", "KANSAS", "67042", "UNITED STATES"]];
    expect(findLossyRows(values, results)).toEqual([]);
  });

  it("ignores punctuation and separators the split legitimately removes", () => {
    expect(findLossyRows(['"A, B","C"'], [["A B", "C"]])).toEqual([]);
  });

  it("skips values too short to judge", () => {
    expect(findLossyRows(["ab"], [[""]])).toEqual([]);
  });

  // A rewrite is meant to shed characters — "STREET" to "ST" is the goal, not a fault —
  // so scalar results are deliberately not measured.
  it("ignores rewrite rows, which legitimately get shorter", () => {
    expect(findLossyRows(["123 MAIN STREET APARTMENT 4"], ["123 MAIN ST APT 4"])).toEqual([]);
  });

  it("flags several lossy rows at once", () => {
    const values = ["ALPHA BRAVO CHARLIE DELTA", "ECHO FOXTROT GOLF HOTEL"];
    const results = [["ALPHA"], ["ECHO FOXTROT GOLF HOTEL"]];
    expect(findLossyRows(values, results)).toEqual([0]);
  });
});
