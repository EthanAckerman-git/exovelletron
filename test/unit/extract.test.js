import { describe, it, expect, vi } from "vitest";
import {
  chunkLines,
  extractSchema,
  buildExtractPrompt,
  parseExtractResult,
  extractRecords,
  MAX_EXTRACT_ROWS,
} from "../../core/llm/extract.js";

describe("chunkLines", () => {
  it("keeps a small input in one chunk", () => {
    const chunks = chunkLines(["a", "b", "c"]);
    expect(chunks).toHaveLength(1);
    expect(chunks[0]).toEqual({ start: 0, lines: ["a", "b", "c"] });
  });

  it("cuts at a blank line so a stacked record is never severed", () => {
    // Two vertical records separated by a blank line, with a budget that forces a cut
    // partway through the second record. The cut must land on the blank line.
    const lines = ["RYAN M HUS", "326 HILLCREST ST", "EL DORADO", "", "DALTON BROKEY", "2355 192ND RD"];
    const budget = "RYAN M HUS 326 HILLCREST ST EL DORADO  DALTON".length;
    const chunks = chunkLines(lines, budget);
    expect(chunks[0].lines).toEqual(["RYAN M HUS", "326 HILLCREST ST", "EL DORADO", ""]);
    expect(chunks[1]).toEqual({ start: 4, lines: ["DALTON BROKEY", "2355 192ND RD"] });
  });

  it("covers every line exactly once, in order", () => {
    const lines = Array.from({ length: 50 }, (_, i) => `line ${i} ${"x".repeat(i % 7)}`);
    const chunks = chunkLines(lines, 80);
    const flat = chunks.flatMap((c) => c.lines);
    expect(flat).toEqual(lines);
    let next = 0;
    for (const c of chunks) {
      expect(c.start).toBe(next);
      next += c.lines.length;
    }
  });

  it("gives an oversized single line its own chunk rather than dropping it", () => {
    const giant = "KANSAS,66534,UNITED STATES,".repeat(50);
    const chunks = chunkLines(["a", giant, "b"], 100);
    expect(chunks.flatMap((c) => c.lines)).toEqual(["a", giant, "b"]);
    expect(chunks.some((c) => c.lines.length === 1 && c.lines[0] === giant)).toBe(true);
  });
});

describe("extractSchema", () => {
  it("requires every field on every record", () => {
    const schema = extractSchema(["Name", "ZIP"]);
    expect(schema.properties.records.items.required).toEqual(["Name", "ZIP"]);
    expect(Object.keys(schema.properties.records.items.properties)).toEqual(["Name", "ZIP"]);
    expect(schema.properties.records.maxItems).toBeGreaterThan(0);
  });
});

describe("buildExtractPrompt", () => {
  it("numbers the lines and names the fields", () => {
    const prompt = buildExtractPrompt("One record per person.", ["Name", "City"], ["alpha", "beta"]);
    expect(prompt).toContain("One record per person.");
    expect(prompt).toContain("Name, City");
    expect(prompt).toContain("1. alpha");
    expect(prompt).toContain("2. beta");
  });
});

describe("parseExtractResult", () => {
  const fields = ["Name", "ZIP"];

  it("maps record objects onto field order", () => {
    expect(parseExtractResult({ records: [{ Name: "RYAN", ZIP: "67042" }] }, fields))
      .toEqual([["RYAN", "67042"]]);
  });

  it("fills missing fields with empty strings", () => {
    expect(parseExtractResult({ records: [{ Name: "RYAN" }] }, fields)).toEqual([["RYAN", ""]]);
  });

  it("drops records where every field is blank", () => {
    expect(parseExtractResult({ records: [{ Name: "", ZIP: "" }, { Name: "A", ZIP: "" }] }, fields))
      .toEqual([["A", ""]]);
  });

  it("tolerates positional arrays and fenced JSON", () => {
    expect(parseExtractResult([["A", "1"]], fields)).toEqual([["A", "1"]]);
    expect(parseExtractResult('```json\n{"records":[{"Name":"B","ZIP":"2"}]}\n```', fields))
      .toEqual([["B", "2"]]);
  });

  it("returns null for garbage", () => {
    expect(parseExtractResult("not json", fields)).toBeNull();
    expect(parseExtractResult({ records: [42] }, fields)).toBeNull();
    expect(parseExtractResult(undefined, fields)).toBeNull();
  });
});

describe("extractRecords", () => {
  const fields = ["Name"];

  it("concatenates records across chunks and reports progress with a record count", async () => {
    const progress = [];
    const complete = vi.fn(async (prompt) => ({
      // One record per input line mentioned in the prompt.
      records: prompt.split("\n").filter((l) => /^\d+\. \S/.test(l)).map((l) => ({ Name: l.split(". ")[1] })),
    }));

    const { records, failed } = await extractRecords({
      values: ["aaa", "bbb", "", "ccc"],
      instruction: "one per line",
      fields,
      complete,
      onProgress: (p) => progress.push(p),
      chunkBudget: 8, // force several chunks
    });

    expect(records).toEqual([["aaa"], ["bbb"], ["ccc"]]);
    expect(failed).toEqual([]);
    const last = progress.at(-1);
    expect(last.done).toBe(4);
    expect(last.total).toBe(4);
    expect(last.records).toBe(3);
  });

  it("never calls the model for blank lines", async () => {
    const complete = vi.fn();
    const { records } = await extractRecords({ values: ["", "  ", ""], instruction: "x", fields, complete });
    expect(records).toEqual([]);
    expect(complete).not.toHaveBeenCalled();
  });

  it("halves a chunk the model fumbles instead of losing the whole span", async () => {
    const complete = vi.fn(async (prompt) => {
      const lines = prompt.split("\n").filter((l) => /^\d+\. \S/.test(l));
      if (lines.length > 1) return "garbage";
      return { records: [{ Name: lines[0].split(". ")[1].toUpperCase() }] };
    });

    const { records, failed } = await extractRecords({
      values: ["a", "b"], instruction: "x", fields, complete,
    });
    expect(records).toEqual([["A"], ["B"]]);
    expect(failed).toEqual([]);
  });

  it("records the line index when even a lone line cannot be parsed", async () => {
    const complete = vi.fn(async () => "garbage");
    const { records, failed } = await extractRecords({
      values: ["only"], instruction: "x", fields, complete,
    });
    expect(records).toEqual([]);
    expect(failed).toEqual([0]);
  });

  it("stops on abort", async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(extractRecords({
      values: ["a"], instruction: "x", fields, complete: vi.fn(), signal: controller.signal,
    })).rejects.toMatchObject({ code: "ABORTED" });
  });

  it("refuses absurdly large jobs", async () => {
    await expect(extractRecords({
      values: new Array(MAX_EXTRACT_ROWS + 1).fill("x"), instruction: "x", fields, complete: vi.fn(),
    })).rejects.toThrow(/limit/);
  });
});
