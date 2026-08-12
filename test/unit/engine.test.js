import { describe, it, expect } from "vitest";
import { Engine, MAX_READS_PER_TURN } from "../../core/llm/engine.js";

/**
 * A llama.cpp stand-in just deep enough for Engine.load + Engine.chat. The test plays
 * the model's part: `driver.run` receives the prompt options — including the tool
 * definitions — and can invoke handlers exactly the way the real model would.
 */
function fakeLlama(driver) {
  class LlamaChatSession {
    async prompt(text, opts) { return driver.run(text, opts); }
    async resetChatHistory() {}
    setChatHistory() {}
    async dispose() {}
  }
  return {
    getLlama: async () => ({
      loadModel: async () => ({
        trainContextSize: 8192,
        createContext: async () => ({ getSequence: () => ({}), dispose: async () => {} }),
        dispose: async () => {},
      }),
      getVramState: async () => ({ total: 1, free: 1, used: 0 }),
    }),
    LlamaChatSession,
    defineChatSessionFunction: (def) => def,
  };
}

async function readyEngine(driver) {
  const engine = new Engine({ importLlama: async () => fakeLlama(driver), config: { contextTokens: 8192 } });
  await engine.load("test-model", "/fake/model.gguf");
  return engine;
}

describe("the llama instance", () => {
  it("is created exactly once even when memoryInfo races the model load", async () => {
    // Two lazy `if (!this.#llama)` checks used to pass at the same time, creating two
    // instances — the model belonged to one, grammars were built on the other, and
    // every transform batch failed with an instance mismatch. Single-flight forever.
    let creations = 0;
    const driver = { run: async () => "ok" };
    const base = fakeLlama(driver);
    const mod = {
      ...base,
      getLlama: async (...args) => {
        creations += 1;
        return base.getLlama(...args);
      },
    };
    const engine = new Engine({ importLlama: async () => mod, config: { contextTokens: 8192 } });
    await Promise.all([
      engine.memoryInfo(),
      engine.load("test-model", "/fake/model.gguf"),
      engine.memoryInfo(),
    ]);
    expect(creations).toBe(1);
  });
});

describe("the read_range tool", () => {
  it("is absent when no readSheet callback is provided", async () => {
    const driver = {};
    const engine = await readyEngine(driver);
    let fns;
    driver.run = (_text, opts) => {
      fns = opts.functions;
      return "ok";
    };
    await engine.chat({ message: "hi" });
    expect(fns.read_range).toBeUndefined();
    expect(fns.write_values).toBeDefined();
  });

  it("hands the model a grid of what the pane read", async () => {
    const driver = {};
    const engine = await readyEngine(driver);
    const requests = [];
    const readSheet = async (req) => {
      requests.push(req);
      return { ok: true, sheet: "Sales", address: "A2:B3", startRow: 2, columnLetters: ["A", "B"], rows: [["x", 1], ["y", 2]] };
    };
    let reply;
    driver.run = async (_text, opts) => {
      reply = await opts.functions.read_range.handler({ address: "A2:B3", sheet: "Sales" });
      return "ok";
    };
    await engine.chat({ message: "hi", readSheet });
    expect(requests).toEqual([{ kind: "read", sheet: "Sales", address: "A2:B3" }]);
    expect(reply).toContain("Contents of Sales!A2:B3");
    expect(reply).toContain("row | A | B");
    expect(reply).toContain("2 | x | 1");
    expect(reply).toContain("3 | y | 2");
  });

  it("relays a failure in words the model can act on", async () => {
    const driver = {};
    const engine = await readyEngine(driver);
    let reply;
    driver.run = async (_text, opts) => {
      reply = await opts.functions.read_range.handler({ address: "A1", sheet: "Nope" });
      return "ok";
    };
    await engine.chat({ message: "hi", readSheet: async () => ({ ok: false, error: "there is no sheet named \"Nope\"" }) });
    expect(reply).toMatch(/could not be read: there is no sheet named "Nope"/);
  });

  it("says so plainly when the range is empty", async () => {
    const driver = {};
    const engine = await readyEngine(driver);
    let reply;
    driver.run = async (_text, opts) => {
      reply = await opts.functions.read_range.handler({ address: "K1:K5" });
      return "ok";
    };
    await engine.chat({
      message: "hi",
      readSheet: async () => ({ ok: true, sheet: "Data", address: "K1:K5", startRow: 1, columnLetters: ["K"], rows: [] }),
    });
    expect(reply).toMatch(/K1:K5 is empty/);
  });

  it("cuts the model off after the per-turn read budget", async () => {
    const driver = {};
    const engine = await readyEngine(driver);
    const replies = [];
    driver.run = async (_text, opts) => {
      for (let i = 0; i <= MAX_READS_PER_TURN; i++) {
        replies.push(await opts.functions.read_range.handler({ address: "A1:A2" }));
      }
      return "ok";
    };
    await engine.chat({
      message: "hi",
      readSheet: async () => ({ ok: true, sheet: "S", address: "A1:A2", startRow: 1, columnLetters: ["A"], rows: [[1], [2]] }),
    });
    expect(replies).toHaveLength(MAX_READS_PER_TURN + 1);
    expect(replies.at(-2)).toContain("Contents of");
    expect(replies.at(-1)).toMatch(/Read budget exhausted/);
  });

  it("evaluates aggregates through the calculate tool instead of guessing", async () => {
    const driver = {};
    const engine = await readyEngine(driver);
    const requests = [];
    const readSheet = async (req) => {
      requests.push(req);
      return { ok: true, sheet: "Big", value: 1173.53 };
    };
    let reply;
    let rejected;
    driver.run = async (_text, opts) => {
      rejected = await opts.functions.calculate.handler({ formula: "MAX(B:B)" });
      reply = await opts.functions.calculate.handler({ formula: "=MAX(Big!B2:B20001)" });
      return "ok";
    };
    await engine.chat({ message: "hi", readSheet });
    expect(rejected).toMatch(/must start with "="/);
    expect(requests).toEqual([{ kind: "calc", formula: "=MAX(Big!B2:B20001)", sheet: null }]);
    expect(reply).toBe("=MAX(Big!B2:B20001) evaluates to: 1173.53");
  });

  it("relays a calculation failure in words the model can act on", async () => {
    const driver = {};
    const engine = await readyEngine(driver);
    let reply;
    driver.run = async (_text, opts) => {
      reply = await opts.functions.calculate.handler({ formula: "=MAXX(B:B)" });
      return "ok";
    };
    await engine.chat({ message: "hi", readSheet: async () => ({ ok: false, error: "Excel rejected that formula" }) });
    expect(reply).toMatch(/could not be calculated: Excel rejected that formula/);
  });

  it("counts calculations against the same per-turn budget as reads", async () => {
    const driver = {};
    const engine = await readyEngine(driver);
    const replies = [];
    driver.run = async (_text, opts) => {
      for (let i = 0; i < MAX_READS_PER_TURN; i++) {
        replies.push(await opts.functions.read_range.handler({ address: "A1:A2" }));
      }
      replies.push(await opts.functions.calculate.handler({ formula: "=SUM(A:A)" }));
      return "ok";
    };
    await engine.chat({
      message: "hi",
      readSheet: async () => ({ ok: true, sheet: "S", address: "A1:A2", startRow: 1, columnLetters: ["A"], rows: [[1], [2]] }),
    });
    expect(replies.at(-1)).toMatch(/Read budget exhausted/);
  });

  it("clamps a huge grid to whole rows and says how much is shown", async () => {
    const driver = {};
    const engine = await readyEngine(driver);
    const rows = Array.from({ length: 200 }, (_, i) => [`${"v".repeat(100)}${i}`]);
    let reply;
    driver.run = async (_text, opts) => {
      reply = await opts.functions.read_range.handler({ address: "A1:A200" });
      return "ok";
    };
    await engine.chat({
      message: "hi",
      readSheet: async () => ({ ok: true, sheet: "Big", address: "A1:A200", startRow: 1, columnLetters: ["A"], rows }),
    });
    // A quarter of an 8K-token window at 3.6 chars a token.
    const maxChars = Math.floor(8192 * 3.6 * 0.25);
    expect(reply).toMatch(/showing the first \d+ of 200 rows/);
    expect(reply.length).toBeLessThanOrEqual(maxChars + 200);
    // Whole rows only: every body line still ends with its row's full value.
    const lastLine = reply.trim().split("\n").at(-1);
    expect(lastLine).toMatch(/\| v{100}\d+$/);
  });
});
