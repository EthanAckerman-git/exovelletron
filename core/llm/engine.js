/**
 * The inference engine: owns the llama.cpp model, context, and chat session.
 *
 * One model is resident at a time. Loading is idempotent and serialised, so several
 * requests arriving during a slow load all wait on the same promise rather than each
 * trying to map a 3 GB file.
 *
 * Tool calls are *proposals*: the handlers validate and queue an action, then hand the
 * model a short confirmation. Nothing touches the workbook until the user approves the
 * preview in the task pane.
 */
import { EventEmitter } from "node:events";
import os from "node:os";
import { ACTION_SPECS, validateAction, lastRowOf } from "./actions.js";
import { SYSTEM_PROMPT, buildUserTurn, CHARS_PER_TOKEN, renderGrid } from "./prompt.js";
import { getModel } from "../models/catalog.js";

/** @typedef {"idle"|"loading"|"ready"|"generating"|"error"} EngineState */

/**
 * ggml probes for an optional Metal shader variant that uses the Metal 4 tensor API and
 * logs at error level when it cannot build it. It then disables that path and runs the
 * standard Metal kernels normally — all layers still offload to the GPU. Surfacing this
 * to users would be a false alarm, so it is filtered out of the log stream.
 */
const BENIGN_BACKEND_NOISE = [
  /ggml_metal_library_init_from_source/i,
  /tensor API is not supported in this environment/i,
];

export const isBenignBackendNoise = (message) =>
  typeof message === "string" && BENIGN_BACKEND_NOISE.some((re) => re.test(message));

/** How many invalid tool calls to tolerate in one turn before telling the model to stop. */
export const MAX_TOOL_REJECTIONS = 4;

/**
 * How many workbook reads one turn may make. Small models sometimes loop on a tool;
 * the budget turns that failure mode into "answer with what you have" instead of a
 * reply that never arrives.
 */
export const MAX_READS_PER_TURN = 8;

export class Engine extends EventEmitter {
  #llamaModule;
  #llama = null;
  #model = null;
  #context = null;
  #session = null;
  #modelId = null;
  #modelPath = null;
  #state = "idle";
  #error = null;
  #loadPromise = null;
  #generation = null; // { controller }
  #config;

  #searchWeb;
  #fetchPage;

  /**
   * @param {object} opts
   * @param {() => Promise<any>} [opts.importLlama] injectable for tests
   * @param {(query:string) => Promise<string>} [opts.searchWeb] injectable for tests
   * @param {(url:string, opts?:object) => Promise<string>} [opts.fetchPage] injectable for tests
   * @param {object} opts.config  { contextTokens, temperature, maxTokens, webSearch }
   */
  constructor({ importLlama = () => import("node-llama-cpp"), searchWeb = null, fetchPage = null, config = {} } = {}) {
    super();
    this.#llamaModule = importLlama;
    this.#searchWeb = searchWeb ?? (async (q) => (await import("../search.js")).searchWeb(q));
    this.#fetchPage = fetchPage ?? (async (url, opts) => (await import("../fetchpage.js")).fetchPage(url, opts));
    this.#config = { contextTokens: 8192, temperature: 0.3, maxTokens: 1536, thinking: false, webSearch: false, ...config };
  }

  get state() { return this.#state; }
  get modelId() { return this.#modelId; }
  get isBusy() { return this.#generation !== null; }

  #setState(state, error = null) {
    this.#state = state;
    this.#error = error;
    this.emit("state", { state, modelId: this.#modelId, error: error?.message ?? null });
  }

  status() {
    return {
      state: this.#state,
      modelId: this.#modelId,
      error: this.#error?.message ?? null,
      contextTokens: this.#config.contextTokens,
      busy: this.isBusy,
    };
  }

  /**
   * Memory actually available to the model, from the GPU backend rather than os.totalmem.
   * On Apple Silicon the usable working set is well below installed RAM, and quoting
   * installed RAM would overpromise how large a model fits.
   */
  async memoryInfo() {
    try {
      const mod = await this.#llamaModule();
      if (!this.#llama) {
        this.#llama = await mod.getLlama({
          logLevel: mod.LlamaLogLevel?.warn ?? undefined,
          logger: (level, message) => {
            if (isBenignBackendNoise(message)) return;
            this.emit("log", { level, message });
          },
        });
      }
      const vram = await this.#llama.getVramState();
      // While a model is resident its own weights show as used; the honest figure for
      // "could another model fit" is free plus whatever this one already holds.
      return { total: vram.total, free: vram.free, used: vram.used };
    } catch {
      return { total: os.totalmem(), free: os.freemem(), used: 0 };
    }
  }

  updateConfig(patch) {
    this.#config = { ...this.#config, ...patch };
  }

  /**
   * Load a model file. Re-loading the same path is a no-op; a different path swaps.
   * @param {string} modelId
   * @param {string} modelPath
   */
  async load(modelId, modelPath) {
    if (this.#modelPath === modelPath && this.#state === "ready") return;
    if (this.#loadPromise) {
      await this.#loadPromise.catch(() => {});
      if (this.#modelPath === modelPath && this.#state === "ready") return;
    }
    this.#loadPromise = this.#doLoad(modelId, modelPath).finally(() => {
      this.#loadPromise = null;
    });
    return this.#loadPromise;
  }

  async #doLoad(modelId, modelPath) {
    this.#setState("loading");
    try {
      await this.#disposeModel();

      const mod = await this.#llamaModule();
      if (!this.#llama) {
        this.#llama = await mod.getLlama({
          logLevel: mod.LlamaLogLevel?.warn ?? undefined,
          logger: (level, message) => {
            if (isBenignBackendNoise(message)) return;
            this.emit("log", { level, message });
          },
        });
      }

      this.#model = await this.#llama.loadModel({ modelPath });

      // Cap requested context by what the model actually supports.
      const trainCtx = this.#model.trainContextSize ?? this.#config.contextTokens;
      const contextSize = Math.max(2048, Math.min(this.#config.contextTokens, trainCtx));

      this.#context = await this.#model.createContext({
        contextSize,
        threads: Math.max(2, Math.min(8, os.cpus().length - 2)),
      });

      const { LlamaChatSession } = mod;
      this.#session = new LlamaChatSession({
        contextSequence: this.#context.getSequence(),
        systemPrompt: SYSTEM_PROMPT,
        autoDisposeSequence: false,
        ...this.#chatWrapperFor(modelId, mod),
      });

      this.#modelId = modelId;
      this.#modelPath = modelPath;
      this.#config.contextTokens = contextSize;
      this.#setState("ready");
    } catch (err) {
      await this.#disposeModel();
      this.#setState("error", err);
      throw new Error(`Could not load the model: ${err.message}`);
    }
  }

  /**
   * Qwen3.5 is a hybrid reasoning model: left to its own devices it opens a <think>
   * block and can spend the entire token budget in it, returning an empty answer. For
   * spreadsheet work the reasoning buys nothing and costs several seconds a turn, so we
   * discourage thoughts outright. Measured: 4.8s and an empty reply becomes 0.8s and a
   * real one.
   *
   * Non-Qwen models fall through to node-llama-cpp's own template detection.
   *
   * The wrapper choice comes from the catalog entry when the id is known; the regex
   * remains only as a fallback for ids the catalog has never heard of.
   */
  #chatWrapperFor(modelId, mod) {
    if (this.#config.thinking) return {};
    const spec = getModel(modelId ?? "")?.chat;
    const wrapper = spec?.wrapper ?? (/^qwen/i.test(modelId ?? "") ? "qwen" : null);
    if (wrapper !== "qwen" || !mod.QwenChatWrapper) return {};
    return {
      chatWrapper: new mod.QwenChatWrapper({
        thoughts: "discourage",
        variation: spec?.variation ?? (/qwen3\.5/i.test(modelId ?? "") ? "3.5" : "3"),
      }),
    };
  }

  async #disposeModel() {
    try { await this.#session?.dispose?.(); } catch { /* ignore */ }
    try { await this.#context?.dispose?.(); } catch { /* ignore */ }
    try { await this.#model?.dispose?.(); } catch { /* ignore */ }
    this.#session = null;
    this.#context = null;
    this.#model = null;
    this.#modelId = null;
    this.#modelPath = null;
  }

  async unload() {
    this.abort();
    await this.#disposeModel();
    this.#setState("idle");
  }

  /** Clear conversation history but keep the model resident. */
  async resetConversation() {
    if (!this.#session) return;
    await this.#session.resetChatHistory();
  }

  /**
   * Re-prime the model with a saved transcript so a reopened conversation can be
   * continued rather than merely read.
   *
   * Only the plain user/assistant turns are replayed. Tool calls are deliberately left
   * out: their results referred to a workbook state that no longer exists, and replaying
   * them would invite the model to believe changes are still pending.
   */
  async restoreConversation(messages = []) {
    if (!this.#session) return;
    const history = [{ type: "system", text: SYSTEM_PROMPT }];
    for (const message of messages) {
      if (message.role === "user" && message.text?.trim()) {
        history.push({ type: "user", text: message.text });
      } else if (message.role === "assistant" && message.text?.trim()) {
        history.push({ type: "model", response: [message.text] });
      }
    }
    this.#session.setChatHistory(history);
  }

  /**
   * Run a row-by-row transformation over a column of values.
   *
   * This gets its own short-lived context rather than sharing the chat session: the
   * batches would otherwise fill the conversation with hundreds of rows of intermediate
   * data, and the memory is better released as soon as the job finishes.
   *
   * @param {object} opts
   * @param {string[]} opts.values
   * @param {string} opts.instruction
   * @param {(p:{done:number,total:number,sample:string[]}) => void} [opts.onProgress]
   * @param {AbortSignal} [opts.signal]
   */
  async transform({ values, instruction, fields = null, onProgress, signal }) {
    if (this.#state !== "ready" || !this.#model) throw new Error("No model is loaded.");
    if (this.isBusy) throw new Error("The model is busy answering a message.");

    const mod = await this.#llamaModule();
    const controller = new AbortController();
    const onAbort = () => controller.abort();
    signal?.addEventListener("abort", onAbort, { once: true });

    this.#generation = { controller };
    this.#setState("generating");

    let context = null;

    try {
      context = await this.#model.createContext({ contextSize: 4096 });
      const session = new mod.LlamaChatSession({
        contextSequence: context.getSequence(),
        systemPrompt:
          "You transform spreadsheet values. You return only the transformed values, as JSON, " +
          "in the same order and the same number as the inputs. You never add commentary.",
        // A split needs room for several fields per row.
        ...this.#chatWrapperFor(this.#modelId, mod),
      });

      const complete = async (prompt, schema) => {
        const grammar = await this.#llama.createGrammarForJsonSchema(schema);
        // Each batch is independent; resetting stops earlier rows steering later ones.
        await session.resetChatHistory();
        const raw = await session.prompt(prompt, {
          grammar,
          temperature: 0,
          maxTokens: 3072,
          signal: controller.signal,
          stopOnAbortSignal: true,
          budgets: this.#config.thinking ? undefined : { thoughtTokens: 0 },
        });
        return grammar.parse(raw);
      };

      const { transformValues, findLossyRows } = await import("./transform.js");
      const result = await transformValues({
        values,
        instruction,
        fields,
        complete,
        signal: controller.signal,
        // Surface the first few finished rows so the pane can show real output while
        // a long job is still running.
        onProgress: ({ done, total, results }) => {
          onProgress?.({ done, total, sample: results.filter((v) => v != null).slice(0, 5) });
        },
      });

      // A split that drops content leaves a plausible-looking sheet, so say so loudly.
      return fields?.length
        ? { ...result, lossy: findLossyRows(values, result.results) }
        : result;
    } finally {
      signal?.removeEventListener("abort", onAbort);
      try { await context?.dispose?.(); } catch { /* already gone */ }
      this.#generation = null;
      if (this.#state === "generating") this.#setState("ready");
    }
  }

  /**
   * Rebuild a messy range as a table of records.
   *
   * Unlike `transform`, the input is read as one document: the model sees each chunk of
   * lines with its neighbours, so a record whose fields are stacked down several rows —
   * or a cell holding a dozen records — comes out as clean rows. The output row count is
   * untied from the input row count. Same lifecycle as `transform`: a short-lived
   * context, disposed as soon as the job ends.
   *
   * @param {object} opts
   * @param {string[]} opts.values      source rows as text, in order
   * @param {string} opts.instruction
   * @param {string[]} opts.fields      output column names
   * @param {(p:{done:number,total:number,records:number,sample:string[][]}) => void} [opts.onProgress]
   * @param {AbortSignal} [opts.signal]
   */
  async extract({ values, instruction, fields, onProgress, signal }) {
    if (this.#state !== "ready" || !this.#model) throw new Error("No model is loaded.");
    if (this.isBusy) throw new Error("The model is busy answering a message.");

    const mod = await this.#llamaModule();
    const controller = new AbortController();
    const onAbort = () => controller.abort();
    signal?.addEventListener("abort", onAbort, { once: true });

    this.#generation = { controller };
    this.#setState("generating");

    let context = null;

    try {
      // Extraction replies repeat every field name for every record, so they run long;
      // the context is sized for a full chunk in and a full record list out.
      const trainCtx = this.#model.trainContextSize ?? 8192;
      context = await this.#model.createContext({ contextSize: Math.min(8192, trainCtx) });
      const session = new mod.LlamaChatSession({
        contextSequence: context.getSequence(),
        systemPrompt:
          "You extract structured records from messy spreadsheet data. You return only the " +
          "records, as JSON in the requested shape. A field the input does not contain is an " +
          "empty string; you never invent data and never add commentary.",
        ...this.#chatWrapperFor(this.#modelId, mod),
      });

      const complete = async (prompt, schema) => {
        const grammar = await this.#llama.createGrammarForJsonSchema(schema);
        // Each chunk stands alone; resetting stops earlier chunks steering later ones.
        await session.resetChatHistory();
        const raw = await session.prompt(prompt, {
          grammar,
          temperature: 0,
          maxTokens: 4096,
          signal: controller.signal,
          stopOnAbortSignal: true,
          budgets: this.#config.thinking ? undefined : { thoughtTokens: 0 },
        });
        return grammar.parse(raw);
      };

      const { extractRecords } = await import("./extract.js");
      return await extractRecords({
        values,
        instruction,
        fields,
        complete,
        signal: controller.signal,
        onProgress,
      });
    } finally {
      signal?.removeEventListener("abort", onAbort);
      try { await context?.dispose?.(); } catch { /* already gone */ }
      this.#generation = null;
      if (this.#state === "generating") this.#setState("ready");
    }
  }

  abort() {
    this.#generation?.controller.abort();
  }

  /**
   * Build node-llama-cpp function definitions from ACTION_SPECS. Each handler records a
   * validated proposal; a rejected proposal returns its reason so the model can correct
   * itself within the same turn instead of failing silently.
   */
  #buildFunctions(defineChatSessionFunction, proposals, rejections, sheetFacts, readSheet) {
    const fns = {};
    for (const [name, spec] of Object.entries(ACTION_SPECS)) {
      fns[name] = defineChatSessionFunction({
        description: spec.description,
        params: spec.params,
        handler: (args) => {
          const result = validateAction(name, args, sheetFacts);
          if (result.ok) {
            proposals.push(result.action);
            this.emit("action", result.action);
            return `Queued for the user to approve: ${result.action.summary}`;
          }

          // Without a cap, a model that keeps sending the same bad argument retries
          // until it exhausts maxTokens and the user sees an empty reply. After a few
          // attempts, tell it to give up and answer in words instead.
          rejections.count += 1;
          if (rejections.count >= MAX_TOOL_REJECTIONS) {
            return (
              `Rejected: ${result.error}. You have failed ${rejections.count} times — ` +
              `stop calling tools now and explain to the user in plain words what you were trying to do.`
            );
          }
          return `Rejected: ${result.error}. Fix the arguments and try again.`;
        },
      });
    }

    // The workbook read tool. The engine cannot touch Excel itself — the workbook
    // lives in the Office webview — so the handler parks on a callback the chat route
    // wires to the task pane, and the pane answers with the cells. This is what lets
    // the model see the whole workbook instead of only the sample it was handed.
    if (readSheet) {
      const reads = { count: 0 };
      fns.read_range = defineChatSessionFunction({
        description:
          "Read cells from the workbook — any sheet, any range — and get their values back as a " +
          "grid. Use it when the answer needs cells you cannot already see: another sheet, rows " +
          "past the sample, or a range the context does not show. Read the smallest range that " +
          "answers the question.",
        params: {
          type: "object",
          properties: {
            address: { type: "string", description: 'The range in A1 notation, e.g. "A2:D40".' },
            sheet: { type: "string", description: "Worksheet name. Omit for the active sheet." },
          },
          required: ["address"],
        },
        handler: async ({ address, sheet }) => {
          reads.count += 1;
          if (reads.count > MAX_READS_PER_TURN) {
            return "Read budget exhausted for this turn — stop reading and answer from what you have already seen.";
          }
          const req = {
            kind: "read",
            address: String(address ?? "").trim(),
            sheet: typeof sheet === "string" && sheet.trim() ? sheet.trim() : null,
          };
          if (!req.address) return "Rejected: address is required.";

          const result = await readSheet(req);
          if (!result?.ok) {
            return (
              `That range could not be read: ${result?.error || "the workbook did not answer"}. ` +
              "If the sheet name or range was wrong, correct it and try once more; otherwise answer from what you can see."
            );
          }

          const where = [result.sheet ?? req.sheet, result.address || req.address].filter(Boolean).join("!");
          if (!result.rows?.length) return `${where} is empty — no values there.`;

          // A read may only take a quarter of the window; drop whole trailing rows
          // rather than shredding the grid mid-line.
          const maxChars = Math.floor(this.#config.contextTokens * CHARS_PER_TOKEN * 0.25);
          const grid = renderGrid(result.startRow, result.columnLetters, result.rows);
          if (grid.length <= maxChars) return `Contents of ${where}:\n\n${grid}`;

          const lines = grid.split("\n");
          const kept = [lines[0], lines[1]];
          let size = kept[0].length + kept[1].length + 2;
          for (let i = 2; i < lines.length && size + lines[i].length + 1 <= maxChars; i++) {
            kept.push(lines[i]);
            size += lines[i].length + 1;
          }
          const shown = kept.length - 2;
          return (
            `Contents of ${where} (showing the first ${shown} of ${result.rows.length} rows — ` +
            `read a narrower range for the rest):\n\n${kept.join("\n")}`
          );
        },
      });

      // Aggregates are Excel's job. A model cannot scan twenty thousand rows, and a
      // model that reads three of them and extrapolates gives a confident wrong
      // answer — so totals, extremes, and lookups go through one formula that Excel
      // evaluates in a scratch cell the user never sees.
      fns.calculate = defineChatSessionFunction({
        description:
          "Calculate one Excel formula and return its result WITHOUT changing the sheet. The only " +
          "reliable way to answer totals, averages, maximums, counts, or lookups across many rows — " +
          "Excel computes over every row so you never have to see them. Examples: =MAX(Big!B2:B20001), " +
          '=SUMIF(C:C,"West",B:B), =INDEX(Big!A:A,MATCH(MAX(Big!B:B),Big!B:B,0)).',
        params: {
          type: "object",
          properties: {
            formula: { type: "string", description: 'The formula, starting with "=". Qualify ranges with the sheet name.' },
            sheet: { type: "string", description: "Worksheet unqualified references resolve against. Omit for the active sheet." },
          },
          required: ["formula"],
        },
        handler: async ({ formula, sheet }) => {
          reads.count += 1;
          if (reads.count > MAX_READS_PER_TURN) {
            return "Read budget exhausted for this turn — stop and answer from what you have already seen.";
          }
          const req = {
            kind: "calc",
            formula: String(formula ?? "").trim(),
            sheet: typeof sheet === "string" && sheet.trim() ? sheet.trim() : null,
          };
          if (!req.formula.startsWith("=")) return 'Rejected: the formula must start with "=".';

          const result = await readSheet(req);
          if (!result?.ok) {
            return (
              `That could not be calculated: ${result?.error || "the workbook did not answer"}. ` +
              "Fix the formula and try once more, or answer from what you can see."
            );
          }
          return `${req.formula} evaluates to: ${JSON.stringify(result.value)}`;
        },
      });
    }

    // A data tool, not a proposal: results flow straight back into the turn. Only
    // offered when the user has explicitly enabled web access — by default this app
    // sends nothing off the machine.
    if (this.#config.webSearch) {
      fns.search_web = defineChatSessionFunction({
        description:
          "Search the web and get back result titles, URLs, and snippets. Use when the user asks " +
          "about current events, prices, external facts, or anything not in the workbook. " +
          "Summarise what you learn and cite the source name.",
        params: {
          type: "object",
          properties: {
            query: { type: "string", description: "A concise search query." },
          },
          required: ["query"],
        },
        handler: async ({ query }) => {
          this.emit("search", { query });
          return await this.#searchWeb(String(query ?? ""));
        },
      });

      fns.fetch_page = defineChatSessionFunction({
        description:
          "Open one web page and read its text. Use when the user pastes a link, or when a " +
          "search result clearly holds the answer and the snippet is not enough. Quote or " +
          "summarise what the page says and name the site.",
        params: {
          type: "object",
          properties: {
            url: { type: "string", description: "The full http(s) address of the page to open." },
          },
          required: ["url"],
        },
        handler: async ({ url }) => {
          this.emit("fetch", { url: String(url ?? "") });
          // A page may only take a quarter of the window: the sheet, the question,
          // and the answer still have to fit alongside it.
          const maxChars = Math.floor(this.#config.contextTokens * CHARS_PER_TOKEN * 0.25);
          return await this.#fetchPage(String(url ?? ""), { maxChars });
        },
      });
    }
    return fns;
  }

  /**
   * Run one chat turn.
   *
   * @param {object} opts
   * @param {string} opts.message        the user's question
   * @param {object} [opts.sheetContext] worksheet context from the task pane
   * @param {(chunk:string) => void} [opts.onToken] streamed text
   * @param {(req:{sheet:string|null,address:string}) => Promise<object>} [opts.readSheet]
   *        answers a mid-turn workbook read; without it the read_range tool is absent
   * @returns {Promise<{text:string, actions:object[], stats:object}>}
   */
  async chat({ message, sheetContext, onToken, readSheet }) {
    if (this.#state !== "ready" || !this.#session) {
      throw new Error(this.#state === "loading" ? "The model is still loading." : "No model is loaded.");
    }
    if (this.isBusy) throw new Error("Already generating a response.");

    const mod = await this.#llamaModule();
    const controller = new AbortController();
    this.#generation = { controller };
    this.#setState("generating");

    const proposals = [];
    const rejections = { count: 0 };
    const startedAt = Date.now();
    let tokenCount = 0;
    let firstTokenAt = null;

    try {
      const prompt = buildUserTurn(message, sheetContext, this.#config.sheetContextTokens ?? 6000);

      const text = await this.#session.prompt(prompt, {
        temperature: this.#config.temperature,
        maxTokens: this.#config.maxTokens,
        signal: controller.signal,
        stopOnAbortSignal: true,
        // Belt and braces alongside the wrapper's "discourage": even if the model opens
        // a thought block, it cannot spend the reply budget inside it.
        budgets: this.#config.thinking ? undefined : { thoughtTokens: 0 },
        functions: this.#buildFunctions(mod.defineChatSessionFunction, proposals, rejections, { lastRow: lastRowOf(sheetContext) }, readSheet),
        // Counting real tokens (not text chunks) keeps the reported rate honest — a
        // chunk can carry several tokens, or none at all.
        onResponseChunk: (chunk) => {
          tokenCount += chunk.tokens?.length ?? 0;
          if (firstTokenAt === null && chunk.tokens?.length) firstTokenAt = Date.now();
          if (chunk.type === undefined && chunk.text) onToken?.(chunk.text);
        },
      });

      const elapsed = (Date.now() - startedAt) / 1000;
      // Rate is measured from the first token so it reflects decode speed rather than
      // being dragged down by prompt ingestion.
      const decodeSeconds = firstTokenAt ? (Date.now() - firstTokenAt) / 1000 : elapsed;
      return {
        text: (text ?? "").trim(),
        actions: proposals,
        stats: {
          seconds: Number(elapsed.toFixed(2)),
          tokens: tokenCount,
          tokensPerSecond: decodeSeconds > 0.05 ? Number((tokenCount / decodeSeconds).toFixed(1)) : 0,
        },
      };
    } catch (err) {
      if (controller.signal.aborted) {
        return { text: "", actions: proposals, stats: { aborted: true } };
      }
      throw err;
    } finally {
      this.#generation = null;
      if (this.#state === "generating") this.#setState("ready");
    }
  }
}
