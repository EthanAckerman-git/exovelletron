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
import { SYSTEM_PROMPT, buildUserTurn } from "./prompt.js";

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

  /**
   * @param {object} opts
   * @param {() => Promise<any>} [opts.importLlama] injectable for tests
   * @param {(query:string) => Promise<string>} [opts.searchWeb] injectable for tests
   * @param {object} opts.config  { contextTokens, temperature, maxTokens, webSearch }
   */
  constructor({ importLlama = () => import("node-llama-cpp"), searchWeb = null, config = {} } = {}) {
    super();
    this.#llamaModule = importLlama;
    this.#searchWeb = searchWeb ?? (async (q) => (await import("../search.js")).searchWeb(q));
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
   */
  #chatWrapperFor(modelId, mod) {
    if (this.#config.thinking) return {};
    if (!/^qwen/i.test(modelId ?? "") || !mod.QwenChatWrapper) return {};
    return {
      chatWrapper: new mod.QwenChatWrapper({
        thoughts: "discourage",
        variation: /qwen3\.5/i.test(modelId) ? "3.5" : "3",
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
  #buildFunctions(defineChatSessionFunction, proposals, rejections, sheetFacts) {
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
   * @returns {Promise<{text:string, actions:object[], stats:object}>}
   */
  async chat({ message, sheetContext, onToken }) {
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
        functions: this.#buildFunctions(mod.defineChatSessionFunction, proposals, rejections, { lastRow: lastRowOf(sheetContext) }),
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
