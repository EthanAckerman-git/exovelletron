/** Throughput + thinking-suppression diagnostics. Development tool. */
import { getLlama, LlamaChatSession, LlamaLogLevel, QwenChatWrapper } from "node-llama-cpp";
import { paths } from "../core/config.js";
import path from "node:path";

const modelPath = path.join(paths.modelsDir, "Qwen3.5-4B-UD-Q4_K_XL.gguf");

const llama = await getLlama({ logLevel: LlamaLogLevel.warn });
console.log("=== BACKEND ===  gpu:", llama.gpu, " vram:", ((await llama.getVramState()).total / 1e9).toFixed(1), "GB");

const model = await llama.loadModel({ modelPath });
console.log("gpuLayers:", model.gpuLayers, " size:", (model.size / 1e9).toFixed(2), "GB");

/* ---- raw decode throughput, isolated context ---- */
{
  const ctx = await model.createContext({ contextSize: 2048 });
  const seq = ctx.getSequence();
  const tokens = model.tokenize("Write a paragraph about spreadsheets and data analysis.");
  let n = 0;
  const t = Date.now();
  for await (const _ of seq.evaluate(tokens, { temperature: 0 })) {
    if (++n >= 150) break;
  }
  const dt = (Date.now() - t) / 1000;
  console.log(`\n=== RAW THROUGHPUT === ${n} tokens in ${dt.toFixed(2)}s = ${(n / dt).toFixed(1)} tok/s`);
  await ctx.dispose();
}

/* ---- thinking suppression ---- */
async function trial(label, wrapperOpts, promptOpts) {
  const ctx = await model.createContext({ contextSize: 4096 });
  const session = new LlamaChatSession({
    contextSequence: ctx.getSequence(),
    ...(wrapperOpts ? { chatWrapper: new QwenChatWrapper(wrapperOpts) } : {}),
  });
  const seen = new Set();
  let visible = "";
  const t = Date.now();
  await session.prompt("Format column D as US dollars with 2 decimals. Reply in one short sentence.", {
    maxTokens: 300,
    ...promptOpts,
    onResponseChunk: (c) => {
      if (c.type === "segment") seen.add(c.segmentType);
      else if (c.text) visible += c.text;
    },
  });
  const dt = (Date.now() - t) / 1000;
  console.log(`\n--- ${label} ---`);
  console.log(`  segments: ${seen.size ? [...seen].join(", ") : "(none)"}   time: ${dt.toFixed(1)}s`);
  console.log(`  visible (${visible.length} chars): ${JSON.stringify(visible.slice(0, 160))}`);
  await ctx.dispose();
}

await trial("baseline (auto thoughts)", null, {});
await trial("thoughts=discourage", { thoughts: "discourage", variation: "3.5" }, {});
await trial("discourage + thoughtTokens:0", { thoughts: "discourage", variation: "3.5" }, { budgets: { thoughtTokens: 0 } });

process.exit(0);
