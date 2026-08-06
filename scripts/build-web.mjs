/**
 * Builds the browser-side assets the local server hands to Excel.
 *
 * The task pane is bundled so the Office webview loads one script instead of a
 * waterfall of ES module requests, which is noticeably faster on cold open.
 */
import { build } from "esbuild";
import { readFile, writeFile, mkdir, cp } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DIST = path.join(ROOT, "dist");
const dev = process.argv.includes("--dev");

async function bundleCss(sources, outFile) {
  const parts = [];
  for (const src of sources) parts.push(await readFile(path.join(ROOT, src), "utf8"));
  await writeFile(outFile, parts.join("\n"), "utf8");
}

async function main() {
  await mkdir(path.join(DIST, "app"), { recursive: true });

  // Development-only visual harness for the transcript components.
  if (dev) {
    await build({
      entryPoints: [path.join(ROOT, "taskpane", "preview.js")],
      outfile: path.join(DIST, "app", "preview.js"),
      bundle: true, format: "esm", target: ["safari16"], logLevel: "warning",
    });
    await cp(path.join(ROOT, "taskpane", "preview.html"), path.join(DIST, "app", "preview.html"));
  }

  await build({
    entryPoints: [path.join(ROOT, "taskpane", "main.js")],
    outfile: path.join(DIST, "app", "taskpane.js"),
    bundle: true,
    format: "esm",
    target: ["safari16"],
    minify: !dev,
    sourcemap: dev,
    // Office.js is loaded by its own <script> tag and lives on window.
    external: [],
    logLevel: "warning",
  });

  await bundleCss(["shared/tokens.css", "taskpane/styles.css"], path.join(DIST, "app", "taskpane.css"));

  for (const [from, to] of [
    ["taskpane/index.html", "taskpane.html"],
    ["taskpane/commands.html", "commands.html"],
    ["taskpane/landing.html", "landing.html"],
  ]) {
    const src = path.join(ROOT, from);
    if (existsSync(src)) await cp(src, path.join(DIST, to));
  }

  console.log(`web assets -> dist/ ${dev ? "(dev)" : "(minified)"}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
