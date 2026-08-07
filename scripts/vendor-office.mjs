/**
 * Vendors the Office.js runtime into dist/vendor/office so the task pane can run
 * with no network access at all.
 *
 * office.js resolves every dependent script (host bundle, locale strings, telemetry
 * sink) relative to its own <script src>. So mirroring the subset of the CDN layout
 * we need into our own origin is enough to make it work fully offline.
 *
 * We copy only the Excel host bundles + en-us strings, not the other Office hosts,
 * which takes the payload from 85 MB down to roughly 5 MB.
 */
import { cp, mkdir, readdir, rm, writeFile, stat } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SRC = path.join(ROOT, "node_modules", "@microsoft", "office-js", "dist");
const OUT = path.join(ROOT, "dist", "vendor", "office");

/** Root-level scripts office.js may pull in for Excel on any platform. */
const KEEP_PREFIXES = ["excel-", "excelwebapp-", "excelios-", "office.js", "es6-promise.js", "custom-functions-runtime.js"];
/** Locales to ship. Office falls back to en-us when a locale is missing. */
const KEEP_LOCALES = ["en-us"];

const isWanted = (name) =>
  !name.includes(".debug.") &&
  name.endsWith(".js") &&
  KEEP_PREFIXES.some((p) => (p.endsWith(".js") ? name === p : name.startsWith(p)));

async function main() {
  if (!existsSync(SRC)) {
    throw new Error(`@microsoft/office-js not installed — expected ${SRC}. Run: npm install`);
  }

  await rm(OUT, { recursive: true, force: true });
  await mkdir(OUT, { recursive: true });

  let copied = 0;
  let bytes = 0;

  for (const entry of await readdir(SRC, { withFileTypes: true })) {
    if (entry.isFile() && isWanted(entry.name)) {
      const dest = path.join(OUT, entry.name);
      await cp(path.join(SRC, entry.name), dest);
      bytes += (await stat(dest)).size;
      copied++;
    }
  }

  for (const locale of KEEP_LOCALES) {
    const from = path.join(SRC, locale);
    if (!existsSync(from)) continue;
    const to = path.join(OUT, locale);
    await mkdir(to, { recursive: true });
    for (const f of await readdir(from)) {
      if (f.includes(".debug.") || !f.startsWith("office_strings")) continue;
      await cp(path.join(from, f), path.join(to, f));
      bytes += (await stat(path.join(to, f))).size;
      copied++;
    }
  }

  // office.js requests telemetry/oteljs_agave.js relative to its base path. We serve a
  // no-op module instead of Microsoft's real sink: it satisfies the loader while making
  // it structurally impossible for the runtime to emit telemetry, online or off.
  await mkdir(path.join(OUT, "telemetry"), { recursive: true });
  const stub = `/* Exovelletron: telemetry sink intentionally stubbed out. No data is collected. */
(function () {
  var noop = function () {};
  var sink = { sendTelemetryEvent: noop, writeEvent: noop, flush: noop, addSink: noop, setSink: noop };
  var g = typeof globalThis !== "undefined" ? globalThis : window;
  g.oteljs = g.oteljs || { Sinks: {}, SimpleEventBuilder: function () { return sink; } };
  g.oteljs_agave = g.oteljs_agave || sink;
  if (g.OTel && g.OTel.OTelLogger) { g.OTel.OTelLogger.sendTelemetryEvent = noop; }
  if (g.Microsoft && g.Microsoft.Office && g.Microsoft.Office.WebExtension) {
    g.Microsoft.Office.WebExtension.telemetrySink = sink;
  }
})();
`;
  await writeFile(path.join(OUT, "telemetry", "oteljs_agave.js"), stub, "utf8");
  await writeFile(path.join(OUT, "telemetry", "oteljs.js"), stub, "utf8");
  copied += 2;

  console.log(`office.js vendored -> dist/vendor/office (${copied} files, ${(bytes / 1e6).toFixed(1)} MB)`);
}

main().catch((err) => {
  console.error(err.message);
  process.exit(1);
});
