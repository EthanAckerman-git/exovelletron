/**
 * Packages the macOS app and DMG.
 *
 * The build output deliberately lands in ~/Library/Caches rather than inside the
 * project. This project tree lives under ~/Desktop, which macOS syncs through iCloud
 * and stamps every written file with com.apple.macl and resource-fork metadata —
 * codesign rejects those as "detritus", and because signing itself writes to the files,
 * they get re-stamped faster than any cleanup can strip them. Building outside the
 * synced tree avoids the problem at the source. The finished DMG is copied back.
 */
import { execFileSync } from "node:child_process";
import { mkdirSync, rmSync, readdirSync, copyFileSync, existsSync } from "node:fs";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const BUILD_DIR = path.join(os.homedir(), "Library", "Caches", "ExovelletronBuild", "release");
const LOCAL_RELEASE = path.join(ROOT, "release");

rmSync(BUILD_DIR, { recursive: true, force: true });
mkdirSync(BUILD_DIR, { recursive: true });

execFileSync(
  "npx",
  ["electron-builder", "--mac", "--arm64", `-c.directories.output=${BUILD_DIR}`],
  { stdio: "inherit", cwd: ROOT },
);

mkdirSync(LOCAL_RELEASE, { recursive: true });
let copied = 0;
for (const file of readdirSync(BUILD_DIR)) {
  if (!file.endsWith(".dmg")) continue;
  copyFileSync(path.join(BUILD_DIR, file), path.join(LOCAL_RELEASE, file));
  copied++;
  console.log(`\n  ${file} -> release/${file}`);
}

const appPath = path.join(BUILD_DIR, "mac-arm64", "Exovelletron.app");
if (existsSync(appPath)) console.log(`  app bundle: ${appPath}`);
if (!copied) {
  console.error("No DMG was produced.");
  process.exit(1);
}
