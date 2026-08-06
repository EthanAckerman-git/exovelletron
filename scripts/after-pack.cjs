/**
 * electron-builder afterPack hook: strip resource forks before signing.
 *
 * codesign refuses any file carrying "resource fork, Finder information, or similar
 * detritus". This project builds from ~/Desktop, which macOS syncs through iCloud and
 * stamps with com.apple.macl and resource-fork metadata.
 *
 * `xattr -cr` is not sufficient — it clears extended attributes but leaves the resource
 * fork itself. `ditto --norsrc --noextattr --noqtn` rebuilds the tree without either,
 * which is what actually lets codesign through. (com.apple.provenance survives the copy
 * because the OS re-applies it on write; codesign tolerates it on its own.)
 *
 * Runs after packaging and before signing — the only window where this sticks.
 */
const { execFileSync } = require("node:child_process");
const { rmSync, renameSync, existsSync } = require("node:fs");

exports.default = async function afterPack(context) {
  if (context.electronPlatformName !== "darwin") return;

  const appPath = `${context.appOutDir}/${context.packager.appInfo.productFilename}.app`;
  if (!existsSync(appPath)) return;

  const scratch = `${appPath}.clean`;
  try {
    rmSync(scratch, { recursive: true, force: true });
    execFileSync("ditto", ["--norsrc", "--noextattr", "--noqtn", appPath, scratch], { stdio: "pipe" });
    rmSync(appPath, { recursive: true, force: true });
    renameSync(scratch, appPath);
    console.log("  • stripped resource forks for codesign");
  } catch (err) {
    rmSync(scratch, { recursive: true, force: true });
    console.warn(`  • could not strip resource forks: ${err.message}`);
  }
};
