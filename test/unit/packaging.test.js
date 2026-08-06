import { describe, it, expect } from "vitest";
import { readFileSync, existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
const pkg = JSON.parse(readFileSync(path.join(ROOT, "package.json"), "utf8"));

/** Top-level directory each electron-builder `files` glob admits. */
const packagedRoots = new Set(
  pkg.build.files
    .filter((g) => !g.startsWith("!"))
    .map((g) => g.split("/")[0]),
);

/** Local hrefs/srcs referenced by an HTML file, ignoring absolute and remote URLs. */
function localRefs(htmlPath) {
  const html = readFileSync(htmlPath, "utf8");
  return [...html.matchAll(/(?:href|src)="([^"]+)"/g)]
    .map((m) => m[1])
    .filter((ref) => !/^(https?:|data:|#|\/)/.test(ref));
}

describe("electron packaging", () => {
  // Regression: shared/tokens.css and assets/trayTemplate.png were referenced at runtime
  // but excluded from build.files. The app launched and looked broken only once packaged
  // — serif fallback text and unstyled buttons, because every CSS variable was undefined.
  const rendererHtml = path.join(ROOT, "desktop", "renderer", "index.html");

  it("ships every file the control panel references", () => {
    for (const ref of localRefs(rendererHtml)) {
      const resolved = path.resolve(path.dirname(rendererHtml), ref);
      expect(existsSync(resolved), `${ref} does not exist on disk`).toBe(true);

      const relative = path.relative(ROOT, resolved);
      expect(relative.startsWith(".."), `${ref} resolves outside the project`).toBe(false);
      expect(
        packagedRoots.has(relative.split(path.sep)[0]),
        `${ref} lives outside build.files (${[...packagedRoots].join(", ")})`,
      ).toBe(true);
    }
  });

  it("ships the runtime files the main process loads by path", () => {
    for (const relative of ["assets/trayTemplate.png", "desktop/preload.cjs"]) {
      expect(existsSync(path.join(ROOT, relative)), `${relative} missing`).toBe(true);
      expect(packagedRoots.has(relative.split("/")[0]), `${relative} not packaged`).toBe(true);
    }
  });

  it("unpacks the native inference addon from the asar", () => {
    // node-llama-cpp loads .dylib/.node files by absolute path, which cannot be read
    // from inside an asar archive.
    const unpack = pkg.build.asarUnpack.join(" ");
    expect(unpack).toContain("node-llama-cpp");
  });

  it("keeps the app entry point and icon wired up", () => {
    expect(existsSync(path.join(ROOT, pkg.main))).toBe(true);
    expect(pkg.build.mac.icon).toBe("assets/icon.icns");
  });
});
