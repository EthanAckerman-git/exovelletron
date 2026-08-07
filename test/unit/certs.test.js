import { describe, it, expect, beforeAll, afterAll } from "vitest";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdtemp, rm } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { createPaths } from "../../core/config.js";
import {
  generateCa, generateLeaf, generateCerts, certFiles, caExists, certsExist,
  daysUntilExpiry, MAX_LEAF_DAYS, CA_COMMON_NAME,
} from "../../core/setup/certs.js";

const run = promisify(execFile);
let home;
let paths;

beforeAll(async () => {
  home = await mkdtemp(path.join(tmpdir(), "eal-certs-"));
  paths = createPaths(home);
  await generateCerts(paths);
}, 60_000);

afterAll(async () => {
  await rm(home, { recursive: true, force: true });
});

const text = async (file) => (await run("openssl", ["x509", "-in", file, "-noout", "-text"])).stdout;

describe("local certificates", () => {
  it("writes all four files", async () => {
    for (const file of Object.values(certFiles(paths))) {
      expect(existsSync(file), file).toBe(true);
    }
  });

  // Regression: a 3650-day leaf made WebKit refuse the task pane with "the content is
  // blocked because it isn't signed by a valid security certificate", even though the
  // issuing root was trusted. Apple caps TLS server certificates at 398 days.
  it("issues a leaf inside Apple's 398-day server certificate limit", async () => {
    const left = await daysUntilExpiry(certFiles(paths).cert);
    expect(left).toBeGreaterThan(0);
    expect(left).toBeLessThanOrEqual(MAX_LEAF_DAYS);
  });

  it("lets the root outlive the leaf so trust is granted only once", async () => {
    const caLeft = await daysUntilExpiry(certFiles(paths).caCert);
    const leafLeft = await daysUntilExpiry(certFiles(paths).cert);
    expect(caLeft).toBeGreaterThan(leafLeft);
  });

  it("covers localhost and loopback addresses", async () => {
    const details = await text(certFiles(paths).cert);
    expect(details).toMatch(/DNS:localhost/);
    expect(details).toMatch(/IP Address:127\.0\.0\.1/);
  });

  it("marks the leaf as a server certificate, not a CA", async () => {
    const details = await text(certFiles(paths).cert);
    expect(details).toMatch(/TLS Web Server Authentication/);
    expect(details).toMatch(/CA:FALSE/);
  });

  it("marks the root as a CA that can sign", async () => {
    const details = await text(certFiles(paths).caCert);
    expect(details).toMatch(/CA:TRUE/);
    expect(details).toMatch(/Certificate Sign/);
    expect(details).toContain(CA_COMMON_NAME);
  });

  it("issues the leaf from our own root", async () => {
    const details = await text(certFiles(paths).cert);
    expect(details).toContain(CA_COMMON_NAME);
  });

  it("reports both halves as present and current", async () => {
    expect(await caExists(paths)).toBe(true);
    expect(await certsExist(paths)).toBe(true);
  });

  it("renews the leaf without touching the root", async () => {
    const before = await run("openssl", ["x509", "-in", certFiles(paths).caCert, "-noout", "-fingerprint"]);
    const leafBefore = await run("openssl", ["x509", "-in", certFiles(paths).cert, "-noout", "-serial"]);

    await generateLeaf(paths);

    const after = await run("openssl", ["x509", "-in", certFiles(paths).caCert, "-noout", "-fingerprint"]);
    const leafAfter = await run("openssl", ["x509", "-in", certFiles(paths).cert, "-noout", "-serial"]);

    // Same root — the user's trust survives; new leaf — the rotation actually happened.
    expect(after.stdout).toBe(before.stdout);
    expect(leafAfter.stdout).not.toBe(leafBefore.stdout);
  }, 30_000);

  it("refuses to issue a leaf with no root present", async () => {
    const orphan = createPaths(await mkdtemp(path.join(tmpdir(), "eal-orphan-")));
    await expect(generateLeaf(orphan)).rejects.toThrow(/authority is missing/i);
    await rm(orphan.home, { recursive: true, force: true });
  });

  it("reports no certificates for an empty home", async () => {
    const empty = createPaths(await mkdtemp(path.join(tmpdir(), "eal-empty-")));
    expect(await caExists(empty)).toBe(false);
    expect(await certsExist(empty)).toBe(false);
    expect(await daysUntilExpiry(certFiles(empty).cert)).toBeNull();
    await rm(empty.home, { recursive: true, force: true });
  });
});
