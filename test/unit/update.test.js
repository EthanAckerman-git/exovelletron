import { describe, it, expect, vi } from "vitest";
import { parseVersion, compareVersions, checkForUpdate } from "../../core/update.js";

describe("parseVersion", () => {
  it("reads plain and v-prefixed versions", () => {
    expect(parseVersion("1.2.0")).toEqual([1, 2, 0]);
    expect(parseVersion("v2.10.3")).toEqual([2, 10, 3]);
    expect(parseVersion("v3")).toEqual([3, 0, 0]);
  });

  it("returns null for versionless text", () => {
    expect(parseVersion("latest")).toBeNull();
    expect(parseVersion("")).toBeNull();
    expect(parseVersion(null)).toBeNull();
  });
});

describe("compareVersions", () => {
  it("orders numerically, not lexically", () => {
    expect(compareVersions("1.10.0", "1.9.0")).toBeGreaterThan(0);
    expect(compareVersions("v1.2.0", "1.2.0")).toBe(0);
    expect(compareVersions("1.2.0", "1.2.1")).toBeLessThan(0);
  });
});

describe("checkForUpdate", () => {
  const release = (tag) => ({
    ok: true,
    json: async () => ({ tag_name: tag, html_url: `https://github.com/EthanAckerman-git/exovelletron/releases/tag/${tag}` }),
  });

  it("reports an update when the release is newer", async () => {
    const result = await checkForUpdate("1.2.0", { fetchImpl: vi.fn(async () => release("v1.3.0")) });
    expect(result).toMatchObject({ status: "update", latest: "v1.3.0" });
    expect(result.url).toContain("/releases/tag/v1.3.0");
  });

  it("reports current when up to date or ahead", async () => {
    expect((await checkForUpdate("1.3.0", { fetchImpl: vi.fn(async () => release("v1.3.0")) })).status).toBe("current");
    expect((await checkForUpdate("2.0.0", { fetchImpl: vi.fn(async () => release("v1.3.0")) })).status).toBe("current");
  });

  // A private repo answers 404; an offline Mac throws. Neither may crash the app or nag.
  it("never throws: offline and non-200 both come back as unknown", async () => {
    expect((await checkForUpdate("1.2.0", { fetchImpl: vi.fn(async () => ({ ok: false, status: 404 })) })).status).toBe("unknown");
    expect((await checkForUpdate("1.2.0", { fetchImpl: vi.fn(async () => { throw new Error("ENOTFOUND"); }) })).status).toBe("unknown");
  });

  it("ignores a release with an unparseable tag", async () => {
    const result = await checkForUpdate("1.2.0", {
      fetchImpl: vi.fn(async () => ({ ok: true, json: async () => ({ tag_name: "latest" }) })),
    });
    expect(result.status).toBe("unknown");
  });
});
