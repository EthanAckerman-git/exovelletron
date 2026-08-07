/**
 * Update checking against GitHub releases.
 *
 * The app is distributed as a DMG on the repo's releases page, so "auto-update" here
 * means: notice when a newer release exists and take the user to it. A single GET to
 * the public releases API — carrying nothing but the request itself — is the only
 * traffic, and it happens once per launch plus whenever the user asks.
 *
 * The version comparison is pure and unit-tested.
 */

export const REPO_URL = "https://github.com/EthanAckerman-git/exovelletron";
const LATEST_API = "https://api.github.com/repos/EthanAckerman-git/exovelletron/releases/latest";
const TIMEOUT_MS = 10_000;

/** "v1.2.0" → [1, 2, 0]; tolerant of missing parts and stray text. */
export function parseVersion(text) {
  const m = /(\d+)(?:\.(\d+))?(?:\.(\d+))?/.exec(String(text ?? ""));
  if (!m) return null;
  return [Number(m[1]), Number(m[2] ?? 0), Number(m[3] ?? 0)];
}

/** Positive when `a` is newer than `b`, negative when older, 0 when equal/unknown. */
export function compareVersions(a, b) {
  const va = parseVersion(a);
  const vb = parseVersion(b);
  if (!va || !vb) return 0;
  for (let i = 0; i < 3; i++) {
    if (va[i] !== vb[i]) return va[i] - vb[i];
  }
  return 0;
}

/**
 * Ask GitHub for the latest release and compare against the running version.
 * Never throws: an offline machine or a private repo simply reports "unknown".
 *
 * @param {string} currentVersion
 * @param {{fetchImpl?: typeof fetch}} [opts]
 * @returns {Promise<{status:"current"|"update"|"unknown", latest?:string, url?:string}>}
 */
export async function checkForUpdate(currentVersion, { fetchImpl = fetch } = {}) {
  try {
    const res = await fetchImpl(LATEST_API, {
      headers: { Accept: "application/vnd.github+json", "User-Agent": "Exovelletron" },
      signal: AbortSignal.timeout(TIMEOUT_MS),
    });
    if (!res.ok) return { status: "unknown" };
    const release = await res.json();
    const latest = String(release.tag_name ?? release.name ?? "").trim();
    if (!parseVersion(latest)) return { status: "unknown" };
    return compareVersions(latest, currentVersion) > 0
      ? { status: "update", latest, url: release.html_url ?? `${REPO_URL}/releases/latest` }
      : { status: "current", latest };
  } catch {
    return { status: "unknown" };
  }
}
