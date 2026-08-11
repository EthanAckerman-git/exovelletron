/**
 * What kind of Mac this is, in words a person recognises.
 *
 * The picker's promise is "recommended for *your* Mac", and "arm64 · 24 GB" does not
 * say that — "Apple M2 Pro · 24 GB" does. macOS exposes the marketing name through
 * sysctl; everything here degrades gracefully when it isn't available (Intel Macs,
 * odd virtualised environments, tests).
 */
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

/**
 * Parse a `machdep.cpu.brand_string` value into something displayable.
 *
 * "Apple M2 Pro" → { chip: "Apple M2 Pro", apple: true }
 * "Intel(R) Core(TM) i7-9750H CPU @ 2.60GHz" → { chip: "Intel", apple: false }
 * garbage/empty → { chip: null, apple: false }
 */
export function parseChipString(raw) {
  const text = String(raw ?? "").trim();
  if (!text) return { chip: null, apple: false };
  if (/^apple\s/i.test(text)) return { chip: text.replace(/\s+/g, " "), apple: true };
  if (/intel/i.test(text)) return { chip: "Intel", apple: false };
  return { chip: null, apple: false };
}

/**
 * Ask macOS what chip this is. Never throws — a Mac we can't identify still works,
 * it just gets the generic fit labels instead of a name.
 */
export async function detectChip({ execImpl = execFileAsync } = {}) {
  try {
    const { stdout } = await execImpl("sysctl", ["-n", "machdep.cpu.brand_string"]);
    return parseChipString(stdout);
  } catch {
    return { chip: null, apple: process.arch === "arm64" };
  }
}
