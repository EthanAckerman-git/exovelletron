/**
 * Local HTTPS certificate: generate once, trust once.
 *
 * Excel renders the task pane as an HTTPS page, so any call to plain http://localhost
 * is blocked as mixed content. Serving the UI *and* the API from https://localhost
 * makes them same-origin, which sidesteps mixed content and CORS entirely.
 *
 * We mint a small private CA and a leaf for localhost, then trust only the CA in the
 * user's login keychain. WKWebView (which backs Excel's task pane on macOS) honours
 * login-keychain trust roots, which is the same mechanism Microsoft's own
 * office-addin-dev-certs uses.
 */
import { execFile } from "node:child_process";
import { readFile, writeFile, mkdir, rm } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { promisify } from "node:util";
import { paths as defaultPaths } from "../config.js";

const run = promisify(execFile);

export const CA_COMMON_NAME = "Exovelletron Development CA";

/**
 * The root is trusted once and lives long, because re-trusting means another macOS
 * password prompt. Root certificates are not subject to the server-certificate lifetime
 * cap.
 */
const CA_DAYS = 3650;

/**
 * Apple caps TLS *server* certificate lifetimes at 398 days. WebKit — which renders
 * Excel's task pane — rejects anything longer outright with "the content is blocked
 * because it isn't signed by a valid security certificate", regardless of whether the
 * issuing root is trusted. A 10-year leaf silently broke the add-in for exactly this
 * reason.
 */
export const MAX_LEAF_DAYS = 398;
const LEAF_DAYS = 397;

/** Renew the leaf well before it lapses so a long-running install never breaks. */
const RENEW_WHEN_DAYS_LEFT_BELOW = 30;

export function certFiles(p = defaultPaths) {
  return {
    caKey: path.join(p.certsDir, "ca-key.pem"),
    caCert: path.join(p.certsDir, "ca-cert.pem"),
    key: path.join(p.certsDir, "server-key.pem"),
    cert: path.join(p.certsDir, "server-cert.pem"),
  };
}

/** Days until `certPath` expires, or null when it is missing/unreadable. */
export async function daysUntilExpiry(certPath) {
  if (!existsSync(certPath)) return null;
  try {
    const { stdout } = await run("openssl", ["x509", "-in", certPath, "-noout", "-enddate"]);
    const raw = stdout.trim().replace(/^notAfter=/, "");
    const expires = new Date(raw);
    if (Number.isNaN(expires.getTime())) return null;
    return Math.floor((expires.getTime() - Date.now()) / 86_400_000);
  } catch {
    return null;
  }
}

export async function certsExist(p = defaultPaths) {
  const f = certFiles(p);
  if (!Object.values(f).every((x) => existsSync(x))) return false;
  const left = await daysUntilExpiry(f.cert);
  return left !== null && left > RENEW_WHEN_DAYS_LEFT_BELOW;
}

/** True when a usable CA is already on disk, so trust does not have to be re-granted. */
export async function caExists(p = defaultPaths) {
  const f = certFiles(p);
  if (!existsSync(f.caKey) || !existsSync(f.caCert)) return false;
  const left = await daysUntilExpiry(f.caCert);
  return left !== null && left > RENEW_WHEN_DAYS_LEFT_BELOW;
}

/** Create the private CA. Kept separate from the leaf so renewals never re-prompt. */
export async function generateCa(p = defaultPaths) {
  const f = certFiles(p);
  await mkdir(p.certsDir, { recursive: true });
  await run("openssl", [
    "req", "-x509", "-newkey", "rsa:2048", "-nodes",
    "-keyout", f.caKey, "-out", f.caCert,
    "-days", String(CA_DAYS),
    "-subj", `/CN=${CA_COMMON_NAME}/O=Exovelletron`,
    "-addext", "basicConstraints=critical,CA:TRUE,pathlen:0",
    "-addext", "keyUsage=critical,keyCertSign,cRLSign",
  ]);
  return f;
}

/**
 * Issue a fresh localhost leaf from the existing CA.
 *
 * Renewing only the leaf matters: regenerating the CA would invalidate the trust the
 * user already granted and cost them another password prompt every time the
 * certificate rolled over.
 */
export async function generateLeaf(p = defaultPaths) {
  const f = certFiles(p);
  await mkdir(p.certsDir, { recursive: true });
  if (!existsSync(f.caCert) || !existsSync(f.caKey)) {
    throw new Error("Certificate authority is missing; generate it before issuing a leaf.");
  }

  const csr = path.join(p.certsDir, "server.csr");
  const extFile = path.join(p.certsDir, "server-ext.cnf");
  await writeFile(
    extFile,
    [
      "basicConstraints=CA:FALSE",
      "keyUsage=critical,digitalSignature,keyEncipherment",
      "extendedKeyUsage=serverAuth",
      "subjectAltName=DNS:localhost,DNS:127.0.0.1,IP:127.0.0.1,IP:::1",
    ].join("\n") + "\n",
    "utf8",
  );

  try {
    await run("openssl", [
      "req", "-newkey", "rsa:2048", "-nodes",
      "-keyout", f.key, "-out", csr,
      "-subj", "/CN=localhost/O=Exovelletron",
    ]);
    await run("openssl", [
      "x509", "-req", "-in", csr,
      "-CA", f.caCert, "-CAkey", f.caKey, "-CAcreateserial",
      "-out", f.cert, "-days", String(LEAF_DAYS),
      "-extfile", extFile,
    ]);
  } finally {
    await rm(csr, { force: true });
    await rm(extFile, { force: true });
  }

  return f;
}

/**
 * True when our CA carries an SSL trust setting in the user's keychain.
 *
 * `security verify-cert` is the obvious call and the wrong one: it consults the system
 * trust store and applies the full SSL policy including Certificate Transparency, which
 * a private CA can never satisfy. It reports "not trusted" for a CA that is, in fact,
 * trusted — which would make setup re-prompt for a password on every launch.
 *
 * The authoritative source is the user's own trust settings, which is what WKWebView
 * (and therefore Excel's task pane) actually consults.
 */
export async function isCaTrusted(p = defaultPaths) {
  const { caCert } = certFiles(p);
  if (!existsSync(caCert)) return false;
  try {
    const { stdout } = await run("security", ["dump-trust-settings"]);
    return stdout.includes(CA_COMMON_NAME);
  } catch {
    // Exits non-zero when the user has no trust settings at all.
    return false;
  }
}

/**
 * Add the CA to the login keychain as a trusted SSL root.
 * Triggers exactly one macOS password prompt — the only one in the whole setup.
 */
export async function trustCa(p = defaultPaths) {
  const { caCert } = certFiles(p);
  if (!existsSync(caCert)) throw new Error("Certificate authority not generated yet.");
  const loginKeychain = path.join(p.home, "Library", "Keychains", "login.keychain-db");
  try {
    await run("security", [
      "add-trusted-cert", "-r", "trustRoot", "-p", "ssl",
      "-k", loginKeychain, caCert,
    ]);
  } catch (err) {
    const msg = String(err.stderr || err.message || "");
    if (/User canceled|The authorization was cancelled|-128/i.test(msg)) {
      throw new Error("Certificate trust was cancelled. Excel cannot load the add-in without it.");
    }
    throw new Error(`Could not trust the local certificate: ${msg.trim() || "unknown error"}`);
  }
}

/** Remove our CA from the keychain — used by the uninstaller. */
export async function untrustCa(p = defaultPaths) {
  const { caCert } = certFiles(p);
  if (!existsSync(caCert)) return;
  try {
    await run("security", ["remove-trusted-cert", caCert]);
  } catch {
    /* already gone */
  }
}

/**
 * Generate whatever is missing or expiring, and return the PEMs.
 *
 * Deliberately does NOT trust: the server can listen on a self-signed certificate
 * perfectly well, and trusting triggers a macOS password dialog. Blocking app startup
 * on a modal the user has not asked for is bad behaviour, and if they dismissed it the
 * app would fail to start at all. Trust is an explicit step in the setup wizard.
 */
export async function ensureCertFiles(p = defaultPaths) {
  if (!(await caExists(p))) await generateCa(p);
  if (!(await certsExist(p))) await generateLeaf(p);
  const f = certFiles(p);
  return { key: await readFile(f.key), cert: await readFile(f.cert) };
}

/** Create both halves from scratch. Used by tests and first-run setup. */
export async function generateCerts(p = defaultPaths) {
  await generateCa(p);
  return generateLeaf(p);
}

/** Generate if needed, then trust. Invoked by the wizard's certificate step. */
export async function ensureTrustedCerts(p = defaultPaths) {
  const credentials = await ensureCertFiles(p);
  if (!(await isCaTrusted(p))) await trustCa(p);
  return credentials;
}
