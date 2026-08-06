/**
 * Generates every icon the app and the Office manifest need, from code.
 *
 * No image dependencies: a tiny PNG encoder (zlib is built into Node) plus 4x
 * supersampled software rendering. Drawing the mark procedurally means it stays crisp
 * at 16px and 1024px alike, and there are no binary blobs in the repo.
 *
 * The mark: a violet-to-indigo squircle holding a 3x3 grid of cells, with the
 * top-right cell lit up — a spreadsheet where one cell filled itself in.
 */
import { deflateSync } from "node:zlib";
import { writeFile, mkdir, rm } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const run = promisify(execFile);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/* ---------------------------------------------------------------- PNG encoder */

const CRC_TABLE = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

/** @param {Uint8Array} rgba  width*height*4 */
function encodePng(rgba, width, height) {
  const stride = width * 4;
  const raw = Buffer.alloc((stride + 1) * height);
  for (let y = 0; y < height; y++) {
    raw[y * (stride + 1)] = 0; // filter: none
    Buffer.from(rgba.buffer, rgba.byteOffset + y * stride, stride).copy(raw, y * (stride + 1) + 1);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // colour type: RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

/* ------------------------------------------------------------------- drawing */

/** Signed distance to a rounded rectangle, negative inside. */
function sdRoundRect(px, py, cx, cy, halfW, halfH, r) {
  const qx = Math.abs(px - cx) - (halfW - r);
  const qy = Math.abs(py - cy) - (halfH - r);
  const ox = Math.max(qx, 0);
  const oy = Math.max(qy, 0);
  return Math.sqrt(ox * ox + oy * oy) + Math.min(Math.max(qx, qy), 0) - r;
}

const mix = (a, b, t) => a + (b - a) * t;

/**
 * Renders the mark at `size` px with 4x supersampling.
 * @returns {Uint8Array} RGBA buffer
 */
function renderIcon(size) {
  const SS = 4;
  const S = size * SS;
  const acc = new Float32Array(size * size * 4);

  // Squircle geometry, inset slightly so the shape has breathing room in the tile.
  const inset = S * 0.045;
  const halfW = (S - inset * 2) / 2;
  const cx = S / 2;
  const r = (S - inset * 2) * 0.235;

  // 3x3 cell grid inside the squircle.
  const gridSpan = (S - inset * 2) * 0.56;
  const gridLeft = cx - gridSpan / 2;
  const gap = gridSpan * 0.14;
  const cell = (gridSpan - gap * 2) / 3;
  const cellR = cell * 0.26;

  for (let sy = 0; sy < S; sy++) {
    for (let sx = 0; sx < S; sx++) {
      const px = sx + 0.5;
      const py = sy + 0.5;

      const d = sdRoundRect(px, py, cx, cx, halfW, halfW, r);
      // Antialias across roughly one supersampled pixel of the edge.
      const shapeA = Math.min(Math.max(0.5 - d, 0), 1);
      if (shapeA <= 0) continue;

      // Diagonal gradient: violet (top-left) -> indigo (bottom-right).
      const t = Math.min(Math.max((px / S) * 0.5 + (py / S) * 0.5, 0), 1);
      let cr = mix(139, 67, t);
      let cg = mix(92, 56, t);
      let cb = mix(246, 202, t);

      // Soft highlight in the upper-left for a little depth.
      const hx = (px - S * 0.3) / (S * 0.55);
      const hy = (py - S * 0.24) / (S * 0.55);
      const hl = Math.max(0, 1 - (hx * hx + hy * hy)) * 0.18;
      cr = mix(cr, 255, hl);
      cg = mix(cg, 255, hl);
      cb = mix(cb, 255, hl);

      // Cells.
      let ink = 0;
      let lit = 0;
      for (let row = 0; row < 3; row++) {
        for (let col = 0; col < 3; col++) {
          const bx = gridLeft + col * (cell + gap) + cell / 2;
          const by = cx - gridSpan / 2 + row * (cell + gap) + cell / 2;
          const cd = sdRoundRect(px, py, bx, by, cell / 2, cell / 2, cellR);
          const a = Math.min(Math.max(0.5 - cd, 0), 1);
          if (a <= 0) continue;
          const isLit = row === 0 && col === 2;
          ink = Math.max(ink, a);
          if (isLit) lit = Math.max(lit, a);
        }
      }

      if (ink > 0) {
        const alpha = lit > 0 ? lit : ink * 0.62;
        const target = 255;
        cr = mix(cr, target, alpha);
        cg = mix(cg, target, alpha);
        cb = mix(cb, target, alpha);
      }

      const ox = Math.floor(sx / SS);
      const oy = Math.floor(sy / SS);
      const oi = (oy * size + ox) * 4;
      const w = shapeA / (SS * SS);
      acc[oi] += cr * w;
      acc[oi + 1] += cg * w;
      acc[oi + 2] += cb * w;
      acc[oi + 3] += 255 * w;
    }
  }

  const out = new Uint8Array(size * size * 4);
  for (let i = 0; i < out.length; i += 4) {
    const a = acc[i + 3];
    if (a <= 0.5) continue;
    // Un-premultiply so edges keep their colour instead of darkening toward black.
    const k = 255 / a;
    out[i] = Math.min(255, Math.round(acc[i] * k));
    out[i + 1] = Math.min(255, Math.round(acc[i + 1] * k));
    out[i + 2] = Math.min(255, Math.round(acc[i + 2] * k));
    out[i + 3] = Math.min(255, Math.round(a));
  }
  return out;
}

/* ---------------------------------------------------------------------- main */

const MANIFEST_SIZES = [16, 32, 64, 80, 128, 256, 512];
const ICNS_SIZES = [
  ["icon_16x16.png", 16], ["icon_16x16@2x.png", 32],
  ["icon_32x32.png", 32], ["icon_32x32@2x.png", 64],
  ["icon_128x128.png", 128], ["icon_128x128@2x.png", 256],
  ["icon_256x256.png", 256], ["icon_256x256@2x.png", 512],
  ["icon_512x512.png", 512], ["icon_512x512@2x.png", 1024],
];

async function main() {
  const assetsDir = path.join(ROOT, "assets");
  const webIconsDir = path.join(ROOT, "dist", "assets");
  await mkdir(assetsDir, { recursive: true });
  await mkdir(webIconsDir, { recursive: true });

  const cache = new Map();
  const png = (size) => {
    if (!cache.has(size)) cache.set(size, encodePng(renderIcon(size), size, size));
    return cache.get(size);
  };

  for (const size of MANIFEST_SIZES) {
    const buf = png(size);
    await writeFile(path.join(webIconsDir, `icon-${size}.png`), buf);
  }
  await writeFile(path.join(assetsDir, "icon.png"), png(512));

  // .icns for the macOS app bundle.
  const iconset = path.join(assetsDir, "icon.iconset");
  await rm(iconset, { recursive: true, force: true });
  await mkdir(iconset, { recursive: true });
  for (const [name, size] of ICNS_SIZES) {
    await writeFile(path.join(iconset, name), png(size));
  }
  try {
    await run("iconutil", ["-c", "icns", iconset, "-o", path.join(assetsDir, "icon.icns")]);
    await rm(iconset, { recursive: true, force: true });
    console.log("icons -> dist/assets/*.png, assets/icon.png, assets/icon.icns");
  } catch (err) {
    console.warn(`icons -> PNGs written; iconutil unavailable (${err.message})`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
