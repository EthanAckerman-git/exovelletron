# Excel AI Local

A fully offline AI assistant for Microsoft Excel on macOS. A chat pane inside Excel that
can read your selection, answer questions about the sheet, and propose changes you
approve before anything is written.

Nothing leaves your Mac. There is no account, no API key, and no network call after
setup.

---

## What it does

- **Chat about the open workbook.** The pane sees your selection, the column headers, and
  a sample of rows.
- **Proposes changes rather than making them.** When the model wants to edit the sheet it
  emits a structured action, which the pane renders as a preview showing the exact cells
  that will change. Nothing is written until you click **Apply**, and every applied change
  has an **Undo**.
- **Prefers formulas over pasted values,** because the model only ever sees a sample of
  your rows and a formula stays correct as data changes.

Supported actions: write values, fill a formula, format cells, insert a column, sort a
range.

## Requirements

- macOS on Apple Silicon
- Microsoft Excel for Mac (tested against 16.111)
- ~3 GB of disk for the default model, 8 GB RAM minimum

## Install

1. Open `Excel AI Local-1.0.0-arm64.dmg` and drag the app to Applications.
2. Launch it. The setup panel walks through three steps:
   - **Local certificate** — one macOS password prompt.
   - **Excel add-in** — see the note below.
   - **Model** — a one-time 2.9 GB download, verified against a SHA-256 digest.
3. Restart Excel. The add-in appears on the **Home** tab as **Excel AI Local**.

### The add-in step needs one drag from you

macOS seals Excel's add-in folder (`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`)
off from every other process. This is not something the app can work around, and it is not
fixed by granting Full Disk Access — that was tested and is still refused. A drag in Finder
works because the drop carries user intent, which macOS honours.

So the app opens two Finder windows and you drag one file across. Ten seconds, once.

## Models

| Model | Download | RAM | Notes |
|---|---|---|---|
| Qwen3.5 2B | 1.34 GB | 4 GB+ | Fastest, weakest at multi-step reasoning |
| **Qwen3.5 4B** | **2.91 GB** | **8 GB+** | **Default.** Best balance |
| Qwen3.5 9B | 5.97 GB | 16 GB+ | Best at complex analysis |

All are Unsloth `UD-Q4_K_XL` dynamic quantisations, which hold up better than a flat
`Q4_K_M` at effectively the same size. Downloads resume if interrupted and are rejected if
the digest does not match.

On an M5 Pro the 4B model runs at roughly **60 tokens/second** with all layers on the GPU.

## How it is put together

```
Excel task pane  ──HTTPS + SSE──▶  local server (127.0.0.1)  ──▶  node-llama-cpp / Metal
   Office.js                        serves pane + API              model in ~/Library
```

The pane and the API are served from the same origin (`https://localhost:39217`). Excel
renders the task pane as an HTTPS page, so a plain `http://localhost` call would be blocked
as mixed content; same-origin HTTPS sidesteps both that and CORS. The app mints a private
CA, trusts it in the login keychain, and issues a leaf for `localhost`.

**Offline:** the Office.js runtime is vendored and served locally rather than loaded from
Microsoft's CDN, and its telemetry sink is replaced with a no-op stub. The pane's
Content-Security-Policy pins `connect-src` to our own origin, so it cannot reach the
internet even if something tried.

**Security:** the server binds `127.0.0.1` only. Every API call requires a random
per-run session token injected into the pane at serve time, plus an origin check, so other
local software and stray browser tabs cannot drive it.

### Layout

```
core/       engine, model catalog/downloader, HTTPS server, setup steps  (unit tested)
taskpane/   the Excel-side UI: sheet reading, action preview, chat
desktop/    Electron shell: control panel, tray, setup wizard
shared/     design tokens used by both UIs
scripts/    build, icon generation, packaging, smoke test
```

## Development

```bash
npm install
npm run prepare:assets   # vendor Office.js, generate icons, build the pane
npm run dev              # run the app from source
npm test                 # 90 unit tests
```

Other commands:

```bash
node scripts/fetch-model.mjs qwen3.5-4b   # download a model from the CLI
node scripts/smoke-engine.mjs             # end-to-end check against the real model
npm run dist                              # build the signed .app and .dmg
```

`npm run build:web -- --dev` also emits `/app/preview.html`, a harness that renders the
transcript components against sample data so the pane can be iterated on in a browser
instead of inside Excel.

### Two build details worth knowing

- **Packaging happens in `~/Library/Caches`, not in the project.** This tree lives under
  `~/Desktop`, which iCloud syncs and stamps with `com.apple.macl` and resource-fork
  metadata. `codesign` rejects those as "detritus", and since signing itself writes to the
  files they get re-stamped faster than any cleanup can strip them. Building outside the
  synced tree avoids it. The finished DMG is copied back to `release/`.
- **The DMG is signed but not notarized.** Notarization needs a paid Apple Developer
  account. On your own machine this is irrelevant; sending the DMG to someone else will hit
  Gatekeeper until it is notarized.

## Known limits

- macOS Apple Silicon only. The core is platform-agnostic, but the installer, certificate
  trust, and add-in registration are macOS-specific.
- One model resident at a time; switching unloads and reloads.
- The model sees a sample of rows, never the whole sheet. This is deliberate — it is what
  keeps a 200,000-row workbook from being serialised across the Office bridge — and it is
  why the prompt pushes hard toward formulas.
