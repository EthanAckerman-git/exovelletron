<div align="center">

# Exovelletron

**A private AI assistant that lives inside Microsoft Excel.**

Ask questions about your spreadsheet, and it answers. Ask it to change something, and it
shows you exactly what it will do before it does it.

Everything runs on your own Mac. No account, no API key, no internet after setup.

</div>

---

## What it can do

| | |
|---|---|
| **Explain your data** | "What am I looking at?" — it reads your selection and tells you. |
| **Write formulas** | Describe the calculation; it fills the column and adjusts references per row. |
| **Split messy columns** | One cell holding `"NAME 123 MAIN ST","CITY","ST","12345"`? It breaks that into proper named columns. |
| **Rebuild truly messy data** | Records with fields stacked down separate rows, or a dozen records crammed into one cell? It reads the whole range in context and rebuilds it as a clean table — one row per record. |
| **Clean up values** | Standardise addresses to USPS format, fix inconsistent dates, normalise names — across **every row**, not just the ones on screen. |
| **Format the sheet** | Bold headers, fills, currency and date formats. |
| **Find problems** | Blanks, duplicates, values that look wrong. |

**Nothing is written until you click Apply**, and every change has an **Undo**.

---

## Install

1. Open **`Exovelletron-1.0.0-arm64.dmg`** and drag the app to Applications.
2. Open Exovelletron. It walks you through three steps:

   | Step | What happens |
   |---|---|
   | **Local certificate** | One macOS password prompt. Excel only loads add-ins over HTTPS. |
   | **Excel add-in** | A folder picker opens. Click **Grant Access**. |
   | **Model** | A one-time 2.9 GB download, verified as it lands. |

3. Quit and reopen Excel.
4. **Home** tab → **Add-ins** → **Exovelletron**.

That's it. From then on it just works — open Excel, click the button.

### Why the folder picker?

macOS seals Excel's add-in folder off from every other program. Granting Full Disk Access
does **not** help — that was tested, and even Terminal *with* Full Disk Access is refused.
Picking the folder yourself counts as permission, so the app opens the picker already
pointing at the right place. One click, once.

---

## Requirements

- **macOS on Apple Silicon** (M1 or newer)
- **Microsoft Excel for Mac** — tested on 16.111
- **8 GB RAM** and about 3 GB of disk

---

## Models

Pick a different one any time from the app; it tells you how well each fits your Mac.

| Model | Download | Best for |
|---|---|---|
| Qwen3.5 2B | 1.3 GB | Older Macs, fastest replies |
| **Qwen3.5 4B** | **2.9 GB** | **Default — the right balance** |
| Qwen3.5 9B | 6.0 GB | Complex analysis, 16 GB+ Macs |

All are Unsloth `UD-Q4_K_XL` dynamic quantisations. Downloads resume if interrupted and
are rejected outright if the checksum does not match.

On an M5 Pro the 4B model runs at roughly **60 tokens/second**, entirely on the GPU.

---

## How it works

```
┌── Microsoft Excel ─────────────────┐
│  Task pane                          │
│    reads your selection             │
│    shows changes before applying    │
└──────────────┬──────────────────────┘
               │ HTTPS, localhost only
┌──────────────▼──────────────────────┐
│  Exovelletron.app                   │
│    local web server (127.0.0.1)     │
│    Qwen3.5 running on the GPU       │
└─────────────────────────────────────┘
```

**Private by construction, not by promise:**

- The server listens on `127.0.0.1` only — invisible to your network.
- The Office.js runtime is bundled with the app instead of loaded from Microsoft's CDN,
  and its telemetry component is replaced with a stub that does nothing.
- The task pane's Content-Security-Policy allows connections to *one* address: itself.
- Every request needs a random token created fresh each time the app starts.

---

## For developers

```bash
npm install
npm run prepare:assets   # bundle Office.js, generate icons, build the pane
npm run dev              # run from source
npm test                 # 172 tests
npm run dist             # build the signed .app and .dmg
```

| Folder | What's in it |
|---|---|
| `core/` | Engine, model catalog, downloader, HTTPS server, setup — all unit tested |
| `taskpane/` | The Excel-side UI: reading the sheet, previewing changes, chat |
| `desktop/` | Electron shell: control panel, tray, setup wizard |
| `shared/` | Design tokens used by both interfaces |

Two development harnesses render the pane in a browser so you don't have to rebuild into
Excel to see a change:

```bash
npm run build:web -- --dev
# /app/preview.html  — the transcript components against sample data
# /app/debug.html    — the whole pane, with Office and the API stubbed
```

### Two build details worth knowing

- **Packaging happens in `~/Library/Caches`.** If the project lives under `~/Desktop` or
  `~/Documents`, iCloud stamps every file with metadata that `codesign` rejects as
  "detritus" — and because signing writes to the files, they get re-stamped faster than
  any cleanup can strip them. Building outside the synced tree avoids it entirely.
- **The DMG is signed but not notarized.** Notarization needs a paid Apple Developer
  account. Irrelevant on your own machine; sending the DMG to someone else will hit
  Gatekeeper until it is notarized.

Architecture notes and the reasoning behind the trickier decisions are in
[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md).

---

## Known limits

- macOS Apple Silicon only. The core is portable; the installer, certificate trust, and
  add-in registration are macOS-specific.
- One model loaded at a time.
- The chat sees a sample of your rows — the first 60 and the last 15 — plus the true row
  count. Row-by-row jobs (splitting, cleaning) read **every** row regardless; that work
  runs in batches rather than through the chat context.
- A column split can only redistribute what you give it. If the model picks too few output
  columns, leftover text would be dropped — so the app measures how much of each row
  survived and warns you when something was lost.

---

## Licence

MIT.
