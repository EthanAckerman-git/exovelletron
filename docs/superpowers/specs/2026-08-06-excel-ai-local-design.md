# Excel AI Local — design

**Date:** 2026-08-06
**Status:** implemented

A fully offline AI assistant for Microsoft Excel on macOS (Apple Silicon): a chat task
pane backed by a local model, with a desktop app that manages setup and model downloads.

## Requirements

Fixed with the user before implementation:

| Decision | Choice |
|---|---|
| Platform | macOS, Apple Silicon only |
| Surface | Chat side pane (no in-cell AI functions) |
| Default model | Qwen3.5 4B, ~2.9 GB |
| Install | One installer; the app performs setup itself |

Non-negotiable: no network access after setup, no account, no telemetry.

## Architecture

```
┌── Microsoft Excel ─────────────────────────┐
│   Task pane (WKWebView)                     │
│     Office.js ⇄ worksheet                   │
│     loads https://localhost:39217           │
└───────────────┬─────────────────────────────┘
                │ same-origin fetch + SSE
┌───────────────▼─────────────────────────────┐
│ Excel AI Local.app (Electron)               │
│   HTTPS server on 127.0.0.1:39217           │
│     /taskpane.html, /api/chat, /api/models  │
│   node-llama-cpp → Metal                    │
│   Control panel: setup wizard, models       │
└─────────────────────────────────────────────┘
```

**Same-origin HTTPS is the load-bearing decision.** Excel renders the pane as an HTTPS
page, so `http://localhost` is blocked as mixed content. Serving the UI *and* the API from
one HTTPS origin removes mixed content and CORS together. The app mints a private CA,
trusts it in the login keychain, and issues a `localhost` leaf.

## Module boundaries

`core/` holds everything pure and testable, with no Electron or Office dependency:

- `llm/engine.js` — model lifecycle, streaming, tool-call plumbing
- `llm/actions.js` — the action vocabulary, schemas, and validators
- `llm/prompt.js` — system prompt, worksheet rendering, token budgeting
- `models/catalog.js` · `downloader.js` · `store.js` — catalog, resumable verified
  downloads, on-disk state
- `server/` — HTTPS server, static serving, auth, routes
- `setup/` — certificates, manifest, preflight checks

`taskpane/` is the Excel-side UI; `desktop/` is the Electron shell; `shared/` holds design
tokens used by both.

## Safety model

The model never writes to the workbook. It calls a tool; the call is validated into a
normalized action; the pane renders that action as a preview and waits for **Apply**.
Applying snapshots the prior cell *formulas* (not values, so a formula is not silently
flattened into a literal) to back a one-step **Undo**.

Validation rejects malformed ranges, grids whose shape disagrees with the target, and
writes above 50,000 cells.

## Offline and security guarantees

- Office.js is vendored and served from our origin; its telemetry sink is replaced with a
  no-op stub, so the runtime structurally cannot emit telemetry.
- CSP pins `connect-src` to our own origin.
- The server binds `127.0.0.1` only.
- Every API call needs a random per-run session token injected into the pane at serve time,
  plus an origin check.

## Findings that changed the design

Three things were discovered during implementation and are recorded because they are not
obvious and would otherwise be re-litigated.

**1. Qwen3.5 is a hybrid reasoning model.** Left on defaults it opened a `<think>` block
and spent the entire token budget inside it, returning an empty answer. Fixed by
constructing `QwenChatWrapper` with `thoughts: "discourage"` plus a zero thought-token
budget. Measured: 4.8s and an empty reply became 0.8s and a real one.

**2. Blank optional tool parameters caused an infinite retry loop.** Constrained decoding
pushes the model to emit every property in the schema, so it sent `fill: ""`. The validator
rejected that as a malformed colour, the model retried, and the turn burned its whole
budget producing nothing. Blank now means absent, and rejections are capped per turn.

**3. The `ggml_metal_library_init_from_source` error is benign.** ggml probes for an
optional Metal 4 tensor-API shader variant, fails to build it, disables that path, and runs
the standard Metal kernels. All 33 layers still offload; measured 61 tok/s. It is filtered
out of the log stream rather than surfaced as an error.

## The add-in installation constraint

Excel reads sideloaded manifests from
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`. On macOS 26+ that directory
is sealed against every other process.

Tested and refused: the shell, a signed and packaged app, scripted Finder automation, and
**Terminal with Full Disk Access already granted**. No TCC prompt is offered. Full Disk
Access is therefore *not* the fix, and the UI deliberately does not suggest it — sending
someone into their security settings for a remedy that fails is worse than asking for the
alternative.

What works is a drag in Finder, because the drop carries user intent and macOS issues a
sandbox extension for it. The wizard detects the blocked state by probing, then stages the
manifest, reveals it selected, and opens the destination alongside it. One drag, once.

This is the single place the "installer does everything" goal could not be fully met, and
the limitation is the platform's rather than the design's.

## Testing

90 unit tests across the action validators, range parsing, formula reference shifting,
config normalization, manifest generation, path-traversal defence, API authorisation,
catalog integrity, prompt budgeting, and packaging completeness.

Regression tests exist specifically for the three findings above and for the packaging bug
where `shared/tokens.css` was referenced at runtime but excluded from the bundle — which
only manifested once packaged, as serif fallback text and unstyled controls.

`scripts/smoke-engine.mjs` exercises the real model end to end: load, stream, tool call,
and action validation.
