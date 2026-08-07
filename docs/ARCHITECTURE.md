# Architecture

How Exovelletron is put together, and why the awkward parts are the way they are.

## The shape of it

```
┌── Microsoft Excel ────────────────────────────┐
│   Task pane (WKWebView)                        │
│     Office.js  ⇄  worksheet                    │
│     loads https://localhost:39217/taskpane     │
└────────────────┬───────────────────────────────┘
                 │ same-origin fetch + SSE
┌────────────────▼───────────────────────────────┐
│ Exovelletron.app (Electron)                    │
│   HTTPS server bound to 127.0.0.1:39217        │
│     /taskpane.html + assets                    │
│     /api/chat        streamed reply            │
│     /api/transform   row-by-row column work    │
│     /api/conversations   saved history         │
│     /api/models      catalog + downloads       │
│   node-llama-cpp → Metal                       │
│   Control panel: setup wizard, model manager   │
└────────────────────────────────────────────────┘
```

## Load-bearing decisions

**Same-origin HTTPS.** Excel renders the task pane as an HTTPS page, so a call to plain
`http://localhost` is blocked as mixed content. Serving the UI *and* the API from one
HTTPS origin removes mixed content and CORS in a single stroke. The app mints a private
CA, trusts it once in the login keychain, and issues a `localhost` leaf from it.

**Proposals, not edits.** The model never writes to the workbook. It calls a tool; the
call is validated into a normalized action; the pane renders that action as a preview and
waits. Applying snapshots the prior cell *formulas* — not values, so a formula is never
silently flattened into a literal — which backs a one-step Undo.

**Two paths for sheet work.** Chat sees a bounded sample of rows. Anything that needs
every row (splitting a column, standardising 500 addresses) goes down a separate batched
path that streams the whole column through the model in small chunks. Trying to serve both
from the chat context would either blow the window or quietly operate on a sample.

## Module boundaries

`core/` is pure and testable with no Electron or Office dependency:

| Module | Responsibility |
|---|---|
| `llm/engine.js` | Model lifecycle, streaming, tool plumbing |
| `llm/actions.js` | The action vocabulary, schemas, validators |
| `llm/transform.js` | Batched row-by-row work, alignment safety |
| `llm/prompt.js` | System prompt, worksheet rendering, token budget |
| `models/` | Catalog, resumable verified downloads, on-disk state |
| `history/` | Conversation persistence |
| `server/` | HTTPS server, static serving, auth, routes |
| `setup/` | Certificates, manifest, preflight checks |

## Findings that shaped the design

These are recorded because none are obvious, and each cost real debugging time.

**Qwen3.5 is a hybrid reasoning model.** On defaults it opens a `<think>` block and can
spend the entire token budget inside it, returning an empty answer. Thoughts are
discouraged via `QwenChatWrapper` plus a zero thought-token budget. Measured: 4.8s and an
empty reply became 0.8s and a real one.

**Blank optional tool parameters caused an infinite retry loop.** Constrained decoding
pushes the model to emit every property in the schema, so it sent `fill: ""`. The
validator rejected that as a malformed colour, the model retried, and the turn burned its
whole budget producing nothing. Blank now means absent, and rejections are capped.

**`ggml_metal_library_init_from_source` is benign.** ggml probes for an optional Metal 4
tensor-API shader variant, fails to build it, disables that path, and runs the standard
Metal kernels. All layers still offload. Filtered out of the log stream rather than
surfaced as an error.

**A 10-year TLS certificate broke the whole add-in.** Excel refused the pane with "the
content is blocked because it isn't signed by a valid security certificate" even though
the issuing root was trusted. Apple caps TLS *server* certificate lifetimes at 398 days
and WebKit enforces it. The CA stays long-lived — re-trusting costs a password prompt —
but the leaf is 397 days and renews from the existing CA, so rotation is silent.

**Assigning `range.formulas` does not fill like the fill handle.** Writing the same
formula string into every cell makes Excel store it verbatim, so every row computed the
anchor row and the column showed one repeated value. The fix is to write the anchor cell
and `copyFrom` it across, which applies Excel's own relative-reference rules.
`offsetFormula` remains, but only for rendering the preview.

**Models pad a fill-down to a round number.** Asked to fill six rows, the model proposed
`E2:E1000`. The prompt states the rule, and validation clamps to the last row of the used
range regardless.

**The `hidden` attribute loses to any explicit `display`.** Several panels are flex
containers, so `panel.hidden = true` did nothing and the history panel rendered
permanently over the transcript. One global `[hidden] { display: none !important }` fixes
the whole class of bug.

**Electron creates its `userData` folder before app code runs.** An all-or-nothing
directory rename during the app rename therefore always saw an existing destination and
skipped, orphaning a 2.9 GB model. Migration is per-item, and treats an empty destination
directory as absent.

**A column split can silently discard data.** If the model chooses five output columns for
data that needs six, the leftover text simply vanishes and the sheet looks fine. The app
compares how many meaningful characters survive each row and warns when a row lost too
much. This is measured only for splits — a rewrite is *supposed* to get shorter.

## The add-in installation constraint

Excel reads sideloaded manifests from
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`. On macOS 26+ that directory
is sealed against every other process.

Tested and refused: the shell, a signed and packaged app, scripted Finder automation, and
**Terminal with Full Disk Access already granted**. No permission prompt is offered. Full
Disk Access is therefore *not* the fix, and the UI deliberately never suggests it.

What works is a folder chosen in a native open panel — that is explicit consent, so macOS
issues a sandbox extension. Measured directly: `EPERM` before the pick, success
immediately after. The wizard detects the blocked state by probing, then opens the picker
already pointing at `wef`. A Finder drag also works and is kept as a fallback.

## Security model

- The server binds `127.0.0.1` only.
- Every API call requires a random per-run session token injected into the pane at serve
  time, plus an origin check. A stale token (the app restarted) makes the pane reload
  itself rather than showing an error.
- CSP pins `connect-src` to our own origin, so the pane cannot reach the internet even if
  something tried.
- Office.js is vendored and served locally; its telemetry sink is replaced with a no-op,
  making telemetry structurally impossible rather than merely disabled.

## Testing

172 tests. Alongside the ordinary coverage, there is a regression test for every finding
above — including one that reads the task pane's routing table and asserts that every
action the model can propose has an apply path, after `split_column` shipped without one.

`scripts/smoke-engine.mjs` exercises the real model end to end.
