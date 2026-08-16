# app-email-service-adapter

**Outlook and Gmail mailbox-sync appviews for the etzhayyim / cloud-itonami
substrate.** Two external email providers behind one repository: an OAuth
connect flow, a mailbox/calendar sync trigger, and the manifests that describe
them to the actor plane.

`README.edn` remains the canonical metadata for this repository
(`:canonical-metadata :edn`). This file is prose, not a second source of truth.

## Status: seed, not a running service

Extracted verbatim from `etzhayyim/root` (`migration.edn`) and **not yet
reconnected**. As measured on 2026-08-16 at `7fdd22f`:

- Nothing here is deployed. `outlook.etzhayyim.com` and `gmail.etzhayyim.com`
  are **NXDOMAIN**, and the workspace surface index has no rows for either.
- The build cannot run in this checkout: `svelte/package.json` uses
  `workspace:*` and the repository has no workspace root.
- Four files describe the HTTP surface and **no two of them agree** — the
  deployed entry point, the manifest, the facade, the browser, and the test
  suite each name a different one.
- The gmail component is a 6.7 MB macOS-arm64 executable committed to git; its
  manifest names a `component.wasm` that does not exist.

None of that is a reason to avoid the repository — it is the work. It is
written down so the next reader starts from the measurements instead of
rediscovering them.

## Start here

```bash
nbb scripts/verify-surface-agreement.cljs --verbose   # ~5 s, no network
```

19 pinned facts (21 with `--live`, which also resolves the hostnames). Exit `0`
they hold · `1` one moved · `2` could not answer.

Then read **[`docs/operator-quickstart.md`](docs/operator-quickstart.md)**,
which explains each fact, what you can and cannot run, and what to do when the
verifier goes red.

To check the checker: `nbb scripts/mutate-surface-agreement.cljs` (13
mutations, ~5 s) breaks each fact in a scratch copy and requires the verifier
to name the specific check that pins it.

## Layout

| path | |
|---|---|
| `appview/outlook-mcp-component/` | SvelteKit UI, `src/app.ts` edge facade, `wrangler.jsonc`, Playwright suite |
| `appview/gmail-mcp-component/` | `kotodama.jsonld` + committed binary; no source |
| `docs/operator-quickstart.md` | what runs, what does not, and the evidence |
| `scripts/` | the verifier and its mutation harness |
| `CLAUDE.md` | actor/PII design notes, incl. two unresolved deviations |
| `MIGRATION-TODO.md` | substrate-boundary checklist inherited from the seed |

Licensed Apache-2.0 with the etzhayyim Charter Compliance Rider v3.1; see
`NOTICE`. The rider text itself lives in `etzhayyim/root`, not here.
