# Operator quickstart

What you can actually run in this checkout, what you cannot, and how to tell the
difference without taking anyone's word for it.

Measured 2026-08-16 at `7fdd22f`. Every command below was run to produce the
numbers quoted; every claim is re-checkable with the one command in §2.

---

## 1. What this repository is

A **seed**, extracted verbatim from `etzhayyim/root` (see `migration.edn`). Two
appview components for external email providers:

| path | what is there |
|---|---|
| `appview/outlook-mcp-component/` | a SvelteKit UI, a `src/app.ts` edge facade, `wrangler.jsonc`, a Playwright suite |
| `appview/gmail-mcp-component/` | a `kotodama.jsonld` manifest and a 6.7 MB committed binary. **No source.** |

24 tracked files. It is not a running service, and nothing in it is deployed —
§4 shows how that was established rather than assumed.

`README.edn` stays the canonical metadata (`:canonical-metadata :edn`). This
file and `README.md` are prose about it, not a second source of truth.

---

## 2. Verify the checkout

One command, no network, no install, ~5 s:

```bash
nbb scripts/verify-surface-agreement.cljs            # 19 facts
nbb scripts/verify-surface-agreement.cljs --verbose  # print all of them
nbb scripts/verify-surface-agreement.cljs --live     # 21 — also resolves the hostnames
```

Exit `0` every pinned fact holds · `1` one moved · `2` **could not answer**, an
input was missing. `2` is deliberately not `0`: a run that could not read its
inputs must not be reportable as agreement.

To check the checker itself:

```bash
nbb scripts/mutate-surface-agreement.cljs            # 13 mutations, ~5 s
```

It breaks each pinned fact in a scratch copy and requires the verifier to exit 1
**naming the check that pins that fact** — an exit code alone would also be
produced by a mutation that broke something unrelated. Three of the thirteen
legitimately fire two checks each; those couplings are annotated at the case.

> The verifier is green today and pins the divergence below *as it currently
> stands*. It is a ratchet, not a verdict: it does not decide which surface
> ought to win, and it turns red the moment anyone changes one of the four
> descriptions without changing the others and this document.

---

## 3. Four descriptions of one HTTP surface, no two alike

`outlook.etzhayyim.com` is described in four places. They disagree.

| # | file | says the surface is |
|---|---|---|
| 1 | `wrangler.jsonc` `main` | `svelte/.svelte-kit/cloudflare/_worker.js` — the SvelteKit build output |
| 2 | `kotodama.jsonld` `component.path` | `src/app.ts` |
| 3 | `src/app.ts` | `/health`, `/_app/meta`, and `/xrpc/` for prefix **`com.etzhayyim.apps.outlook.`** |
| 4 | `svelte/src/App.svelte` | `/xrpc/etzhayyim.outlook.v1.OutlookService` **`.`** `Method` |
| — | `svelte/e2e/outlook.spec.ts` | `/xrpc/etzhayyim.outlook.v1.OutlookService` **`/`** `Method`, plus `/_worker/health` |

Consequences, each pinned by a named check:

**The file a reader opens is not the file that runs.** `wrangler.jsonc` deploys
the SvelteKit output (1); `kotodama.jsonld` names the facade (2). There are
**zero** server routes under `svelte/src/` — no `+server.ts`, no
`hooks.server.ts`, no `+page.server.ts` — so the deployed worker serves the SPA
and nothing else. `src/app.ts` is never imported by anything. Its `/health`
answers nobody.

> This is not local colour. The workspace's own fleet detector,
> `scripts/verify-appview-facade.cljs`, reports
> `health-only-in-undeployed-facade` for this repository, and for 88 others out
> of 331 appviews scanned.

**The browser calls a namespace the facade does not proxy.** The UI calls
`etzhayyim.outlook.v1.OutlookService`; the facade proxies
`com.etzhayyim.apps.outlook.`. So even if someone "fixed" the deploy to point at
`src/app.ts`, **zero** of the UI's six calls would be proxied — they would fall
through to the 404. Fixing the deploy target alone does not connect these two.

The six the browser makes: `GetAuthStatus`, `GetConnection`, `ExchangeCode`,
`StartAuth`, `SyncMailbox`, `Disconnect`.

**The test suite exercises a third surface.** Its 13 tests across 3 describe
blocks call `GetOAuthConfig`, `GetConnection`, `card.home`, `card.compose`,
`card.action`. It shares exactly **one method name** with the UI
(`GetConnection`) and **zero URLs**, because the UI joins with `.` and the suite
joins with `/`.

Where a path *is* shared, the expected body still differs: the suite asserts
`status: "ok"` and `app: "outlook"` on `/health`, and `appId` on `/_app/meta`;
`src/app.ts` returns `ok: true` and `nanoid` and has no `status`, `app`, or
`appId` field, nor a `/_worker/health` route at all. Pointing the deploy at the
facade would leave the suite red on shape as well as on routing.

---

## 4. What you cannot do here, and how that was established

Not "unsupported" — measured. Each has a check id in §2's verifier.

**You cannot build.** `svelte/package.json` pins
`"@etzhayyim/design-system": "workspace:*"`. The `workspace:` protocol requires
a workspace root; this repository has no `pnpm-workspace.yaml` and no root
`package.json`, so the install that would produce `main` cannot resolve here.
The dependency is also never imported by any source file, so the pin currently
buys nothing. → `workspace-protocol-with-no-workspace-root`

**You cannot deploy to the documented host.** `wrangler.jsonc` routes
`outlook.etzhayyim.com/*`. That name is **NXDOMAIN**:

```console
$ host outlook.etzhayyim.com
Host outlook.etzhayyim.com not found: 3(NXDOMAIN)
$ curl -s -o /dev/null -w '%{http_code}\n' https://outlook.etzhayyim.com/
000
```

Same for `gmail.etzhayyim.com`. `etzhayyim.com` itself resolves, so this is the
subdomain, not the zone. The workspace surface index
(`90-docs/surface/surface.datoms.edn`) has **0 rows** for either host.
→ `dns-outlook`, `dns-gmail` (under `--live`)

Both `@id`s — `did:web:outlook.etzhayyim.com` and `did:web:gmail.etzhayyim.com` —
therefore do not resolve either, since `did:web` resolution is an HTTPS GET
against that host.

**You cannot run the e2e suite meaningfully.** `playwright.config.ts` sets
`baseURL: https://outlook.etzhayyim.com`. All 13 tests fail at DNS. They would
fail on routing and on body shape too (§3), so a green run here would mean the
suite had been pointed somewhere else.

**The gmail component is not the artifact its manifest names.**
`kotodama.jsonld` declares `component.path = "component.wasm"`. That file does
not exist. What is committed is `gmail-mcp-component`, 6,712,786 bytes:

```console
$ file appview/gmail-mcp-component/gmail-mcp-component
Mach-O 64-bit executable arm64
$ shasum -a 256 appview/gmail-mcp-component/gmail-mcp-component
b2d054a5f969f17716db1e247803865c94637fa0573c47f327614eb6699c06ff
```

A macOS-arm64 native executable — not wasm (magic `cffaedfe`, not `0061736d`),
and not runnable on the SpinApp/Worker runtime `PROJECT.jsonld` claims for it.
It is also a 6.7 MB binary in git history, which the workspace routes to
DataLad/git-annex instead (skill `large-binary-datalad`).
→ `gmail-manifest-component-path`, `gmail-manifest-component-is-absent`,
`gmail-binary-is-mach-o-not-wasm`, `gmail-binary-size`

**The gmail manifest subscribes to collections that cannot exist.** Its three
`triggers.subscribeRepos.collections` read
`com.etzhayyim.apps.emailUserviceUadapter.*`. The outlook manifest, for the same
app, reads `com.etzhayyim.apps.emailServiceAdapter.*`. A codemod appears to have
replaced `-` with `U` in `email-service-adapter`; the resulting NSIDs match
nothing. → `gmail-collections-are-mangled`

---

## 5. Known-stale references

Files named by the docs in this repository that are not in it — carried over
from the monorepo this was extracted from, so treat them as pointers into
`etzhayyim/root`, not local paths:

- `NOTICE` and `MIGRATION-TODO.md` cite `CHARTER-RIDER.md` — absent here.
- `CLAUDE.md` cites `20-actors/…`, `30-graph/…`, `00-contracts/lexicons/…`,
  `10-protocol/wproto/src/signal.ts`, `deps.toml` — all absent here.
- `CLAUDE.md` says the lexicons are "`emailServiceAdapter/` (2 files)". There is
  no lexicon directory in this repository.
- `MIGRATION-TODO.md`'s codemod evidence cites
  `appview/outlook-mcp-component/static-ui/_app/immutable/chunks/By41dYui.js`.
  There is no `static-ui/` here, so that scan cannot be reproduced from this
  checkout.

Two open deviations are recorded in `CLAUDE.md` and are **not** resolved:
`signal:v1:{base64}` is an encoding, not encryption; and `subjectEnc` /
`bodyPreviewEnc` / `nameEnc` are embedded in AT Records against ADR-0014.

Three other copies of this app exist in the workspace
(`etzhayyim/com-etzhayyim-app-email-service-adapter`,
`etzhayyim/root/60-apps/etzhayyim-project-email-service-adapter`, and worktrees
under it). The first carries the same facade finding. A fix here does not
propagate to them.

---

## 6. If you change any of this

The verifier pins the *current* shape, so a real fix will turn it red. That is
intended. When it does:

1. Decide which of the four descriptions is authoritative. The verifier will not
   decide for you, and it should not — connecting the UI to a working backend
   needs the pod-side LangServer this repository does not contain.
2. Update this document to match.
3. Only then update `:expect` in `scripts/verify-surface-agreement.cljs`, and
   add or adjust the case in `scripts/mutate-surface-agreement.cljs`.

Updating `:expect` alone silently decouples the doc from the code, which is the
failure this pair exists to prevent.
