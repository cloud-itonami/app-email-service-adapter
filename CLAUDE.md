# etzhayyim-project-email-service-adapter — Outlook OAuth + Mailbox Sync

> **T2 Logical Actor**: Manifest-driven (`20-actors/email-service-adapter/actor-manifest.jsonld`). **PII Tier 3**.

`outlook.etzhayyim.com` (nanoid: `outlook`) — Outlook OAuth + mailbox sync. Mailbox content stored as **Tier 3 (Preferences only)**.

## Lexicons
`emailServiceAdapter/` (2 files): syncMailbox, listSyncs.

## cross-actor
- `gmail` — peer adapter (Gmail OAuth)
- `mailer` — outbound transactional email
- `briefing` — invitation flow
- `kyber-inbox` (`did:web:kyber-qzzg06nh.etzhayyim.com:dept:inbox`, nanoid `inb0x4k2`) — receives derived `com.etzhayyim.apps.kyber.inbox.{emailSignal,calendarSignal,documentSignal}` records (dept-routed signal/noise classification). Wired via `kotodama.jsonld` `derive` rule from `syncJob`. RisingWave sink: `vertex_email_message` / `vertex_calendar_event` / `vertex_office_document` + `edge_kyber_routed`.

## PII (per ADR-0014)
- email body / headers: Tier 3 (Preferences only)
- AT Repo: hashed message-id + sync watermark のみ
- **kyber lexicons (CURRENT DEVIATION)**: `subjectEnc` / `bodyPreviewEnc` / `nameEnc` を AT Record 内に埋め込み中。federation 時に opaque 暗号文が拡散可能性。本来は ADR-0014 厳格遵守で hash + routing metadata のみ AT Record、暗号化本文は Preferences。`deps.toml [[migrations.outlook-kyber-integration-2026-04-14]]` で follow-up tracking。
- **暗号化レベル (CURRENT DEVIATION)**: `signal:v1:{base64}` は demo envelope (base64 = encoding ≠ encryption)。本番は `10-protocol/wproto/src/signal.ts` `encryptFieldVal()` (Signal Protocol X3DH+DoubleRatchet) に差替え必要。

## Design
- ADR-0014: PII Tier 3 + Cohort-First Pattern
- `30-graph/graph-schema/migrations/0027_outlook_kyber_integration.ts` (email/event/contact + dept routing)
- `30-graph/graph-schema/migrations/0028_office_documents_kyber.ts` (OneDrive/SharePoint office documents)
- `00-contracts/lexicons/com/etzhayyim/apps/kyber/inbox/` (3 record lexicons)
