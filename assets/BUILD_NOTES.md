---
name: project-value-report-generic
description: Generic TrendAI Vision One value-report deck builder for any customer, via the /value-report skill — asks who the customer is and who's presenting instead of being hardcoded to one tenant
metadata:
  type: project
  originSessionId: c58c7bf4-e092-40d2-8d07-87467022f0e7
---

**Trigger: "Run a value report" / "Build a value report for <customer>"** (also
the `/value-report` skill). Created 2026-07-01 by generalizing
[[project_intralot_value_report_deck]] — same 10-slide design/pipeline, but not
locked to one customer.

- Master template dir: `~/Documents/Claude/value_report/` — `build_value_report.py`
  (fully generic: `OUTPUT` is computed at runtime from `D['customer']`, so it's
  copied verbatim into each customer's folder with **no per-customer script
  edits needed** for the output path), `template.pptx`, a placeholder
  `value_report_data_template.json` skeleton, `RUNBOOK_TEMPLATE.md` (generic
  runbook with `{{CUSTOMER}}` tokens).
- Per-customer working dir: `~/Documents/<Customer>/value_report/` — same
  layout as the existing JHC/Intralot folders (build script + template.pptx +
  value_report_data.json + a generated `<Customer> Value Report.md` runbook).
  Created on first run for that customer, reused on reruns (rebuild preserves
  slide 1 + last slide, same as JHC/Intralot).

**Key design point:** the skill ALWAYS asks two questions before touching any
data — which customer, and who's presenting — rather than assuming or reusing
a prior answer. Doesn't force these into rigid multiple-choice UI; uses
`AskUserQuestion` with discovered-existing-customers as quick-picks (Other
covers anything else) when that fits, or just asks in plain conversation when
it doesn't (e.g. no customers exist yet).

**Gotcha caught during setup:** the placeholder skeleton originally used `0`
for several chart-scaling denominators (`action_max`, `threat_max`,
`credits.purchased`) which crashes `build_value_report.py` with a
`ZeroDivisionError` if someone tries to build before pulling real data. Fixed
by setting those three placeholder fields to `1` in
`value_report_data_template.json`. Smoke-tested with a full build + a rebuild
(preserve-slides path) — both work.

Existing dedicated skills (`/jhc-value-report`, `/intralot-value-report`) are
untouched and still work exactly as before — this is an additional, separate
skill, not a replacement.

**2026-07-01 enhancements (Sierra-Cedar, LLC run) — now a 13-slide deck:**
- **Slide 2 (Cyber) redesigned** from a 3-column (Devices/Vulnerabilities/Accounts)
  layout to a 2x3 grid covering all 6 risk categories: Devices, Internet-Facing
  Assets, Accounts, Applications, Cloud Assets, Vulnerabilities. Each panel =
  headerbar + risk-level chip (green/orange/red) + headline count + breakdown
  sub-line. Data shape: `D['cyber']['categories']` = list of 6
  `{title, risk_level, headline, sub}` dicts (replaces the old flat
  discovered/assessable/managed/domain/service/cves_total/mttp fields).
- **New "Risk Factors Detail" slides (3-5)**, generated dynamically 2-categories-
  per-slide from `D['risk_detail']` (a flat list of
  `{title, risk_level, factors: [[count,label],...], note?}`). `main()` builds
  `pairs = [risk_detail[i:i+2] ...]` and splices one `build_risk_detail_slide`
  call per pair into `content_builders` right after `build_cyber` — slide COUNT
  now varies with `len(risk_detail)`, so `expected` (used by preserve-mode) is
  computed fresh each run from the current `content_builders` length. If
  `len(risk_detail)` changes between runs, preserve-mode's slide-count check
  fails and it silently falls back to a full rebuild (hand-edited presenter/
  team fields need re-entering after that — documented in the runbook).
  A category with empty `factors` renders a green-check "no risk factors /
  no data source" note instead of an empty table (matches what the live
  dashboard itself shows for an unconfigured category).
- **New Vision One API endpoints discovered** (all same-origin, `uic-token`
  header from `localStorage.getItem('uic-token')`, sync XHR):
  - `GET /public/ass/api/v1/trilogy/riskOverview/internet/riskLevel` → Internet-Facing Assets (`ipRisk`/`domainRisk` total+low/medium/high)
  - `GET /public/ass/api/v1/trilogy/riskOverview/appAssets/riskLevelBar` → Applications total+low/medium/high (only fires after clicking the Applications card — not prefetched)
  - `GET /public/sase/api/v1/bigtable/forward/v1/cloud-assets/asset-level-bar` → Cloud Assets total+low/medium/high (only fires after clicking the Cloud Assets card)
  - `GET /public/ass/api/v1/trilogy/riskOverview/riskFactor?assetType=<device|internet|account|cloudAsset>` → per-category risk factor counts (cves, cyberThreats, compromiseEvents, complianceViolations, sentryCves, functionCves, imageCves, highRiskConfigurationRisks, etc.)
  - `GET /public/cta/api/v1/trilogy/riskOverview/riskFactors?assetType=<endpoint|account>` → richer per-category breakdown (e.g. account: weakSignInCnt, excessiveCnt, legacyProtocolCnt, excessivePrivilegeCnt, cyberThreatCnt) — endpoint/account only, `application`/`cloudAsset`/`device`/`internet` return 400 on this particular endpoint.
  - No working `assetType` found for Applications on the `riskFactor` endpoint when no app-inventory data source is connected — that's expected, not a bug; render the empty-state note instead.
- **Two site-wide Vision One dashboard quirks discovered (apply to ANY tenant,
  not just Sierra-Cedar)** — now documented in `RUNBOOK_TEMPLATE.md`:
  1. A `.iframe-overlay` div sits on top of the whole Cyber Risk Overview iframe
     and silently swallows clicks (`browser_interact` reports hitting
     `iframe#__SASE_ES_CONTAINER` regardless of target). Fix once per session:
     `document.querySelector('.iframe-overlay').style.pointerEvents='none'`.
  2. Even with the overlay disabled, clicking the category cards by screen
     coordinate is unreliable (silently selects the wrong tab). Reliable fix:
     find the element by exact text inside the dashboard's iframe
     (`document.querySelectorAll('iframe')` → `f.contentDocument` →
     `querySelectorAll('*')` filtered by `textContent.trim()===label`), walk up
     exactly 3 parents to the ancestor with a real `onclick`, and call
     `.click()` on it directly via `browser_evaluate` — full snippet in the
     runbook. This also explains why earlier attempts to reach the Email &
     Collaboration Security dashboard failed (never confirmed fixed).
- Verification workflow for this session: no PowerPoint/LibreOffice CLI
  renderer on this machine, so slide previews were generated by opening the
  deck in **Keynote** via AppleScript (`open pptxAlias` then
  `export theDoc to (POSIX file outFolder) as slide images`) and reading the
  resulting PNGs — more reliable than trying to script PowerPoint's own
  AppleScript dictionary (its `save ... as save as PNG` command didn't
  actually produce files in testing).

**2026-07-01, same day — two more additions, deck now 14 slides:**
- **MDR now always asks first** (see [[feedback_value_report_mdr_ask_first]]):
  `build_mdr` checks `D['mdr'].get('applicable', True)` — if the user says the
  customer doesn't have MDR, set `applicable: false` (+ optional
  `not_applicable_note`) and it renders a clean "not part of this
  subscription" card instead of an empty funnel.
- **New slide 7: "Data Source and Log Management"**, inserted between MDR and
  Identity (`build_data_source`, wired into `content_builders` right after
  `build_mdr`). Shows 4 headline metrics (New Analytic Ingestion, New
  Archival Ingestion, Total Extended Analytic Retention, Total Extended
  Archival Retention) in a stat-row matching the Credits slide's style, plus a
  full "Ingestion and Retention by Data Source" table (name, category,
  ingestion GB, retention GB — computed by summing analytic+compliance
  ingestion and all four retention fields per source). Data shape:
  `D['data_source'] = {period, metrics: [{label,value},...], sources: [[name,category,ingestion,retention],...]}`.
  Source page: **Agentic SIEM and XDR → Data Source and Log Management →
  Data Monitoring tab** (NOT the "Data sources and retention" tab on the same
  page — that's just a per-source retention-period config list, no usage
  figures). APIs: `GET /ui/dsr/uss/api/v1/data_usage/overview` (headline
  tiles — a **daily snapshot**, not a date range) and
  `POST /ui/dsr/uss/api/v1/data_usage/data_usage_by_source` (per-source
  breakdown — the page sends its own request body with all configured source
  codes; just read the response rather than hand-constructing the body).
  Reached this page via the same `data-menu-id` click trick as Credits/Threat
  Intel (`AGENTIC SIEM AND XDR` popup → `data-menu-id="10044"`), then had to
  click the "Data Monitoring" tab specifically — clicking the exact leaf text
  node itself worked (no parent-walk needed here, unlike the category cards).
  Sierra-Cedar quirk: "Server & Workload Protection" and "Trend Cloud One -
  Endpoint & Workload Security" show up as two identical-figure rows in the
  by-source table (pCodes `sds_v1es` / `sds`) — looks like one underlying
  service double-counted under two product-naming conventions; left both in
  as returned rather than silently merging.
- **Slide order changed same day:** Intelligence Reports moved from its old
  spot (after Email, before Credits) to right after Data Source and Log
  Management (before Identity) — per Deb's explicit request.

**2026-07-01, same day — third addition, deck now 15 slides:**
- **New slide 9: "Workflow and Automation"**, inserted between Intelligence
  Reports and Identity (`build_workflow_automation`, wired into
  `content_builders` right after `build_intel`). Two side-by-side cards:
  "SECURITY PLAYBOOKS" (list, or a green-check "none configured" note if
  empty) and "THIRD-PARTY INTEGRATIONS · CONFIGURED" (list). Data shape:
  `D['workflow_automation'] = {playbooks: [...], playbooks_note, integrations:
  [...], integrations_note}`. Reuses the same empty-state / striped-row
  pattern as `build_risk_detail_slide`/`build_data_source`.
  - **Security Playbooks source:** Workflow and Automation → **Playbooks**
    tab (NOT "Execution Results" — that's just execution history and can say
    "No execution results" even when playbooks exist; "No playbooks created"
    on the Playbooks tab is the real "none configured" signal).
  - **Sidebar navigation gotcha (same class of bug as the unresolved Email
    dashboard issue):** the "Workflow and Automation" left-rail group's
    submenu popup would not mount in the DOM under single click, double
    click, hover, or hover+mouse_move — tried extensively. **Fix: use
    `window.__MENU_FEATURES`**, an undocumented global JS array exposing
    `{id, title, path, group, description}` for every menu item in the app
    (present regardless of whether any popup is mounted). Find the entry by
    `title` (e.g. `"Security Playbooks"` → `path: "/app/workflow-automation-eco"`)
    and navigate directly via `location.hash` / URL — bypasses the need to
    click through the sidebar at all. Worth trying this FIRST on any future
    "popup won't mount" sidebar problem (e.g. if the Email & Collaboration
    Security dashboard is revisited) before burning time on click techniques.
  - **Third-Party Integrations source:** Cyber Risk Exposure Management →
    click the "Data sources" button (top of dashboard) → navigates to
    `#/app/dsr/asrm`. `GET /ui/dsr/uss/api/v2/data_source/data_sources`,
    filter `status==1` (Configured), take `sourceName`. Deb's instruction was
    literal — list the full Configured set from this page, not a narrower
    hand-picked "genuinely third-party vendor" subset.
  - Sierra-Cedar result: 0 playbooks (confirmed via UI, not a pull failure);
    15 Configured integrations.
- `content_builders` is now `[build_cyber] + risk_detail_builders +
  [build_mdr, build_data_source, build_intel, build_workflow_automation,
  build_identity, build_endpoint, build_email, build_credits,
  build_eoy_prediction]`. Slide order at that point (15 total): Title(1),
  Cyber(2), Risk Factors Detail(3-5), MDR(6), Data Source and Log
  Management(7), Intelligence Reports(8), Workflow and Automation(9),
  Identity(10), Endpoint(11), Email(12), Credits(13), EOY(14), Team(15).

**2026-07-02 — fourth addition, deck now 16 slides:**
- **New slide 10: "AI Security"**, inserted between Workflow and Automation
  and Identity (`build_ai_security`, wired into `content_builders` right
  after `build_workflow_automation`). Data shape:
  `D['ai_security'] = {applicable: bool, not_applicable_note?, modules?}`.
  - **`applicable: false` (the common case today)** renders a plain centered
    "Not Applicable" card with an explanatory note — same visual language as
    `build_mdr`'s not-applicable branch. This is the DEFAULT (missing key ⇒
    false), the inverse of MDR's default — most tenants won't have adopted
    AI Security yet.
  - **`applicable: true`** renders one card per module in `modules` (list of
    `{name, status, metrics: [[label,value],...]}`) — a headerbar + status
    chip (`STATUS_CHIP = {"In Use": green, "Not Configured": gray}`) + a
    stacked metrics list. This branch is UNTESTED against real data (no
    customer has adopted AI Security yet) — treat the layout as provisional
    and adjust once a real tenant has live figures to render.
  - **How to check a tenant:** `window.__MENU_FEATURES.filter(x => x.group
    === "AI Security")` returns the exact paths for the 3 sub-modules
    (AI Security Blueprint `/app/ai-security-blueprint`, AI Application
    Security `/app/ai-app-sec`, AI Secure Access `/app/zero/overview2` —
    note this one redirects to `/app/zero/overview`, the same Zero Trust
    Secure Access overview page, since AI Secure Access is a feature within
    that module rather than a separate dashboard). Navigate to each by URL
    directly (same `__MENU_FEATURES` trick as Workflow and Automation) —
    **if any/all show a "Start paid usage" / "Start free trial" / "Opt in"
    onboarding landing page** (recognizable by mock/blurred screenshot
    numbers rather than a live dashboard), that module is not in use. Don't
    mistake the onboarding page's illustrative screenshot figures for real
    usage data.
  - Sierra-Cedar result: all three sub-modules showed onboarding pages →
    `applicable: false`.
- `content_builders` was `[build_cyber] + risk_detail_builders +
  [build_mdr, build_data_source, build_intel, build_workflow_automation,
  build_ai_security, build_identity, build_endpoint, build_email,
  build_credits, build_eoy_prediction]`. Slide order at that point (16
  total): Title(1), Cyber(2), Risk Factors Detail(3-5), MDR(6), Data Source
  and Log Management(7), Intelligence Reports(8), Workflow and Automation(9),
  AI Security(10), Identity(11), Endpoint(12), Email(13), Credits(14),
  EOY(15), Team(16).

**2026-07-02 — fifth addition, deck now 17 slides:**
- **New slide 12: "Data Security"**, inserted right after Identity Security
  Posture (`build_data_security`, wired into `content_builders` right after
  `build_identity`). Data shape: `D['data_security'] = {score, risk_level,
  total_assets, sensitive_assets, monitored_assets, categories: [[name,
  count],...], top_assets: [[name,type,sensitive_types_joined,score],...]}`.
  Layout: a headline stat row (total/sensitive/monitored assets, same style
  as Credits/EOY), then two side-by-side cards — left "SENSITIVE DATA
  DETECTIONS BY CATEGORY" (horizontal bar chart, reuses the Email slide's
  bar-row pattern), right "TOP RISKY ASSETS WITH SENSITIVE DATA" (4-column
  table: Asset Name, Type, Sensitive Data Type, Risk Score — reuses the Data
  Source table's striped-row pattern).
  - **Source page:** Cyber Risk Exposure Management → **Data Security
    Posture** (`#/app/data-posture`; `window.__MENU_FEATURES` group
    `"Data Security"` also lists `Data Inventory` `/app/data-inventory`,
    `Data Policy` `/app/data-policy`, and `Sensitive Data Classification`
    `/app/sensitive-data-classification` as sibling pages, not used for this
    slide).
  - **APIs (same-origin, fire on page load — no click-through needed):**
    `GET /public/sase/api/v1/bigtable/forward/v1/sensitive-assets/summary?storageType=all`
    → `statistics` (`totalAssets`, `sensitiveAssets`, `monitoredAssets`) +
    `distributions` (`sensitiveType`: personal/financial/credentials/other/
    custom → `count`) for the headline stats and category bar chart.
    `GET /public/sase/api/v1/bigtable/forward/v1/sensitive-assets/list?storageType=all&limit=5`
    → `assets[]` (`assetName`, `assetType`, `sensitiveTypes[]`,
    `latestRiskScore`), already sorted by risk score descending — matches
    the on-screen "Top Risky Assets with Sensitive Data" table exactly. Join
    `sensitiveTypes` with `", "` for the table's display column.
  - Sierra-Cedar result: score 67.2/100 (Medium), 13,907 total assets, 12
    sensitive, 0 monitored; categories Personal 4 / Financial 12 /
    Credentials 1 / Other 0 / Custom 0; top asset UCPPUD02 at 78.0.
- `content_builders` was `[build_cyber] + risk_detail_builders +
  [build_mdr, build_data_source, build_intel, build_workflow_automation,
  build_ai_security, build_identity, build_data_security, build_endpoint,
  build_email, build_credits, build_eoy_prediction]`. Slide order at that
  point (17 total): Title(1), Cyber(2), Risk Factors Detail(3-5), MDR(6),
  Data Source and Log Management(7), Intelligence Reports(8), Workflow and
  Automation(9), AI Security(10), Identity(11), Data Security(12),
  Endpoint(13), Email(14), Credits(15), EOY(16), Team(17).

**2026-07-02 — sixth addition, deck now 18 slides:**
- **New slide 11: "Zero Trust"**, inserted right after AI Security
  (`build_zero_trust`, wired into `content_builders` right after
  `build_ai_security`) — Deb asked for this as a follow-up ("sorry I missed
  a slide") after the AI Security addition. Code is a near-exact copy of
  `build_ai_security`'s structure (same not-applicable card, same per-module
  status-card grid for the applicable branch) — deliberately kept identical
  since both slides answer the same question ("is this newer module in
  use?") for a different Vision One product family. Data shape:
  `D['zero_trust'] = {applicable: bool, not_applicable_note?, modules?}`,
  same semantics as `ai_security` (default `applicable: false`).
  - **How to check a tenant:** `window.__MENU_FEATURES.filter(x => x.group
    === "Zero Trust Secure Access")` — entry point is **Secure Access
    Overview** `/app/zero/overview`. If it shows a "Start paid usage" /
    "Start free trial" onboarding landing page (same visual pattern as AI
    Security's onboarding pages), the module is not in use. Confirmed
    cross-link: AI Secure Access's own path (`/app/zero/overview2`, listed
    under the "AI Security" `__MENU_FEATURES` group) redirects to this exact
    Zero Trust overview page — AI Secure Access is a *feature within* Zero
    Trust Secure Access, not an independently-hosted dashboard, so checking
    one often tells you about the other.
  - Sierra-Cedar result: onboarding page shown → `applicable: false`.
- `content_builders` was `[build_cyber] + risk_detail_builders +
  [build_mdr, build_data_source, build_intel, build_workflow_automation,
  build_ai_security, build_zero_trust, build_identity, build_data_security,
  build_endpoint, build_email, build_credits, build_eoy_prediction]`. Slide
  order at that point (18 total): Title(1), Cyber(2), Risk Factors
  Detail(3-5), MDR(6), Data Source and Log Management(7), Intelligence
  Reports(8), Workflow and Automation(9), AI Security(10), Zero Trust(11),
  Identity(12), Data Security(13), Endpoint(14), Email(15), Credits(16),
  EOY(17), Team(18).

**2026-07-02 — seventh change, deck now 19 slides (net +1 — one slide
removed, two added):**
- **Removed the single generic "Endpoint Security" slide** (`build_endpoint`,
  and its `D['endpoint']` data key) — deleted entirely, not just unwired,
  since it's no longer called from anywhere. Deb pointed out it conflated
  two genuinely distinct Vision One products under ambiguous "standard"/
  "sensor_only" labels. Investigating the split surfaced that the labels
  were actually wrong: Endpoint Inventory's Security Deployment facet for
  Sierra-Cedar only has **"Server & Workload Protection"** and **"Sensor
  only"** categories — there is no separate "Standard Endpoint Protection"
  facet there at all. The old `endpoint.standard` figure (2,468) was almost
  certainly the Server & Workload Protection facet count under a stale/wrong
  label, not a real Standard Endpoint Protection metric.
- **New slide 14: "Standard Endpoint Protection"** (`build_standard_endpoint_protection`).
  Data shape: `D['standard_endpoint_protection'] = {applicable: bool,
  not_applicable_note?, endpoint_status: {managed, at_risk, outdated,
  offline}, threats: [[name,count],...]}`. Same not-applicable-card /
  4-stat-header-row-+-threat-bar-chart pattern as the other product slides.
  **How to check a tenant:** `window.__MENU_FEATURES` group `"Endpoint
  Security"` → entry `/app/epp/endpoint-protection`. If it shows "Protection
  Manager is not set up", not in use. Cross-check against Endpoint
  Inventory's facet list as a second signal (see above). The `applicable:
  true` branch is UNTESTED against real data (Sierra-Cedar doesn't use this
  product) — provisional layout, revisit once a real customer has it.
- **New slide 15: "Server & Workload Protection"** (`build_server_workload_protection`).
  Data shape: `D['server_workload_protection'] = {applicable: bool,
  not_applicable_note?, computer_status: {managed, critical, warning,
  unmanaged}, alert_status: {critical, warning}, latest_alerts: [...],
  infected_computers: [[name,count],...]}`. Layout: 4-stat header row (from
  `computer_status`) + two side-by-side cards — left "ALERT STATUS" (big
  Critical/Warning numbers + optional latest-alerts list), right
  "ANTI-MALWARE STATUS · TOP INFECTED COMPUTERS" (table, or a green-check
  "No infected computers detected" empty state).
  **How to check a tenant:** same `__MENU_FEATURES` group, entry
  `/app/epp/workload-protection`. **This console is hosted cross-origin**
  at `cloudone.trendmicro.com` (Cloud One Workload Security / Deep
  Security) — confirmed by `browser_network_requests` showing a
  `cloudone.trendmicro.com/_workload_iframe/...` request. Cross-origin means
  `browser_evaluate`'s `contentDocument` walk into that iframe silently
  fails (returns nothing, no error) — **read this dashboard via screenshot
  only**, don't waste time trying to sniff its API or DOM-walk it. If it
  loads a live Dashboard tab with widgets (Alert Status, Computer Status
  pie, Anti-Malware Status), it's in use.
  Sierra-Cedar result: in use — 1,893 managed / 220 critical / 330 warning /
  221 unmanaged computers; 0 infected computers ("No Information Available"
  on the page = genuinely clean, not a pull failure).

**2026-07-02 — eighth change, same slide count (19), Alert Status re-sourced:**
- **The "ALERT STATUS" card on the Server & Workload Protection slide no
  longer uses the Deep Security console's own Critical/Warning counts** (Deb
  asked to source it from **Endpoint Inventory → Endpoint Event Viewer →
  SECURITY EVENTS** instead — that cross-origin console's counts and its 5
  identically-truncated "Latest Alerts" strings weren't very useful anyway).
  `alert_status` schema changed from `{critical, warning}` +
  `latest_alerts: [...]` to `{period, categories: [[name,count],...]}` — a
  full replacement of that field's shape, not additive. The card now renders
  as an 8-row striped table (category, count) instead of two big numbers.
  Any count containing `"+"` renders in `REDB` with a footnote ("+ = API
  result cap reached; actual count is higher.") — `build_server_workload_protection`
  checks `"+" in str(count)` per row, so the data file is the single source
  of truth for which rows show as capped.
- **Data source:** Endpoint Inventory → **Endpoint Event Viewer**
  (`#/app/endpoint-event-viewer`) lists 8 categories in its left rail under
  "SECURITY EVENTS": Anti-Malware, Web Reputation, Intrusion Prevention
  (THREAT PREVENTION); Application Control, Device Control, Firewall
  (ACCESS CONTROL); Log Inspection, Integrity Monitoring (ADVANCED
  CAPABILITIES). Same-origin API, one sync-XHR call per category:
  `POST /public/aew/v1/security-events/<slug>/search` (slug = kebab-case
  category name) with body `{searchCriteria:[], period:{start,end},
  eventType:["all"], groups:[], customTags:[], searchTimePerBatch:20,
  eventCountPerBatch:500, maxEventCount:500, nextToken:"", nextLink:""}` —
  count `events.length`.
- **Confirmed this session: async `browser_evaluate` functions are NOT
  awaited by the tool** — an `async () => { return 'hello'; }` came back as
  `{}`, and the same for a `fetch()`-based version. Must use **sync XHR**
  (`xhr.open(..., false)`), matching the pre-existing "must use synchronous
  XHR" rule in [[reference_v1_xdr_portal_extraction]] — this reconfirms it
  applies to `browser_evaluate`'s `function` mode too, not just older
  `browser_interact`-based flows.
- **New gotcha: the 500-event cap on this endpoint is flaky and silent.**
  Many categories return exactly 500 events with `nextToken: null` (no
  signal more data exists) even for a single day's window — the cap isn't
  reliably flagged via `hasNext`. Don't trust an exact-500 result as a true
  count. Mitigation used: re-query in ~5-day chunks across the 30-day window
  and sum; if a chunk comes back under 500, that chunk's count is trustworthy;
  if every chunk hits 500, report the sum with a `"+"` suffix (e.g.
  `"3,000+"`) as an honest lower bound rather than fabricating precision.
- **New gotcha: too many sequential sync-XHR calls in one `browser_evaluate`
  trips the CDP command's own timeout**, independent of the page/API being
  slow. An 8-page while-loop across all 8 categories in a single call timed
  out (`Request timeout: forwardCDPCommand`); reconnecting via `status` and
  splitting into 3-6 sequential calls (each doing at most ~3 pages, or one
  category's worth of chunks) resolved it. Rule of thumb: keep any single
  `browser_evaluate` call's total sync-XHR round-trip time under ~15-20s.
- Sierra-Cedar's 30-day results: Anti-Malware 143, Web Reputation 0, Device
  Control 0 (all exact, well under the cap); Intrusion Prevention 2,628+,
  Application Control 3,000+, Firewall 3,000+, Log Inspection 3,000+,
  Integrity Monitoring 3,000+ (all capped in every ~5-day chunk tried — this
  tenant genuinely generates very high event volume in these categories,
  most likely automated agent/system-level logging rather than real
  security incidents; not called out as anomalous in the deck itself, just
  reported as-is).

**2026-07-02 — ninth change, same slide count (19), Alert Status gains an
Action Taken column + Anti-Malware Status card removed:**
- **`alert_status.categories` is now a list of `[name, count, action]`
  triples**, not 2-tuples — a schema change, not additive (every category
  row must include the 3rd element or unpacking in `build_server_workload_protection`
  raises). The card is now full-width (`px,pw=0.5,12.33`, matching the
  Standard Endpoint Protection slide's single-card-per-row pattern) with 3
  columns (Category / Events / Action Taken) instead of the previous
  narrower 2-column list, since the right half of the slide is now free.
- **Removed the "ANTI-MALWARE STATUS · TOP INFECTED COMPUTERS" card
  entirely** per Deb's request ("Don't need the anti-malware status
  chart") — deleted its rendering code and the `infected_computers` data
  key from both scripts and both data files (generic template + Sierra-
  Cedar's live JSON). It's simply gone, not replaced by anything.
- **How to derive the Action Taken value per category:** the underlying
  event schema's action field name varies — `actionTaken` (array) for some
  categories (e.g. anti-malware), `action` (array) for others (e.g.
  intrusion-prevention, application-control) — check both:
  `e.action || e.actionTaken`. Two categories (Log Inspection, Integrity
  Monitoring) have **no action field populated at all** (`null`/`undefined`
  on every event) — they're audit/change-detection-only categories by
  design (Log Inspection = Windows/Linux security-log entries like
  `WinEvtLog: AUDIT_SUCCESS...`; Integrity Monitoring = file/registry
  change records with `change`/`type`/`previousValue`/`currentValue` fields
  instead of an action). For these, write a descriptive action string like
  "Logged (audit/monitoring, no remediation action)" rather than leaving
  the cell blank or guessing an action that doesn't exist. A quick 7-day
  single-page sample (not the full 30-day chunked count query) is enough to
  tally the dominant action(s) — don't re-run the expensive capped-count
  chunking just to get the action breakdown.
- Sierra-Cedar's action tallies (7-day sample): Anti-Malware → mostly "Log",
  1 "Quarantined" out of 31 sampled (displayed as "Log (occasional
  Quarantine)"); Intrusion Prevention → 100% "Reset"; Application Control →
  100% "Allowed"; Firewall → mixed, "Fail Open: Deny" 371/500 + "Log Only"
  129/500 in the sample (displayed as "Fail Open: Deny (majority), Log
  Only"); Log Inspection / Integrity Monitoring → no action field, displayed
  as "Logged (..., no remediation action)"; Web Reputation / Device Control
  → 0 events, displayed as "—".

**2026-07-02 — tenth addition, deck now 20 slides:**
- **New slide 16: "Cloud Security"**, inserted right after Server &
  Workload Protection (`build_cloud_security`, wired into `content_builders`
  right after `build_server_workload_protection`). Per Deb's explicit
  request, covers all 4 tabs of the **Cloud Security Posture** module in a
  2x2 grid, one panel per tab — the FIRST slide in this deck to summarize an
  always-applicable module (no "Not Applicable" branch; Cloud Security
  Posture is part of Cyber Risk Exposure Management, not an optional
  add-on product like AI Security/Zero Trust/the two Endpoint products).
  - **Data shape:** `D['cloud_security'] = {risk_index, risk_level,
    cloud_overview, entitlements, ai_spm, apis}`, where each of the last 4
    keys is `{headline_value, headline_label, headline_sub, metrics:
    [[label,value],...]}` (exactly 4 metrics rows expected per panel to fit
    the layout, though the renderer doesn't hard-require exactly 4 — it
    just divides available height by `len(metrics)`).
  - **Panel layout (`build_cloud_security`'s inner `panel()` helper):**
    headerbar (title + accent color) → big headline number + label inline
    → a sub-line → up to 4 striped label/value metric rows. Reused nowhere
    else yet; if another slide ever needs a "headline + secondary metrics"
    card, consider factoring `panel()` out, but don't over-abstract for a
    single reuse.
  - **Source module:** Cyber Risk Exposure Management → **Cloud Security
    Posture** (`#/app/server-cloud/cloud-posture/cloud-overview`; also
    listed twice in `window.__MENU_FEATURES` — id 711 under group
    `"Cloud Security"` is the real page, id 7111 under group `"Cyber Risk
    Exposure Management"` is a `cloud-overview-redirect` alias to the same
    place). Its 4 tabs, all consumed:
    - **Cloud Overview:** Cloud Risk Index gauge (hover the trend line for
      the exact score), Account Distribution and Configuration, Protection
      (threat/XDR alerts), Potential Attack Paths, Security Posture
      (vulnerabilities/misconfigurations rings), Compliance (% average +
      per-framework progress bars). No single clean summary API found —
      read via screenshot.
    - **Entitlements** *(marked Preview in the UI)*: Cloud Identity Summary
      (total/human-admin/non-human-admin/overprivileged/unused identity
      counts) + Top Identity Misconfiguration Risk Events table.
    - **AI - Security Posture Management:** a completely separate counter
      set from the main Entitlements tab (AI-scoped cloud accounts/
      services/models/workloads/data storage/entitlements), plus an
      "AI-related cloud assets" ring and mini threat/attack-path/
      vulnerability/misconfiguration panels.
    - **APIs:** a plain paginated table (no tab-level Summary/Details
      toggle like the other three tabs have) — `Total: N` in the footer;
      characterize dominant values (risk score, exposure, type, activity)
      across the visible rows rather than computing an exact distribution.
      Cross-check the total against Cloud Overview's "Assets at Risk → APIs"
      count — same number, useful sanity check.
  - **New DOM-interaction gotchas specific to this module (apply broadly to
    any tenant's Cloud Security Posture page):**
    - **Tab click depth:** the leaf-text-match-then-walk-up-parents click
      technique needed exactly **2** parent levels here (not 3, the depth
      that's worked for category cards elsewhere in this session) — the
      button is closer to the text node on this page. Don't hardcode a
      parent-walk depth across pages; verify per page (dump the parent
      chain and look for the first ancestor with `hasOnclick` truthy, or
      just try 2 then 3).
    - **Scrolling:** this module's content lives in a div with class
      `_content_o8kpz_1` nested inside a same-origin `ui/sase/cp` iframe.
      Setting `document.scrollingElement.scrollTop` (or any window/body
      level scroll) silently no-ops — you must find and scroll that
      specific inner div directly (`el.scrollTop = el.scrollHeight`) to
      reach below-the-fold widgets. The generic recursive
      "walk all iframes looking for scrollHeight > clientHeight" helper
      used earlier in this session (for Data Security's page, etc.) still
      finds it — just at a deeper nesting level than usual.
  - Sierra-Cedar result: 93 cloud accounts (91 AWS/1 Azure/1 Oracle), Risk
    Index 72.6 (High), 620 threat alerts + 194 attack paths (30 days, all in
    AWS), compliance 44% average; 19,011 total identities (8,129 unused, 0
    overprivileged); 18 AI-related cloud assets (0 high-risk); 108 API
    collections (101 Medium risk, dominant profile Public/REST/Inactive).

**2026-07-02 — eleventh addition, deck now 21 slides:**
- **New slide 17: "Network Security"**, inserted right after Cloud Security
  (`build_network_security`, wired into `content_builders` right after
  `build_cloud_security`). Code is a near-exact copy of `build_zero_trust`/
  `build_ai_security`'s structure (not-applicable card + per-module
  status-card grid for the applicable branch) — same family of "is this
  optional module in use?" slides. Data shape:
  `D['network_security'] = {applicable: bool, not_applicable_note?,
  modules?}`, default `applicable: false`.
  - **How to check a tenant:** `window.__MENU_FEATURES` group `"Network
    Security"` → entry points **Network Overview**
    `/app/network-security/network-overview` and **Network Inventory**
    `/app/network-security/ni`. If either shows a "Choose a deployment
    option to continue" onboarding wizard (three appliance-choice cards —
    Virtual Network Sensor, Deep Discovery Inspector, TippingPoint — each
    just a description + "Continue with ..." button, plus a "Skip this step
    for now" link), the module is not in use. Checking both entry points is
    a cheap cross-check (both agreed for Sierra-Cedar).
  - Sierra-Cedar result: neither page showed a live dashboard → `applicable:
    false`.
  - The `applicable: true` branch is UNTESTED against real data (no
    customer in this session has a deployed network appliance) — same
    caveat as the other provisional-branch slides (AI Security, Zero
    Trust, Standard Endpoint Protection); revisit the per-appliance
    metrics shape once a real tenant has one deployed.
- `content_builders` is now `[build_cyber] + risk_detail_builders +
  [build_mdr, build_data_source, build_intel, build_workflow_automation,
  build_ai_security, build_zero_trust, build_identity, build_data_security,
  build_standard_endpoint_protection, build_server_workload_protection,
  build_cloud_security, build_network_security, build_email, build_credits,
  build_eoy_prediction]`. Current slide order (21 total): Title(1), Cyber(2),
  Risk Factors Detail(3-5), MDR(6), Data Source and Log Management(7),
  Intelligence Reports(8), Workflow and Automation(9), AI Security(10),
  Zero Trust(11), Identity(12), Data Security(13), Standard Endpoint
  Protection(14), Server & Workload Protection(15), Cloud Security(16),
  Network Security(17), Email(18), Credits(19), EOY(20), Team(21). If asked
  to reorder again, just move the relevant builder function(s) within that
  list — no other code changes needed, since `expected` is computed from
  `len(content_builders)` fresh each run.
- `content_builders` is now `[build_cyber] + risk_detail_builders +
  [build_mdr, build_data_source, build_intel, build_workflow_automation,
  build_ai_security, build_zero_trust, build_identity, build_data_security,
  build_standard_endpoint_protection, build_server_workload_protection,
  build_email, build_credits, build_eoy_prediction]`. Current slide order
  (19 total): Title(1), Cyber(2), Risk Factors Detail(3-5), MDR(6), Data
  Source and Log Management(7), Intelligence Reports(8), Workflow and
  Automation(9), AI Security(10), Zero Trust(11), Identity(12), Data
  Security(13), Standard Endpoint Protection(14), Server & Workload
  Protection(15), Email(16), Credits(17), EOY(18), Team(19). If asked to
  reorder again, just move the relevant builder function(s) within that
  list — no other code changes needed, since `expected` is computed from
  `len(content_builders)` fresh each run.

**2026-07-02 — twelfth change, same slide count (21), Email data source
switched + gains a Not Applicable branch:**
- **`build_email` now checks `d.get('applicable', False)`** (default
  False, matching AI Security/Zero Trust/Network Security's pattern) — a
  new not-applicable card renders instead of the stat row + charts when
  `applicable` is false or missing. `email.not_applicable_note` follows the
  same convention as the other optional-module slides.
- **Data source changed** from "Cloud Email and Collaboration Protection →
  Dashboard" (`/app/email-and-collaboration/dashboard`) to **"Email and
  Collaboration Security → Configuration and Operations → Overview"**
  (`#/app/email/overview`, `window.__MENU_FEATURES` id 8729) per Deb's
  explicit request. This resolves a long-standing unresolved issue from
  earlier in the session: the old Dashboard path's iframe would load but
  never become the visually active view no matter what click technique was
  tried, across several separate attempts. The new Overview page loads
  reliably.
  - **Overview page layout:** a fixed ~7-day date range (not adjustable via
    a visible date picker) + 4 widgets: Top Users with Account Takeover
    Risks, Top Users with Targeted Attack Risks, Scanning Breakdown,
    Threats Detection Count. Each shows either real data or a plain "No
    data to display." + Reload button (not an error state).
  - **How to determine Not Applicable:** if all 4 widgets show "No data to
    display.", set `applicable = false` — this is a legitimate "nothing to
    report" state, not a pull failure, so don't leave placeholder zeros
    (they'd misleadingly read as "scanned 0 messages" rather than "no data
    available"). Sierra-Cedar result: all 4 widgets empty → `applicable:
    false`.
  - The `applicable: true` branch (stat row + Threat Detections chart + Top
    5 High-Risk Recipients) is the SAME rendering code from before this
    change — untouched, just now gated behind the applicability check.
    Still untested against real data from this specific Overview page
    (the field names `scanned`/`threats_total`/`threats`/`recipients` were
    designed against the older Dashboard page's shape) — when a customer
    actually has data here, verify the Overview page's widgets map cleanly
    onto those same fields, or adjust if the new page's data shape differs.

**2026-07-02 — thirteenth change, same slide count (21), Identity's Risky
Accounts column gains a graceful empty state:**
- **`build_identity`'s "RISKY ACCOUNTS · TOP 5" column now checks
  `d.get('risky_accounts')`** — if empty, it renders `identity.risky_accounts_note`
  centered in the column instead of trying to render a fake ranked row.
  Previously the data file carried a literal placeholder row
  (`["1", "[ not captured this run ]", "[ not captured this run ]", "N/A"]`)
  that rendered verbatim on the slide — Deb spotted this looking like a bug
  ("What is the problem with the Identity Security Posture slide?") and
  asked for a fix.
  - **Rule going forward: never leave bracketed placeholder text
    (`"[ ... ]"`) in a customer-facing data file** — any field that can't be
    filled in should use an empty value (`[]`, `""`, `0`) paired with a
    `*_note` explanation field that the renderer displays gracefully, not a
    literal "not captured" string that ends up on the slide itself. This is
    the same pattern already used for `not_applicable_note`,
    `playbooks_note`, `integrations_note`, etc. — apply it uniformly to any
    NEW optional/nullable field added to this deck in the future.
  - **Where I looked for real risky-account data before giving up (document
    so a future run doesn't have to re-discover this):** Identity Inventory
    (`#/app/identity-inventory`) is a plain directory (User/Enterprise
    Application/Device/Group/Role/Access Policy/Granted Permission tabs) —
    no risk-score column anywhere, confirmed via
    `doc.querySelectorAll('th')` header dump. Clicking a Priority Risk
    Event's name (Identity Security Posture Overview tab) or an Exposure
    Event row's small icon (Exposure tab) both just launch the same generic
    Identity Inventory app in a new iframe (`ui/sjg/idInventory/...`) rather
    than a scored/filtered account list — confirmed via
    `browser_network_requests` showing the same `idInventory` API calls
    either way. **New discovery:** Identity Inventory's Overview tab shows a
    "Data sync: Unhealthy" / "Policy enforcement: Disconnected" status for
    Microsoft Entra ID specifically (separate from the already-known
    on-premises AD connector retirement) — likely contributes to why no
    risk-scored account view is surfaced for Sierra-Cedar right now.

**2026-07-02 — fourteenth change, first InsidePlastics run, several new generalizable findings:**
- **Risky Accounts: check "Accounts leaked on dark web" on Identity Security
  Posture's Overview tab before concluding no ranking exists.** For
  InsidePlastics this panel had named users ranked by asset criticality
  (0-10 scale) — a real usable ranking, unlike Sierra-Cedar's dead end. Add
  this to the standard search checklist for the Risky Accounts column.
- **A CREM Data Sources page "Configured" label doesn't guarantee the
  module's own app shows live data — always verify the app itself.**
  InsidePlastics' Data Sources page listed "Zero Trust Secure Access -
  Private Access" as Configured, but the Zero Trust app itself never got
  past the "Start paid usage" onboarding splash (mock screenshot data). Set
  `applicable` based on what the module's own dashboard shows, not the
  connector-status label.
- **Conversely, Network Overview's onboarding wizard doesn't mean Network
  Security is unused — check Network Inventory too.** InsidePlastics'
  Network Overview page showed the same "Choose a deployment option"
  wizard Sierra-Cedar had, but Network Inventory (`#/app/network-security/ni`)
  showed a live, healthy Deep Discovery Inspector appliance. Cross-checked
  against the MDR PDF's "Monitoring sources" table (DDI supplied 91.4% of
  MDR's event volume) to confirm it's genuinely active. **Always check both
  pages before concluding Network Security is not applicable.**
- **Data Security and Cloud Security have no `applicable`/not-applicable
  branch in the build script** — they always render their full layout. When
  a tenant genuinely has no cloud accounts or no sensitive-data scanning
  configured, fill in real zeros/"N/A" strings (not bracketed placeholders)
  rather than trying to force a not-applicable state that the renderer
  doesn't support. `top_assets: []` and empty `cloud_overview.metrics`
  values render gracefully already.
- **AI Security Blueprint can show a genuinely live (not onboarding-splash)
  dashboard that still has zero real usage** — distinguish "live dashboard,
  all zeros" from "onboarding splash with mock screenshots" when writing the
  `not_applicable_note`; both count as not-in-active-use but the live-zeros
  case is worth describing precisely (incl. any "no permission" sub-panels)
  since it's a different failure mode than a pure sales splash.
- **`build_email`'s "THREAT DETECTIONS · LAST 30 DAYS" header text is
  hardcoded in the script**, not derived from data — if the actual pulled
  window is shorter (InsidePlastics' was 7 days, matching the Email
  Overview page's fixed ~7-day window), the slide header will say "30 days"
  regardless. Not fixed this run; flag as a known script inaccuracy if a
  customer's actual window is ever shorter than 30 days.
- **Credits `purchased`/`used`: the raw purchases API sums across ALL
  historical renewal terms (lifetime total), which is useless for a
  single-term slide.** Better approach: back-calculate `purchased` from the
  dashboard's own `balance` and `balance_pct` (`purchased = balance /
  balance_pct`) — this is guaranteed internally consistent with what's
  displayed, rather than trying to filter/sum the purchases list by status
  and date, which can disagree with the dashboard by a wide margin depending
  on which purchases you include (confirmed a ~180k discrepancy between the
  two methods on this tenant).
- **Endpoint Event Viewer: `nextToken: null` is the authoritative
  "search exhausted, this count is exact" signal** — don't assume a count
  needs a `"+"` suffix just because a single response hit the 500-per-call
  cap; keep following the chained `nextToken` until a response returns
  `null`, then trust the summed total as exact (InsidePlastics needed this
  distinction: several categories summed to exactly 1,000 across 2 chained
  calls with a final `nextToken: null`, meaning the true count was exactly
  1,000, not "1,000+").
