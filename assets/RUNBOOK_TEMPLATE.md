# {{CUSTOMER}} Value Report — How to Run

**Trigger:** say **"Run the {{CUSTOMER}} value report"**, or invoke the generic
`/value-report` skill and answer "{{CUSTOMER}}" when asked which customer.

**Output:** `~/Desktop/{{CUSTOMER}} Value Report.pptx` — a 21-slide TrendAI
Vision One™ value-report deck (dark theme), {{CUSTOMER}} tenant.

**Golden rule:** always **re-pull live data current as of the date asked** —
never just rebuild from the old JSON. (Set `as_of` / `month_year` to today.)

---

## Slides
1. **Title** — presenter name (hand-filled); date auto-updates.
2. **Cyber Risk Exposure Management** — CRI + 2x3 grid: Devices, Internet-Facing Assets, Accounts, Applications, Cloud Assets, Vulnerabilities (headline count + risk chip each).
3–5. **Risk Factors Detail** (2 categories per slide, generated dynamically from `risk_detail` — 3 slides for 6 categories) — the specific risk factors/signals driving each category's risk level. A category with no factors renders a "no risk factors / no data source" note instead of an empty table.
6. **MDR** — monthly Managed Detection & Response summary (from the downloaded PDF).
7. **Data Source and Log Management** — 4 headline metrics (New Analytic Ingestion, New Archival Ingestion, Total Extended Analytic Retention, Total Extended Archival Retention) + a full "Ingestion and Retention by Data Source" table.
8. **Intelligence Reports** — matched sweeps (last 30 days).
9. **Workflow and Automation** — Security Playbooks list (or a "none configured" note) + Third-Party Integrations list (Cyber Risk Exposure Management → Data Sources, status = Configured only).
10. **AI Security** — status of AI Security Blueprint / AI Application Security / AI Secure Access. Renders "Not Applicable" if the customer isn't using any of them (the common case today), or per-module status/metrics cards if they are.
11. **Zero Trust** — status of Zero Trust Secure Access (AI Secure Access, Internet Access, Private Access). Renders "Not Applicable" if the customer isn't using it (the common case today), or per-module status/metrics cards if they are — same pattern as AI Security.
12. **Identity Security Posture** — 4 equal columns: Risk Events · Exposure Events · Attack Events · Risky Accounts.
13. **Data Security** — headline stats (total assets, assets with sensitive data, monitored assets) + Sensitive Data Detections by Category (bar chart) + Top Risky Assets with Sensitive Data (table).
14. **Standard Endpoint Protection** — status of the Apex Central-managed endpoint protection product. Renders "Not Applicable" if the customer isn't using it, or headline stats + a threat-detections chart if they are.
15. **Server & Workload Protection** — status of the Cloud One Workload Security product. Renders "Not Applicable" if the customer isn't using it, or Computer Status stats + a full-width Alert Status table (Security Events by category, with an Action Taken column, sourced from Endpoint Event Viewer) if they are.
16. **Cloud Security** — 2x2 grid covering Cloud Security Posture's 4 tabs: Cloud Overview, Entitlements, AI - Security Posture Management, APIs — each panel a headline stat + 4 secondary metrics.
17. **Network Security** — status of network appliance deployment (Virtual Network Sensor / Deep Discovery Inspector / TippingPoint). Renders "Not Applicable" if none is deployed (the common case today), or per-appliance status/metrics cards if configured — same pattern as AI Security/Zero Trust.
18. **Cloud Email & Collaboration** — sourced from Email and Collaboration Security → Overview. Renders "Not Applicable" if that page shows no data (the common case today), or scanned/threats breakdown if it does.
19. **Credits** — purchased / used / balance, monthly drawdown.
20. **End of Year Prediction Summary** — purchased vs predicted annual usage (avg monthly drawdown × 12).
21. **TrendAI Account Team** — contacts (hand-filled).

**Note on slide count:** slides 3–5 (Risk Factors Detail) are generated dynamically — 2 categories per slide, from the `risk_detail` list in `value_report_data.json`. Total slide count depends on `len(risk_detail)`; if it changes between runs, the rebuild falls back to a full rebuild from template (preserve-mode requires a matching slide count) — hand-edited presenter/team-contact fields would need re-entering after such a change.

---

## Files (`~/Documents/{{CUSTOMER}}/value_report/`)
| File | Purpose |
|------|---------|
| `build_value_report.py` | Builds the deck (generic — copied verbatim from `~/Documents/Claude/value_report/`). Run: `cd "~/Documents/{{CUSTOMER}}/value_report" && python3 build_value_report.py` |
| `value_report_data.json` | The only thing that changes per run — all figures live here, including `customer` and `presenter`. |
| `template.pptx` | Base template (title slide + "Quote 01" layout). |
| `value_report_data.json.bak_<date>` | Auto-backup of the prior run's data. |

---

## Data pull (live, via Blueprint MCP on `portal.xdr.trendmicro.com`, {{CUSTOMER}} tenant)

Switch the org to **{{CUSTOMER}}** via the top-right org switcher first.

Best method for the dashboard numbers: **read the dashboard's own JSON APIs**
with `browser_network_requests` (reload page → `list urlPattern='/public/'` →
`details` on the IDs below) rather than scraping the DOM.

- **Cyber (slide 2)** — CRI + one panel per category (Devices, Internet-Facing
  Assets, Accounts, Applications, Cloud Assets, Vulnerabilities), each with a
  headline count + risk-level chip:
  - CRI: read the Cyber Risk Overview score + risk level.
  - Devices: `GET /public/ass/api/v1/trilogy/riskOverview/visibilityAttackSurface?type=device` → discovered = `total`, assessable/managed = `full`.
  - Internet-Facing Assets: `GET /public/ass/api/v1/trilogy/riskOverview/internet/riskLevel` → `ipRisk.total`, `domainRisk.total` (+ low/medium/high breakdown).
  - Accounts: `GET /public/ass/api/v1/trilogy/riskOverview/userAccounts/riskLevelBar` → `domainAccounts.total`, `serviceAccounts.total`.
  - Applications: click the "Applications" card first (see tab-click note below), then `GET /public/ass/api/v1/trilogy/riskOverview/appAssets/riskLevelBar` → `appAssets.total`/low/medium/high. If the card shows "No risk factors detected — configure data sources", there's genuinely no app inventory connected; don't fabricate a count.
  - Cloud Assets: click the "Cloud Assets" card, then `GET /public/sase/api/v1/bigtable/forward/v1/cloud-assets/asset-level-bar` → `cloudAssets.total`/low/medium/high.
  - CVEs (Vulnerabilities panel): `GET /public/ass/api/v1/risks/factorStatistics?period=30&scoringAlgorithm=v2` → factor "Vulnerability detection": `count`=High, `mediumCount`, `lowCount`.
  - MTTP / Avg Unpatched Time / regional average: gauges on the Risk Overview — not always reachable via API; note as "not captured" rather than blocking on it.
- **Risk Factors Detail (slides 3–5)** — 2 categories per slide, generated from
  `risk_detail` (a flat list of `{title, risk_level, factors: [[count,label],...], note?}`).
  These are the specific signals behind each category's risk level, pulled
  per-category from the SAME dashboard (click each card, then read its risk
  factor endpoints):
  - Devices: `GET /public/ass/api/v1/trilogy/riskOverview/riskFactor?assetType=device` (`cves`, `cyberThreats`, `topIndustrialThreats`, `zeroDayAlerts`) + `GET /public/cta/api/v1/trilogy/riskOverview/riskFactors?assetType=endpoint` (`endpoint.cyberThreatCnt` — often more accurate than the `ass` endpoint's own `cyberThreats` field for the "cyber threats" figure shown on the card).
  - Internet-Facing Assets: `riskFactor?assetType=internet` → `insecureDomains`, `unexpectedServices`, `cves`.
  - Accounts: `riskFactor?assetType=account` → `compromiseEvents`; richer breakdown from `GET /public/cta/api/v1/trilogy/riskOverview/riskFactors?assetType=account` → `weakSignInCnt`, `excessiveCnt`, `legacyProtocolCnt`, `excessivePrivilegeCnt`, `cyberThreatCnt`.
  - Applications: no working `assetType` value found when no data source is connected — leave `factors: []` with a `note` explaining why (matches the card's own "no risk factors detected" message).
  - Cloud Assets: `riskFactor?assetType=cloudAsset` → `complianceViolations`, `sentryCves` (cloud VMs), `functionCves` (serverless), `imageCves` (container images), `highRiskConfigurationRisks`, `cves` (container clusters).
  - Vulnerabilities: reuse the severity breakdown from the Cyber slide's CVE factor (High/Medium/Low counts) — this slide's "factors" are just those three rows.
  - **If `len(risk_detail)` differs from a prior run, the deck falls back to a full rebuild** (slide count changed) — hand-edited presenter/team fields need re-entering after that.
- **MDR (slide 6):** **ASK the user first, every time** — "Does {{CUSTOMER}}
  have MDR (Managed Detection and Response)?" Never infer this from the
  dashboard and never reuse a prior run's answer.
  - **If yes:** the user downloads the latest "Managed Detection and Response
    (MDR) – Monthly Report" PDF from Reports → Generated reports (downloads
    are blocked in automation) — wait for them to confirm it's downloaded,
    then read it with `pypdf` and fill in `mdr`.
  - **If no:** set `mdr.applicable = false` (and optionally
    `mdr.not_applicable_note`) in `value_report_data.json`. `build_mdr`
    renders a clean "not part of this subscription" message instead of an
    empty funnel — don't leave placeholder zeros, they misleadingly read as
    "MDR ran and found nothing" rather than "not applicable."
- **Data Source and Log Management (slide 7):** Agentic SIEM and XDR →
  Data Source and Log Management → **Data Monitoring** tab (not "Data sources
  and retention" — that's a different tab on the same page, just a retention-
  period config list, not usage figures). Same-origin APIs:
  - Headline metrics: `GET /ui/dsr/uss/api/v1/data_usage/overview` →
    `newAnalyticIngestionGB`, `newComplianceIngestionGB` (= archival ingestion),
    `totalChargedAnalyticRetentionGB`, `totalChargedComplianceRetentionGB`
    (= archival retention). This is a **daily snapshot** (today's date), not a
    range — labelled accordingly in `data_source.period`.
  - Per-source table: `POST /ui/dsr/uss/api/v1/data_usage/data_usage_by_source`
    (body: `{ingestionType:[], startDate:<epoch ms>, endDate:<epoch ms>, dataSources:[...]}` —
    the page sends this itself on load with all configured source codes; just
    read the response, don't hand-construct the body) → array of
    `{sourceName, category, analyticIngestionGB, complianceIngestionGB,
    freeAnalyticRetentionGB, chargedAnalyticRetentionGB,
    freeComplianceRetentionGB, chargedComplianceRetentionGB}`. Build each
    `sources` row as `[name, category[0], ingestion=analytic+compliance,
    retention=sum of all four retention fields]`. Don't filter out all-zero
    rows — a source with zero ingestion but nonzero retention (e.g. a risk-
    event feed) is still meaningful.
- **Intel (slide 8):** Intelligence Reports → set View = "Matched sweeps only"
  → count those in the last 30 days. If zero, `build_intel` already renders a
  positive-signal panel — just fill in `intel.note` with the specifics.
- **Workflow and Automation (slide 9):**
  - **Security Playbooks:** Workflow and Automation → **Playbooks** tab
    (not "Execution Results" — that's just a history of past runs and may say
    "No execution results" even when playbooks exist). "No playbooks created"
    means genuinely none configured — leave `workflow_automation.playbooks: []`
    with a `playbooks_note`, don't leave it looking like a data-pull failure.
    **Sidebar navigation gotcha:** the "Workflow and Automation" left-rail
    group's submenu popup may never mount in the DOM no matter how you
    click/hover/double-click it (same failure mode as some other sidebar
    groups). Reliable workaround: read `window.__MENU_FEATURES` (an
    undocumented global array of `{id, title, path, group, description}` for
    every menu item) via `browser_evaluate`, find the entry with
    `title === "Security Playbooks"`, and navigate directly to
    `https://portal.xdr.trendmicro.com/#` + its `path` (e.g.
    `/app/workflow-automation-eco`) instead of clicking through the UI.
  - **Third-Party Integrations:** Cyber Risk Exposure Management → click the
    "Data sources" button (top of the dashboard) → this navigates to
    `#/app/dsr/asrm`. Read `GET /ui/dsr/uss/api/v2/data_source/data_sources`
    and filter to `status == 1` (Configured) — take the `sourceName` field of
    each. Per Deb's instruction, list **all** Configured sources here (don't
    try to separate "genuinely third-party" from Trend's own products — the
    literal ask is the Configured list from this page).
- **AI Security (slide 10):** check all three sub-modules —
  `window.__MENU_FEATURES.filter(x => x.group === "AI Security")` gives the
  exact `path` for each (as of this writing: **AI Security Blueprint**
  `/app/ai-security-blueprint`, **AI Application Security** `/app/ai-app-sec`,
  **AI Secure Access** `/app/zero/overview2`, which redirects to the Zero
  Trust Secure Access overview since it's the same underlying module).
  Navigate to each directly by URL. **If any/all show a "Start paid usage" /
  "Start free trial" / "Opt in" onboarding landing page** (with blurred
  sample screenshots, not live numbers), that module is NOT in use — set
  `ai_security.applicable = false` and let the default not-applicable note
  render (or write a custom `not_applicable_note` naming which modules you
  checked). Don't mistake the onboarding page's mock screenshot numbers for
  real data. Only set `applicable = true` (and fill in `modules`, one entry
  per module with `name`/`status`/`metrics`) if a module actually opens into
  a live dashboard with real figures.
- **Zero Trust (slide 11):** check the module the same way as AI Security —
  `window.__MENU_FEATURES.filter(x => x.group === "Zero Trust Secure Access")`
  gives every sub-page's path; the entry point is **Secure Access Overview**
  `/app/zero/overview`. **If it shows a "Start paid usage" / "Start free
  trial" onboarding landing page** (mock screenshots, not live numbers —
  same visual pattern as AI Security's onboarding pages, and in fact AI
  Secure Access's own path `/app/zero/overview2` redirects to this exact
  page, since AI Secure Access is a feature within Zero Trust Secure Access
  rather than a separate dashboard), set `zero_trust.applicable = false` and
  let the default not-applicable note render. Only set `applicable = true`
  (and fill in `modules`) if it opens into a live dashboard with real risk
  control / access figures instead.
- **Identity (slide 12):** Identity Security Posture → Overview (risk level,
  risk-event table), Exposure tab (exposure events), Attack tab
  (`#/app/identity-posture/attack` → Attack Events). Risky accounts = top execs
  by asset risk score. **If you can't find a risky-accounts ranking for this
  tenant** (checked so far: Identity Inventory `#/app/identity-inventory` has
  no risk-score column, just a plain directory; clicking through risk-event
  detail links/icons on the Exposure tab just re-opens the same Identity
  Inventory app rather than a scored view) — set `identity.risky_accounts = []`
  and fill in `identity.risky_accounts_note` with a one-line explanation of
  what you checked. `build_identity` renders that note centered in the
  column instead of the ranked list. **Never leave bracketed placeholder
  text like `"[ not captured this run ]"` in the data file** — it renders
  literally on the slide and reads as a bug, not as a null-state.
- **Data Security (slide 13):** Cyber Risk Exposure Management → **Data
  Security Posture** (`#/app/data-posture`, also reachable via
  `window.__MENU_FEATURES` group `"Data Security"`). Same-origin APIs (no
  click-through needed — they fire on page load):
  - Headline stats + category bar chart: `GET /public/sase/api/v1/bigtable/forward/v1/sensitive-assets/summary?storageType=all`
    → `statistics` (`totalAssets`, `sensitiveAssets`, `monitoredAssets`) and
    `distributions` (`sensitiveType`: personal/financial/credentials/other/custom
    → `count`). Map `sensitiveType` to the display labels Personal/Financial/
    Credentials/Other/Custom.
  - Top risky assets table: `GET /public/sase/api/v1/bigtable/forward/v1/sensitive-assets/list?storageType=all&limit=5`
    → `assets[]`, each `{assetName, assetType, sensitiveTypes[], latestRiskScore}`
    — already sorted by risk score descending; join `sensitiveTypes` with
    ", " for the table's "Sensitive Data Type" column.
  - Overall score/risk level for the subtitle: same score gauge shown at the
    top of the Data Security Posture page (0-100 with a Low/Medium/High
    band) — read via screenshot if no dedicated summary endpoint is found.
- **Standard Endpoint Protection (slide 14) and Server & Workload Protection
  (slide 15):** these are two DISTINCT products, both reachable via
  `window.__MENU_FEATURES` group `"Endpoint Security"` — **Standard Endpoint
  Protection** `/app/epp/endpoint-protection` (Apex Central-based centralized
  management) and **Server & Workload Protection** `/app/epp/workload-protection`
  (Cloud One Workload Security / Deep Security). Check each independently —
  a tenant can have neither, either, or both.
  - **Standard Endpoint Protection:** if the page shows "Protection Manager
    is not set up", it's NOT in use — set
    `standard_endpoint_protection.applicable = false`. Corroborate against
    **Endpoint Inventory** (`#/app/ves-inventory`) left-rail "Security
    Deployment" facet: if there's no "Standard Endpoint Protection" facet
    value listed (only "Server & Workload Protection" / "Sensor only"
    appear), that's a second confirming signal. If genuinely in use, pull
    managed/at-risk/outdated/offline counts and a threat-detection breakdown
    from the Protection Manager dashboard.
  - **Server & Workload Protection:** this console is hosted **cross-origin**
    (`cloudone.trendmicro.com`, a separate Cloud One / Deep Security domain)
    — `browser_evaluate` can't reach into its DOM/iframes (a
    `contentDocument` access on a cross-origin frame silently fails), so read
    figures via **screenshot**, not API sniffing. If it loads a live
    Dashboard tab with widgets (Alert Status, Computer Status pie, Anti-
    Malware Status, etc.), it's in use — set `applicable = true` and read the
    **Computer Status** widget → `computer_status` (`managed`, `critical`,
    `warning`, `unmanaged`). The page can take several seconds to move past
    a loading spinner — wait before screenshotting. If the tenant has never
    provisioned this product, the page will not resolve to a live dashboard
    (confirm the entry point doesn't redirect back to a Cloud Security
    landing/marketing page) — set `applicable = false` in that case.
    **Neither Alert Status nor Anti-Malware Status/Top-Infected-Computers
    are sourced from this console** — the slide no longer shows an
    Anti-Malware Status card at all (removed per Deb's request; wasn't
    needed once Alert Status carries the Action Taken detail), and Alert
    Status comes from Endpoint Event Viewer instead (see below).
  - **Alert Status table** (full-width card on the Server & Workload
    Protection slide, 3 columns: Category, Events, Action Taken) is pulled
    from **Endpoint Inventory → Endpoint Event Viewer → SECURITY EVENTS**
    (`#/app/endpoint-event-viewer`), which lists 8 categories in its left
    rail: Anti-Malware, Web Reputation, Intrusion Prevention (under THREAT
    PREVENTION), Application Control, Device Control, Firewall (under
    ACCESS CONTROL), Log Inspection, Integrity Monitoring (under ADVANCED
    CAPABILITIES). `alert_status.categories` is a list of
    `[name, count, action]` triples (NOT a 2-tuple — the "action" column was
    added per Deb's request; the Anti-Malware Status / Top Infected
    Computers card that used to sit next to this table was removed the same
    day since it wasn't needed alongside the action breakdown).
    Same-origin API, one call per category:
    `POST /public/aew/v1/security-events/<slug>/search` (slugs are the
    kebab-case category names, e.g. `anti-malware`, `intrusion-prevention`)
    with body `{searchCriteria:[], period:{start,end}, eventType:["all"],
    groups:[], customTags:[], searchTimePerBatch:20, eventCountPerBatch:500,
    maxEventCount:500, nextToken:"", nextLink:""}`. Use **sync XHR**
    (`uic-token` header), not `fetch`/async (this tool doesn't await async
    `browser_evaluate` functions — confirmed by testing, an async function
    returning a string comes back as `{}`).
    - **Count:** `events.length`, subject to the same 500-event-cap
      gotcha documented below.
    - **Action Taken:** the field name varies by category schema — most use
      `actionTaken` (an array, e.g. anti-malware), some use `action` instead
      (an array for intrusion-prevention/application-control, but `null` for
      log-inspection/integrity-monitoring — those two categories are audit/
      change-detection only and never populate an action field at all).
      Tally `e.action || e.actionTaken` across the sampled events returned
      and pick the dominant value(s); if a category's action field is always
      null/absent, describe it descriptively instead (e.g. "Logged
      (audit/monitoring, no remediation action)") rather than leaving it
      blank. A 7-day sample (single page, fast) is usually enough to see the
      dominant action(s) — you don't need to re-derive this from the full
      30-day/chunked count query.
    **Gotcha — the 500-event cap is flaky and easy to hit:** many categories
    silently cap at exactly 500 with `nextToken: null` (i.e. no signal that
    more data exists) even over just a single day. If a category returns
    exactly 500, don't trust it as exact — re-query in a handful of ~5-day
    chunks across the 30-day window and sum; if EVERY chunk still hits 500,
    report the sum with a `"+"` suffix (e.g. `"3,000+"`) rather than
    presenting a false-precision number — `build_server_workload_protection`
    already renders any count containing `"+"` in red with a footnote
    explaining the cap. Don't chunk finer than ~5 days per call or issue too
    many sequential sync-XHR calls in one `browser_evaluate` — both make the
    call block long enough to trip the CDP command's own timeout (confirmed:
    an 8-page loop across 8 categories in one call timed out; 3-6 sequential
    calls of a few pages each is safe). Sierra-Cedar result (30 days):
    Anti-Malware 143 (Log, occasional Quarantine), Web Reputation 0 (—),
    Device Control 0 (—) — all exact, well under the cap; Intrusion
    Prevention 2,628+ (Reset), Application Control 3,000+ (Allowed),
    Firewall 3,000+ (Fail Open: Deny majority, Log Only), Log Inspection
    3,000+ (Logged, audit-only), Integrity Monitoring 3,000+ (Logged,
    change-detection-only) — all capped in every chunk tried.
- **Cloud Security (slide 16):** Cyber Risk Exposure Management → **Cloud
  Security Posture** (`#/app/server-cloud/cloud-posture/cloud-overview`,
  also reachable via `window.__MENU_FEATURES` — the "Cloud Security Posture"
  entry under group `"Cloud Security"`, id 711 — id 7111 under group
  `"Cyber Risk Exposure Management"` is a redirect-only duplicate). This
  module has 4 tabs, all needed for this slide: **Cloud Overview**,
  **Entitlements** *(Preview)*, **AI - Security Posture Management**, and
  **APIs**. Switching tabs by click works — find the tab's leaf text node,
  walk up exactly **2** parents (not 3, unlike the category-card pattern
  elsewhere) to reach the clickable element, and `.click()` it directly (the
  tab bar and its content live inside a same-origin iframe, `ui/sase/cp`).
  Content is inside a scrollable div with class `_content_o8kpz_1` (not the
  outer document) — set `.scrollTop` on that element directly to see
  below-the-fold widgets; window/body-level scroll calls do nothing.
  - **Cloud Overview:** read (via screenshot; no single clean summary
    endpoint found) — Cloud Risk Index gauge (hover the trend line's last
    point for the exact score) + risk level; Account Distribution and
    Configuration card (per-provider account counts); Protection card
    (Threat alerts, XDR alerts, both with a "Last 24 hrs" delta); Potential
    Attack Paths card; Compliance card (% average + per-framework bars, each
    with a `matched/total` pair). `cloud_overview.metrics` in the data file
    only needs 4 rows — pick the most report-worthy (threat alerts, attack
    paths, high-risk asset count, compliance average); the rest is
    available if a customer wants a deeper cut later.
  - **Entitlements (Preview tab):** Cloud Identity Summary card (total
    identities, human/non-human admins, overprivileged, unused) + Top
    Identity Misconfiguration Risk Events table (risk event name + impacted
    count) — the top row is usually "Unused AWS IAM Role" for AWS-heavy
    tenants.
  - **AI - Security Posture Management:** left-rail counts (Cloud accounts,
    Services, Models, Workloads, Data Storage, Entitlements — all
    AI-specific, distinct from the main Entitlements tab's numbers) + the
    big "AI-related cloud assets" ring (total + high-risk + 24hr delta) +
    threat detections / potential attack paths / vulnerabilities /
    misconfiguration risk events mini-panels.
  - **APIs:** a plain table (API collection name, asset risk score,
    endpoint counts, exposure, API type, activity status), `Total: N` in
    the pagination footer. No per-tenant aggregate/summary view was found on
    this tab — characterize the dominant values across the visible rows
    (API type, exposure, activity, typical risk score) rather than trying to
    compute a precise distribution; cross-check the total against the
    "APIs" row in Cloud Overview's "Assets at Risk" grid (same number).
- **Network Security (slide 17):** check via `window.__MENU_FEATURES` group
  `"Network Security"` — entry points **Network Overview**
  `/app/network-security/network-overview` and **Network Inventory**
  `/app/network-security/ni`. **If either shows a "Choose a deployment
  option to continue" onboarding wizard** (three cards: Virtual Network
  Sensor, Deep Discovery Inspector, TippingPoint — no appliance data, just
  "Continue with ..." buttons and a "Skip this step for now" link), the
  module is NOT in use — set `network_security.applicable = false` and let
  the default not-applicable note render. Checking both pages is a good
  cross-check (both showed the same wizard for Sierra-Cedar). Only set
  `applicable = true` (and fill in `modules`, one entry per deployed
  appliance type with `name`/`status`/`metrics`) if either page opens into
  a live dashboard with real network traffic/appliance figures instead.
- **Email (slide 18):** source is **Email and Collaboration Security →
  Configuration and Operations → Overview** (`#/app/email/overview`, also
  reachable via `window.__MENU_FEATURES` — group `"Email and Collaboration
  Security"`, id 8729). This page loaded reliably where the older "Cloud
  Email and Collaboration Protection → Dashboard" path
  (`/app/email-and-collaboration/dashboard`) previously did not — its
  iframe would load but never become the visually active view no matter
  what click technique was tried (documented as an unresolved issue for
  several runs before this fix). The Overview page has 4 widgets: Top Users
  with Account Takeover Risks, Top Users with Targeted Attack Risks,
  Scanning Breakdown, Threats Detection Count — all scoped to a fixed
  ~7-day window shown at the top (not adjustable via a visible date
  picker). **If all 4 widgets show "No data to display."** (with a
  "Reload" button, not an error), set `email.applicable = false` — this is
  the common case, not a pull failure; don't fabricate zeros or leave stale
  placeholder numbers. Only set `applicable = true` (and fill in
  `scanned`/`threats_total`/`threats`/`recipients`) if the widgets actually
  populate with figures.
- **Credits (slide 19):** TrendAI Flex Licensing → Platform Usage and Credits
  (read the headline tiles via screenshot, or the same-origin `/ui/bc/api/credit/...`
  endpoints — `purchases`, `statistic/credit/summary`, `statistic/plannedCredit/monthly`
  — for exact figures instead of eyeballing).
- **EOY (slide 20):** computed automatically from the credits data — no manual input.

### Site-wide quirks discovered on this dashboard (apply to ANY tenant)
- **A `.iframe-overlay` div sits on top of the whole Cyber Risk Overview iframe
  and silently swallows clicks** (they report hitting `iframe#__SASE_ES_CONTAINER`
  no matter where you click, and nothing happens). Fix once per session:
  `document.querySelector('.iframe-overlay').style.pointerEvents='none'` via
  `browser_evaluate`.
- **Clicking the category cards (Devices/Internet-Facing Assets/Accounts/
  Applications/Cloud Assets) by screen coordinate is unreliable** even after
  disabling the overlay — coordinates silently select the wrong tab. Instead,
  find the actual element by text inside the dashboard's iframe and click its
  clickable ancestor directly:
  ```js
  () => { const iframes = document.querySelectorAll('iframe');
    for (const f of iframes) { try { const doc = f.contentDocument; if (!doc) continue;
      const els = Array.from(doc.querySelectorAll('*')).filter(el =>
        (el.textContent||'').trim() === 'Cloud Assets' && el.children.length===0);
      if (els.length) { let node = els[0]; for (let i=0;i<3;i++) node = node.parentElement;
        node.click(); return 'clicked'; } } catch(e) {} } return 'not found'; }
  ```
  (Walk up exactly 3 parents from the leaf text node to reach the element with
  `onclick` — verify with a quick parent-chain dump if a different tenant's DOM
  nesting differs.) Clicking a card only swaps the *displayed* cached data for
  some widgets but does trigger a fresh fetch for that category's own
  `riskFactor`/`riskLevelBar` endpoint — capture network requests right after
  the click, not just at initial page load.

### First run for this tenant — expect to discover quirks
The build script already has graceful fallbacks for a few known patterns (empty
email recipients, zero intel matches). If {{CUSTOMER}}'s tenant has its own
quirks (a dashboard widget that's structurally empty, a metric that dominates
differently than usual, missing MDR, etc.), handle them in **this customer's
own copy** of `build_value_report.py` — don't assume other customers behave
the same way. Document what you find in this section so the next rerun doesn't
have to rediscover it.

<!-- Add discovered quirks below this line as they come up. -->

---

## Rebuild behavior (important)
Re-running the script **preserves your hand-edits**:
- **Slide 1 (title):** the **date auto-updates**, but the **presenter name and
  everything else you typed is kept**.
- **Last slide (Account Team):** fully preserved — fill in the contacts once
  and they survive every rerun.
- All slides in between (currently 19, but varies with `len(risk_detail)` —
  see the slide-count note above) are regenerated from `value_report_data.json`.
- (First run, or if the deck is missing, or the slide count changed since last
  run, does a full build from the template instead of preserving anything.)

## Manual steps each run
1. Ask whether this customer has MDR; if yes, download the MDR monthly PDF (Claude will prompt you).
2. After the build, open the deck and fill in **slide 1 presenter** (should
   already be pre-filled from the answer given when the report was first set
   up) and the **last slide's Account Team** contacts (only needed the first
   time — they persist after that).
3. No PowerPoint renderer on this machine — Claude verifies geometry only;
   eyeball the deck in PowerPoint.
