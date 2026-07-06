---
name: value-report
description: Build a TrendAI Vision One customer value-report PowerPoint deck for ANY customer ("Run a value report" / "Build a value report" / "Create a value report for <customer>"). Asks who the customer is and who's presenting, then re-pulls Vision One portal data and runs the generic build script.
---

# Value Report — generic (any customer)

Builds a 21-slide (varies with `risk_detail` length — see `RUNBOOK_TEMPLATE.md`),
dark-theme TrendAI Vision One™ value-report deck for **any customer**, by asking
who it's for instead of being hardcoded to one tenant.

**Bundled assets (this plugin):** the build script, template, data skeleton,
runbook, and portal-extraction notes all ship inside this plugin under
`${CLAUDE_PLUGIN_ROOT}/assets/`. Nothing needs to pre-exist on the machine
except the runtime prerequisites below.

**Prerequisites (each user must have these):**
- **Blueprint MCP** configured and logged into *their own* TrendAI Vision One
  tenant (`portal.xdr.trendmicro.com`) — used to pull live portal data.
- **Python 3** with `python-pptx` (deck build) and `pypdf` (MDR report parsing):
  `pip3 install python-pptx pypdf`.

**Bundled assets:** `${CLAUDE_PLUGIN_ROOT}/assets/` —
  - `build_value_report.py` — fully generic: reads `customer`/`presenter` from
    the data file, writes `~/Desktop/<customer> Value Report.pptx` (no
    per-customer script edits needed).
  - `template.pptx` — title slide + all slide layouts, incl. "Quote 01".
  - `value_report_data_template.json` — skeleton with placeholder values.
  - `RUNBOOK_TEMPLATE.md` — generic runbook with `{{CUSTOMER}}` tokens.
  - `PORTAL_EXTRACTION.md` — Blueprint MCP techniques for the Vision One portal
    (nested same-origin iframes, DOM-walk clicks, react-select filters, etc.).
  - `BUILD_NOTES.md` — slide-by-slide build notes, dashboard quirks, and the
    Keynote/AppleScript PNG-export render technique used to verify the deck.

**Per-customer working dir:** `~/Documents/<Customer>/value_report/` — created
on first run for a customer (by copying the bundled assets), reused on every
rerun. This is the user's own scratch area and persists across plugin updates.

## Step 0 — Ask who this is for (ALWAYS do this first, never assume or reuse stale answers)

Ask the user, before touching any data:
1. **Which customer** is this value report for?
2. **Who is presenting/creating** the report?

How to ask: check `~/Documents/*/value_report/` for customers that already
have a deck set up — if any exist, use `AskUserQuestion` with those as
quick-pick options for the customer question (the tool always adds a free-text
"Other" automatically, so a brand-new customer name is one click away). For the
presenter question, default to the current session user if known (check the
`userEmail` context) as one option, plus rely on "Other" for anyone else. If
offering multiple-choice options feels forced for either question (e.g. no
customers exist yet), just ask directly in plain conversation instead — the
point is that both facts come from the user this run, not from a prior run's
JSON or a guess.

## Step 1 — Set up (or reuse) the customer's working directory
Target: `~/Documents/<Customer>/value_report/` (use the customer name as given,
e.g. "Acme Corp" → `~/Documents/Acme Corp/value_report/`).

- **If it already exists** (rerunning for a known customer): reuse
  `build_value_report.py`, `template.pptx`, and `value_report_data.json` as-is.
  Update `value_report_data.json`'s `presenter` field if the answer to Step 0
  differs from what's currently in there.
- **If it doesn't exist** (first run for this customer):
  1. `mkdir -p` the directory.
  2. Copy `${CLAUDE_PLUGIN_ROOT}/assets/build_value_report.py` and
     `${CLAUDE_PLUGIN_ROOT}/assets/template.pptx` into it verbatim.
  3. Copy `${CLAUDE_PLUGIN_ROOT}/assets/value_report_data_template.json` to
     `value_report_data.json` in the new directory, then fill in `customer`,
     `presenter`, `report_title`, `month_year`, and `as_of` (today's date) from
     the Step 0 answers — leave every other section as placeholders for now,
     they get overwritten in Step 2.
  4. Copy `${CLAUDE_PLUGIN_ROOT}/assets/RUNBOOK_TEMPLATE.md` to
     `~/Documents/<Customer>/value_report/<Customer> Value Report.md`,
     replacing every `{{CUSTOMER}}` token with the actual customer name.

## Step 2 — Re-pull live data (per the "always rerun the full pipeline" rule)
Open the portal via **Blueprint MCP**, switch the org/tenant to the customer
named in Step 0, and follow the customer's own runbook
(`<Customer> Value Report.md`) for exactly which numbers go where. For the
portal-navigation mechanics (nested iframes, DOM-walk clicks, filters), see
`${CLAUDE_PLUGIN_ROOT}/assets/PORTAL_EXTRACTION.md`. On a first run, you're
discovering this tenant's quirks live (which endpoint action type dominates,
whether email/intel widgets are empty, etc.) — handle them in the customer's
own copy of `build_value_report.py` if they need bespoke logic, and note them
in the runbook so the next rerun doesn't have to rediscover them.

**When you reach the MDR slide: ASK the user first, every time — never infer
or reuse a prior answer.** Ask "Does {customer} have MDR (Managed Detection
and Response)?"
- **If yes:** ask the user to download the latest MDR Monthly Report PDF
  (Reports → Generated reports — downloads are blocked in automation) and
  wait for confirmation before reading it with `pypdf` and filling in `mdr`.
- **If no:** set `mdr.applicable = false` and (optionally) `mdr.not_applicable_note`
  in `value_report_data.json`. `build_mdr` renders a clean "not part of this
  subscription" slide instead of an empty/zeroed-out funnel — don't leave
  placeholder zeros in that case, they read as "MDR ran and found nothing"
  rather than "MDR doesn't apply here."

## Data quality rule: never leave placeholder text on a slide

If a figure genuinely can't be pulled this run (no data source, no widget
found, a connector is broken, etc.), **never leave bracketed placeholder
text like `"[ not captured this run ]"` or `"N/A"` baked into
`value_report_data.json`** — it renders verbatim on the slide and reads as
a bug, not as an intentional null state. This happened on the Identity
Security Posture slide's "Risky Accounts" column and the customer had to
point it out.

Instead, use an empty value (`[]`, `""`, `0`) for the field, paired with a
sibling `*_note` field explaining why (e.g. `not_applicable_note`,
`playbooks_note`, `integrations_note`, `risky_accounts_note`). The build
script already renders these gracefully — centered explanatory text instead
of the empty/fake list — for every slide that currently supports it
(Workflow and Automation, AI Security, Zero Trust, Standard Endpoint
Protection, Network Security, Email, Identity's Risky Accounts). **Apply
this same empty-value-plus-note pattern to any NEW optional/nullable field
you add to a future slide** — don't invent a new bracketed-placeholder
convention. Before pulling data on any field, do a real search for it
first (a genuinely missing data source is common — e.g. no risk-score
column, a broken connector, an unreleased pre-release feature — but always
check rather than assuming and defaulting straight to a placeholder).

## Step 3 — Build
```bash
cd "~/Documents/<Customer>/value_report" && python3 build_value_report.py
```
Writes `~/Desktop/<Customer> Value Report.pptx`. **Rebuild preserves slide 1
(title) and the last slide (Account Team)** — if the deck already exists with
the expected slide count, the script reopens it and rebuilds only the 8
content slides in place; the date on slide 1 is still refreshed. Hand edits to
presenter/team survive reruns.

## Step 4 — Verify and report
No PowerPoint renderer on this machine — render via the Keynote/AppleScript
PNG-export technique (documented in `${CLAUDE_PLUGIN_ROOT}/assets/BUILD_NOTES.md`)
and read the images rather than trusting geometry alone. Before reporting
done, grep `value_report_data.json` for stray bracketed placeholder text
(`grep -n '\[ ' value_report_data.json`) — anything matching means a field
still needs the empty-value-plus-note treatment described above, not a
literal placeholder string. Report the key figures pulled this run and the
output path.
