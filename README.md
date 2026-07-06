# Value Report — Claude Code plugin

Builds a dark-theme **TrendAI Vision One™ customer value-report PowerPoint deck**
for any customer, straight from live portal data. It asks who the report is for
and who's presenting, pulls the numbers from Vision One via Blueprint MCP, and
runs a bundled build script that writes `~/Desktop/<Customer> Value Report.pptx`.

This is the generalized version of the per-customer decks (John H. Carter,
Intralot) — one skill, any tenant.

## Prerequisites

Each user needs these on their own machine — the plugin does **not** install them:

1. **Blueprint MCP** configured in Claude Code and logged into *your own*
   TrendAI Vision One tenant (`portal.xdr.trendmicro.com`). The skill drives the
   portal through it to pull live data.
2. **Python 3** with two packages:
   ```bash
   pip3 install python-pptx pypdf
   ```
   (`python-pptx` builds the deck; `pypdf` parses the MDR monthly report PDF.)

## Install

```
/plugin marketplace add DebLarkAU/claudevaluereport
/plugin install value-report@trendai-value-report
```

- The first line registers this repo as a plugin marketplace.
- The second installs the `value-report` plugin from the `trendai-value-report`
  marketplace defined in this repo.

To try it locally before pushing, from inside a clone of this repo:
```
/plugin marketplace add .
/plugin install value-report@trendai-value-report
```

## Use

In any Claude Code session:

> Run a value report

or `Create a value report for <customer>`. The skill will:
1. Ask which customer and who's presenting.
2. Set up (or reuse) a working dir at `~/Documents/<Customer>/value_report/`,
   copying the bundled template/script/runbook into it on first run.
3. Pull live data from Vision One (asking about MDR explicitly), fill in
   `value_report_data.json`.
4. Build `~/Desktop/<Customer> Value Report.pptx`.

Reruns preserve the title slide and Account Team slide; only the 8 content
slides are rebuilt in place.

## What's in here

```
.claude-plugin/
  plugin.json          plugin manifest
  marketplace.json     self-listing marketplace so the repo installs directly
skills/value-report/
  SKILL.md             the skill (references bundled assets via ${CLAUDE_PLUGIN_ROOT})
assets/
  build_value_report.py            generic deck builder (finds its own siblings)
  template.pptx                    base deck: title slide + layouts
  value_report_data_template.json  data skeleton (placeholders)
  RUNBOOK_TEMPLATE.md              generic per-customer runbook ({{CUSTOMER}} tokens)
  PORTAL_EXTRACTION.md             Blueprint MCP techniques for the Vision One portal
  BUILD_NOTES.md                   slide-by-slide notes, dashboard quirks, render trick
```

The per-customer working directory (`~/Documents/<Customer>/value_report/`) is
created outside the plugin and persists across plugin updates, so a `/plugin
update` never clobbers a customer's filled-in data.

## Updating

Push changes to the repo, then users run:
```
/plugin marketplace update trendai-value-report
/plugin update value-report@trendai-value-report
```
