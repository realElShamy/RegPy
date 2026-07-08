# Three-Statement Financial Model — Claude Code plugin

Installable packaging of the `three-statement-model` skill. The skill builds a
fully linked, self-balancing, **best-practice multi-sheet** three-statement model
workbook — Cover · Model · Checks (audit) · Dashboard · Sensitivity — with
jurisdiction-aware tax and accounting treatment for **Egypt and the GCC** (EAS,
IFRS, Zakat, corporate tax, DMTT, VAT), driven by whatever data source the session
has connected (accounting/ERP MCP connectors, uploaded financials, or manual input).

The build is deterministic (a verified Python builder) and every workbook is put
through a verified acceptance harness before delivery.

## Install

This repository is a Claude Code plugin marketplace. In Claude Code:

```
/plugin marketplace add realElShamy/RegPy
/plugin install three-statement-model@regpy-marketplace
```

Then invoke it by asking for a financial model, or run `/three-statement-model`.

See [`INSTALL.md`](../../INSTALL.md) at the repository root for all install
methods (plugin, personal skill, project skill, portable copy into any agent
harness).

## What's inside the skill

```
skills/three-statement-model/
├── SKILL.md                      # orchestrating pipeline (entry point)
├── references/                   # TEMPLATE_SPEC, MODEL_ANATOMY, JURISDICTIONS,
│                                 #   DATA_SOURCES, BEST_PRACTICES, DASHBOARD, SENSITIVITY
├── scripts/                      # build_workbook.py / validate_workbook.py (full),
│                                 #   build_model.py / validate_model.py (core single sheet)
└── assets/                       # inputs.schema.json, example_inputs.json,
                                  #   jurisdictions/ (eg, sa, ae, qa, kw, bh, om, intl)
```

## Requirements

- `python3` + `openpyxl` to build.
- LibreOffice headless **or** the `formulas` pip package to validate (the validator
  detects what is present and reports which engine verified the numbers).

## Canonical source

The canonical, auto-loading project copy of this skill lives at
`.claude/skills/three-statement-model/` in the repository root. This plugin copy is
kept in sync with it by [`plugins/sync_plugin.sh`](../sync_plugin.sh); edit the
canonical copy, then re-run the sync script.
