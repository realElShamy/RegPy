---
name: three-statement-model
description: Build a fully linked, self-balancing three-statement financial model (income statement, balance sheet, cash flow statement + supporting schedules) as a professionally formatted, best-practice Excel workbook — a cover/contents sheet, the model, an audit/Checks sheet that rolls every integrity tie-out into one status light, an executive Dashboard (KPI tiles + charts incl. cash-flow & profit waterfalls), and a Sensitivity sheet (scenarios, two-way tables, tornado) — from a connected accounting system (Zoho Books or any ERP MCP connector), uploaded financials, or manual inputs, with jurisdiction-aware tax and accounting treatment focused on Egypt and the GCC (EAS, IFRS, Zakat, corporate tax, VAT context). Use whenever the user asks for a financial model, 3-statement model, financial forecast/projection, budget model, feasibility study model, dashboard, or asks to "model" a company's financials.
---

# Three-Statement Model Builder

You build annual three-statement models to a proven, self-balancing template. The
build is **deterministic**: a verified Python script generates the workbook and a
verified harness accepts or rejects it. The default deliverable is a complete,
best-practice **multi-sheet workbook** — Cover · Model · Checks · Dashboard ·
Sensitivity — built to recognised modelling standards (FAST/ICAEW/SMART:
colour-coded cells, one-row-one-formula, aggregated integrity checks, dashboards,
scenario/sensitivity analysis; see `references/BEST_PRACTICES.md`). Your job is
everything around the build: resolving the jurisdiction, getting the data,
constructing the input payload, choosing defensible forecast assumptions, running
the scripts, and delivering the result with honest disclosures. Never hand-write
the workbook cell by cell and never skip validation.

All paths below are relative to this skill's directory. Do your work (payloads,
built workbooks, logs) in a scratch/output directory OUTSIDE the skill folder —
skill directories are often read-only and must not be polluted; invoke `scripts/`
and reference `assets/` by absolute path.

## Pipeline (follow in order; each stage gates the next)

### Stage 0 — Understand the ask
Establish: which company, which country/jurisdiction, how many historical years exist,
how many forecast years are wanted (default 5), currency and units (default: local
currency, thousands). If the user's ask is exploratory ("can you build me a model?"),
confirm these five parameters in ONE compact question round, not a drip of questions.

### Stage 1 — Resolve the jurisdiction
1. Determine the country: explicit user statement > connected accounting system's
   organization country > currency hints > ask.
2. Load the matching pack from `assets/jurisdictions/`:
   `eg` (Egypt), `sa` (Saudi Arabia), `ae` (UAE), `qa` (Qatar), `kw` (Kuwait),
   `bh` (Bahrain), `om` (Oman), `intl` (IFRS/generic fallback for anything else).
3. Read `references/JURISDICTIONS.md` for that country: it explains the tax line
   treatment (CIT vs Zakat vs blended), the default rate the pack encodes, when to
   override it (free zones, small-business thresholds, mixed ownership, DMTT), and
   the disclosure language to include with the deliverable.
4. If the entity's situation deviates from the pack default (e.g. UAE qualifying
   free-zone person, KSA mixed ownership split, Egypt exporter incentives), compute
   the effective rate per JURISDICTIONS.md and put it in the payload's
   `tax_pct_ebt` explicitly — the pack default only applies when you leave
   `tax_pct_ebt` empty.
5. Set the payload's `units` (e.g. `"EGP thousands"`) AND, whenever the currency
   or scale differs from the pack's `units_label` (always the case for the `intl`
   pack outside USD, or for millions presentation), set the payload's
   `units_label` (e.g. `"EGP m"`, `"USD '000"`) — it overrides the pack's label
   on the workbook's unit-suffixed rows in both scripts.

### Stage 2 — Acquire historical data
Follow `references/DATA_SOURCES.md`. Discover what's connected before assuming:
use the harness's tool-discovery mechanism if it has one (e.g. ToolSearch in
Claude Code), otherwise scan the available tool list for accounting/ERP, drive,
and file connectors (accounting systems → files/drives → manual entry as
fallback). Build the historical
arrays of the input contract (`assets/inputs.schema.json`), folding real trial
balances into the template's line items per the playbook's mapping rules, and record
every mapping judgement for the final report.

**Hard gate:** the historical payload must satisfy the §1 identities of
`references/TEMPLATE_SPEC.md` (balance-sheet identity, cash/PP&E/debt/retained-
earnings roll-forwards, all ±1 unit). `scripts/build_model.py` enforces them and
exits non-zero listing each break. Fix breaks by correcting the mapping (see
DATA_SOURCES.md), never by plugging numbers.

### Stage 3 — Set forecast assumptions
Derive defaults from history, then adjust to the user's stated plans:
- Revenue growth: recent trajectory unless the user gives targets.
- COGS % / salaries % of revenue: recent averages unless told otherwise.
- Rent & overhead: last actual, escalated if the user indicates.
- D&A % of opening PP&E and interest % of average debt: back-calculated recent rates.
- Working-capital days: recent actuals.
- Capex, debt and equity movements: user's plans; otherwise hold capex at recent
  levels and leave financing flat.
- Tax: leave `tax_pct_ebt` empty to accept the jurisdiction default, or set the
  computed effective rate from Stage 1.
Present the assumption set to the user for confirmation when the session is
interactive; proceed with documented defaults when it is not.

### Stage 4 — Build
Default (the full best-practice workbook — Cover · Model · Checks · Dashboard ·
Sensitivity):
```bash
python3 scripts/build_workbook.py --inputs INPUTS.json \
    --jurisdiction assets/jurisdictions/<iso2>.json --output MODEL.xlsx \
    [--prepared-by "Name"]
```
Requires only python3 + openpyxl. It validates inputs (§1), fills the tax default
from the pack if `tax_pct_ebt` is empty, applies the pack's label/units overrides,
and writes the whole workbook in one pass: the byte-identical validated model
sheet (blue inputs, black formulas, borders, conditional-formatted balance check,
two charts) **plus** the Cover (contents, colour legend, disclosure), the Checks
audit sheet (15 integrity tie-outs → one OK/ERROR light + amber plausibility
alerts), the Dashboard (KPI tiles + 6 charts incl. live cash-flow & profit
waterfalls), and the Sensitivity sheet (scenario summary, two-way heat-map tables,
tornado). See `references/BEST_PRACTICES.md`, `references/DASHBOARD.md`,
`references/SENSITIVITY.md` for the conventions each applies.

Core-only alternative (just the single model sheet, e.g. for embedding):
```bash
python3 scripts/build_model.py --inputs INPUTS.json \
    --jurisdiction assets/jurisdictions/<iso2>.json --output MODEL.xlsx
```
Both reuse the same `populate_model_sheet`, so the model sheet is identical either
way. (Build all sheets in one pass — never load-and-resave to "add" sheets:
openpyxl drops charts on load.)

### Stage 5 — Validate (mandatory, never skip)
For the full workbook:
```bash
python3 scripts/validate_workbook.py MODEL.xlsx --inputs INPUTS.json \
    --jurisdiction assets/jurisdictions/<iso2>.json
```
It runs the core acceptance gate (`validate_model.py`) on the model sheet, then
independently re-checks the supporting sheets: the Checks master recalculates to
"OK" with no ERROR, every Dashboard KPI equals an independent simulation, and
every Sensitivity value matches the deterministic economics. For a core-only build
run `validate_model.py MODEL.xlsx --inputs … --jurisdiction …` directly.

Exit code 0 required. The harness re-derives every number with an independent
simulation, checks the formula canon and cross-column consistency, and recalculates
the workbook — this REQUIRES either LibreOffice headless (`soffice`) or the Python
`formulas` package; if neither is installed, install one (`pip install formulas`
is the lightweight option) before validating. If validation fails, read the
failure list and fix the payload/mapping or report the data problem — do not edit
the workbook by hand and do not deliver a failing model. If the recalculation
method reported is `cached-partial`, the computed values were NOT verified —
treat that as a failure, not a pass.

### Stage 6 — Deliver
Send the workbook to the user (and to a connected drive if they want it filed).
Point them at the Cover sheet (contents + colour legend) and the Checks status
light. Accompany it with a compact report:
1. Forecast net earnings and closing cash by year.
2. The driver set used (growth, margins, days, capex, financing, tax rate).
3. Jurisdiction disclosure from JURISDICTIONS.md (tax basis, VAT note, standards
   context, and the "model-level simplification, not tax advice" caveat).
4. Data-mapping disclosures from Stage 2.
5. Validation status (harness pass, recalc method) and a one-line read of the
   Sensitivity sheet (what most moves the answer, per the tornado).
6. The honesty caveats for the enhanced sheets: the Sensitivity grids are
   computed at build time (a live-Excel Data Table recipe is on the sheet), and
   scenario switching is a documented extension (`references/SENSITIVITY.md` §5).

## Deep references (read on demand, not all upfront)
- `references/TEMPLATE_SPEC.md` — the normative template: input contract, cell map,
  build phases, formula canon, formatting spec, self-validation protocol. Read §1
  before Stage 2 and skim the rest the first time you use this skill.
- `references/MODEL_ANATOMY.md` — why the template is shaped this way (dependency
  graph, corkscrews, sign conventions, no-circularity design). Read when you need to
  explain or defend the model's mechanics.
- `references/BEST_PRACTICES.md` — the modelling-standards guide behind the
  workbook: standards (FAST/ICAEW/SMART), formatting, colour code, number formats,
  signs, layout/white-space/alignment, sheet architecture & interlinking & input
  sheets, error-checking & auditing. Read when adapting formatting/structure or
  explaining why the model is built the way it is.
- `references/DASHBOARD.md` — visualisation & dashboarding: chart choice, KPI
  tiles, the waterfall (invisible-base) technique, alignment/sizing. Read when
  changing the Dashboard.
- `references/SENSITIVITY.md` — scenario & sensitivity analysis: data tables,
  tornado, live-Excel Data Table and CHOOSE scenario-switch recipes. Read when
  changing the Sensitivity sheet or a user wants live scenarios.
- `references/JURISDICTIONS.md` — Egypt + GCC accounting standards and tax regimes,
  pack semantics, override recipes, disclosure texts.
- `references/DATA_SOURCES.md` — MCP connector playbook and trial-balance mapping
  rules.

## Scope and honesty rules
- The template is an annual, single-entity, going-concern operating model: one tax
  line as % of EBT, term debt only (no revolver), no dividends line, no OCI, no
  consolidation, no monthly granularity. If the user needs those, say plainly that
  this template doesn't cover it, deliver what it does cover, and note the extension.
- Jurisdiction packs encode model-level simplifications of tax law. Always include
  the disclosure text; never present the output as tax or audit advice.
- Zakat, DMTT/Pillar-Two, free-zone and mixed-ownership treatments follow
  JURISDICTIONS.md recipes — apply them through the effective tax rate and say so.
- Report validation results truthfully, including the recalculation method used.
