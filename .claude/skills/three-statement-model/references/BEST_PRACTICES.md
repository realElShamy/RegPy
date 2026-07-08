# Financial-Modelling Best Practices

The reference guide behind this skill. It collects the conventions that separate a
model people trust from a spreadsheet people fear — formatting, colour, layout,
sheet architecture, error-checking and auditing — and states, for each, both the
*principle* and the *concrete spec* the skill applies. Visualisation and
sensitivity have their own files (`DASHBOARD.md`, `SENSITIVITY.md`).

> A model's job is not only to compute the right number but to let a reviewer
> *see* that it is right. Every practice below serves transparency, consistency,
> flexibility or integrity — the four goals every recognised standard shares.

---

## 1. Standards & principles

Modelling is a discipline with published standards. They differ in detail but
agree on the fundamentals.

- **FAST Standard** (fast-standard.org) — **F**lexible, **A**ppropriate,
  **S**tructured, **T**ransparent. Its most operational rule is **one row, one
  formula, one calculation**: a row's formula is written once and copied across —
  never a different formula in different columns of the same row — so a reviewer
  who checks one cell has checked the row. Also: inputs separated from
  calculations, short formulas built from intermediate rows rather than one
  mega-formula, consistent time series left-to-right.
- **ICAEW "Twenty Principles for Good Spreadsheet Practice"** — a chartered-
  accountant checklist: agree the purpose; adopt a standard and stick to it;
  separate inputs, calculations and outputs; use one formula per row/column; be
  disciplined about workbook structure; build in checks; document; control
  versions.
- **SMART** (Mazars/Corality) — **S**tructured, **M**anageable, **A**daptable,
  **R**obust, **T**ransparent — emphasises a clear inputs→calcs→outputs flow and
  a "no. surprises" review culture.
- **Operis / "The Operis Way"** and the **Institute for Financial Modelling** add
  rigour around review, versioning and audit.
- **Spreadsheet risk (EuSpRIG)** is why all of this matters: field studies find a
  large share of production spreadsheets contain material errors. Structure,
  consistency and built-in checks are the antidote — they make errors *visible*.

Common thread: **transparency + consistency + separation of concerns + built-in
integrity checks**. The rest of this document is those four ideas made concrete.

---

## 2. Formatting, fonts, colour, number formats, signs, borders

Formatting in a model is not decoration — it is a **legend**. A consistent scheme
lets a reviewer decode a cell's role from its appearance alone.

### 2.1 The colour code (the single most important convention)

Font colour encodes **what kind of cell this is**, so a reviewer can tell an
assumption from a calculation at a glance and knows exactly which cells are safe
to change:

| Font colour | Hex | Meaning |
|---|---|---|
| **Blue** | `#0000FF` | **Hardcoded input / assumption** — a typed number you may change |
| **Black** | `#000000` | **Formula on this sheet** — a calculation; don't overtype |
| **Green** | `#008000` | **Link to another sheet** in this workbook |
| **Red** | `#FF0000` | **Link to another workbook** / external file, or a warning |
| (Purple/grey, house-specific) | — | inputs pulled from another file, or flags |

Rule of thumb: **if it's blue it's a number you can touch; if it's black it's a
formula you shouldn't.** This skill applies it strictly — every blue cell is a
typed constant, every black data cell is a formula, and cross-sheet links on the
Cover/Checks/Dashboard are green. The Cover sheet prints the legend so the reader
knows the code.

Cell **fills** are secondary and lighter-touch: a pale fill to mark an input
block or a banner; the Excel *Good/Bad/Neutral* palette for status flags — green
`#C6EFCE`/font `#006100`, red `#FFC7CE`/`#9C0006`, amber `#FFEB9C`/`#9C6500`.
Don't fill every input cell a garish yellow; the font colour already carries the
message.

### 2.2 Number formats (use custom format strings, never round the data)

Formatting changes *display*, never the stored value — never type `1.2` for
1,200. Standard custom format strings:

| Use | Format string |
|---|---|
| Money, thousands, `()` negatives, dash for zero | `_-* #,##0_-;\(#,##0\)_-;_-* "-"_-;_-@` |
| Percent (1 dp) | `0.0%` |
| Multiple / turns | `0.0"x"` |
| Ratio | `0.00` |
| Balance-check (signed, tight) | `0.0000_ ;\-0.0000\ ` |
| Date | `dd-mmm-yyyy` or `mmm-yy` for headers |
| Signed delta (sensitivity axes) | `+0.0%;-0.0%` |

The leading `_-* ` / trailing `_-` give **hanging indentation** so figures align
on the decimal regardless of sign, and the `"-"` shows a dash for exact zeros
(cleaner than `0`). State **units once** in a header ("EGP '000") rather than
repeating a currency symbol on every row.

### 2.3 Signs

Pick a sign convention and hold it. This template's (documented on the Cover and
in `MODEL_ANATOMY.md`): **expenses are shown positive and subtracted in the
subtotal rows**; **capex is shown positive** under investing and subtracted in
the net-cash row; **issuance is +, repayment is −** on a single combined row; an
increase in working capital *consumes* cash (subtracted). The virtue is not which
convention you pick but that it never changes mid-model.

### 2.4 Fonts, borders, alignment, indentation

- **One body font**, consistently sized (this template: Arial Narrow 12; section
  banners Open Sans 14 bold). Don't mix fonts within the body.
- **Borders mark structure, not everything.** Thin **top** border above a
  subtotal; thin top + **double bottom** under a grand total; a thin bottom to
  close a summed block. No boxes around every cell.
- **Bold** subtotals/totals and section headers; leave line items regular.
- **Indent** sub-items under their header (line items under "Expenses").
- **Right-align numbers** (via the number format's spacing), left-align labels.
- **Alignment discipline** is what makes a model look engineered rather than
  assembled — see §3.

---

## 3. Layout, white space, alignment, print & navigation

### 3.1 The grid: time across, items down

- **One period per column, time left-to-right**, the *same* columns for the same
  years on every schedule, so a period is one vertical slice through the whole
  model. Historical columns then forecast columns, visually distinguished (this
  template shades the forecast year headers navy).
- **Labels in column A; B/C spacer columns** so long labels overflow cleanly and
  the number columns start at a consistent position.
- **Section banners** introduce each block (Assumptions, IS, BS, CFS, Schedules).

### 3.2 White space / "empty areas"

Deliberate blank rows and columns are structure, not waste. A blank **separator
row** between sections, a spacer column between labels and data, and even gutters
between dashboard objects **group** related content and give the eye rest.
(This template's `MODEL_ANATOMY.md` documents exactly which rows are intentional
separators — SUM ranges are drawn to *avoid* swallowing them.)

### 3.3 Alignment of charts and images

Objects that don't line up make a model look untrustworthy. Snap every chart and
image to the **cell grid** (hold **Alt** while dragging in Excel), give charts a
**uniform size**, and anchor them so a row of charts shares top and bottom edges.
Keep a consistent gutter. The skill anchors its dashboard charts at fixed cells
with one width/height so the 2-across grid aligns. (More in `DASHBOARD.md` §6.)

### 3.4 Navigation, print, view

- **Freeze panes** so headers/labels stay on screen while scrolling (this
  template freezes the top 3 rows, keeping years and the balance-check flag
  always visible — the check becomes a permanent dashboard).
- **Gridlines off**, a single zoom, for a clean presentation surface.
- **Group/outline** long schedules so they collapse to a summary (the skill
  groups the supporting-schedule and chart-feed rows).
- **Cover + contents** sheet with **hyperlinks** to every sheet; **tab colours**
  to distinguish input / model / check / output sheets at a glance.
- **Print setup:** landscape, fit-to-width, a **header/footer** carrying filename,
  sheet name, date and "Page X of Y", and a defined **print area** — so a printed
  or PDF'd model is as legible as the on-screen one.

---

## 4. Sheet architecture — inputs, calculations, outputs, and the prerequisite sheets

### 4.1 The discipline: separate Inputs, Calculations, Outputs (ICO)

The foundational structural rule of every standard: **keep inputs, calculations
and outputs distinct**, and let calculation flow **one direction** — top-to-
bottom, left-to-right, later sheets depending on earlier ones, never a backward
reference. This makes the model auditable (you can trace any output back to its
inputs) and flexible (change an input in one known place).

### 4.2 The prerequisite sheets a complete model has

A model that will interlink cleanly is built from a small set of standard sheets,
each with one job, wired in one direction:

| Sheet | Role | Links |
|---|---|---|
| **Cover / title** | Entity, currency, period, version, disclaimer, colour legend, **contents with hyperlinks**, and a live status light from Checks | reads Checks |
| **Contents / index** | Navigation (can live on the Cover) | hyperlinks to all |
| **Input / Assumptions** | *Every* hardcoded assumption, colour-coded, one source of truth; scenario switch lives here | feeds Calcs |
| **Timeline / dates** | The period axis all schedules reference | feeds all |
| **Calculations / model** | The three statements + supporting schedules | reads Inputs |
| **Checks** | Every integrity tie-out rolled to one status light | reads Calcs |
| **Outputs / Dashboard** | KPI tiles, charts, summary for the reader | reads Calcs |
| **Sensitivity / scenarios** | Scenario summary, data tables, tornado | reads Calcs |

This skill delivers **Cover · Model · Checks · Dashboard · Sensitivity** and the
Cover carries the contents + legend. The **single validated model sheet** keeps
its assumptions in a colour-coded **Assumptions block at the top** rather than on
a separate Input sheet — a deliberate, defensible choice for an annual single-
entity model (see §4.3); the Cover, Checks, Dashboard and Sensitivity are the
interlinking supporting sheets, each reading *from* the model and never writing
back to it.

### 4.3 Single-sheet vs multi-sheet — and where the input sheet goes

There are two schools, both legitimate:

- **Separate input sheet** (FAST/ICAEW default for larger models): all
  assumptions on one Inputs sheet; calculations link to them in green. Best when
  many schedules share inputs, when several people own different inputs, or when
  the model is large.
- **Single calculation sheet with an assumptions block** (this template): inputs
  live as blue cells at the top of the one sheet, each period in one column across
  all sections. The provider of the reference template *recommends* this for a
  compact model because it removes cross-sheet mis-linking risk and keeps a period
  in a single column. It is the right call for an annual, single-entity model; it
  does **not** scale to multi-entity or monthly models, where a separate input
  sheet wins.

Either way the *principle* is the same — inputs are **identified** (colour),
**centralised** (one place per input) and **flow one way** into the calculations.
To move this template toward a separate live input sheet with scenario switching,
see the CHOOSE pattern in `SENSITIVITY.md` §5.

### 4.4 Named ranges

Use **named ranges** for a handful of key, widely-referenced cells (final-year
revenue, net earnings, closing cash) so formulas read
`=FinalYearNetEarnings` rather than `='Three Statement Model'!$M$38`, and so
outputs survive row inserts. Don't name *everything* — over-naming hurts
auditability as much as under-naming. The skill defines
`FinalYearRevenue`, `FinalYearNetEarnings`, `FinalYearClosingCash` and
`ModelBalanceChecks`.

---

## 5. Error-checking & auditing

The difference between a model and a liability is that a model **tells you when
it's wrong**. Two layers: *checks* built into the workbook, and *auditing*
techniques you apply while reviewing.

### 5.1 Build checks in

- **Check cells** compute something that *must* be zero (or true) and flash if
  it isn't. The canonical one is the **balance-sheet check** — Total Assets −
  Total Liabilities & Equity = 0 — but a complete model checks far more:
  cash on the BS equals the CFS closing cash; balance-sheet PP&E/debt equal the
  schedule closing balances; IS D&A/interest equal the schedules; every corkscrew
  closes (`closing = opening + additions − reductions`); every roll-forward
  reconciles (retained earnings = prior + net income).
- **Aggregate every check into one master flag** — a single "OK/ERROR" the
  reviewer watches. Green means every tie-out in the model holds; red means at
  least one broke, and the per-check grid shows which. This skill's **Checks
  sheet** rolls **15 integrity tie-outs across every year** into one master light
  (via `COUNTIF(...,"ERROR")`), and keeps the model's own always-visible balance
  flag frozen at the top of the model sheet.
- **Separate errors from alerts.** A *check* failing means the model is
  mechanically broken (fix the wiring). An *alert* — negative cash, negative
  equity, a loss year, a margin outside 0–100% — means the model is *fine* but the
  business result deserves attention. The skill shows integrity checks in
  red/green and plausibility alerts in **amber** as a distinct "REVIEW/CLEAR"
  master, so a genuine break is never confused with a bad-but-correct outcome.
- **Sensible tolerances.** Compare with a small tolerance (this template: ±1 unit,
  i.e. ±$1k on $000s figures) so floating-point/rounding noise doesn't trip a
  check, but real breaks do.
- **Use IFERROR sparingly and locally.** Wrap `IFERROR` only around a cell that
  can *legitimately* error (a display flag, a ratio with a possible zero
  denominator) — never blanket-wrap calculations, because that **hides real
  errors**. The template wraps only the check *flag* and zero-denominator audit
  ratios, never the economics.
- **No hardcoded numbers buried in formulas.** A constant inside a formula (a tax
  rate, a `1.1` growth factor) is an invisible input. Put it in a blue input cell
  and reference it. (The template documents its five permitted literals — `1` in
  `(1+growth)`, `-1`/`+1` in the year and back-calc rows, the `1` check threshold,
  and `365` — and the validator enforces the list.)

### 5.2 Audit / review techniques

When reviewing (yours or someone else's model):

- **Trace Precedents / Dependents** (Formulas ▸ Formula Auditing) to see what
  feeds a cell and what it feeds — the fastest way to understand wiring.
- **Evaluate Formula** to step through a complex formula's calculation.
- **Show Formulas** (`Ctrl+``) to see the whole sheet as formulas at once and spot
  a cell that's a hardcode where its neighbours are formulas.
- **Go To Special** (`Ctrl+G ▸ Special`) to select all **constants**, all
  **formulas**, or all **errors** at once — the single best way to confirm the
  colour-coding is honest (select constants → they should all be blue) and to find
  stray hardcodes or `#REF!`s.
- **Check the consistency of a row**: FAST's one-formula-per-row means a
  column-shifted copy of any forecast cell should equal its neighbour. (This
  skill's validator asserts exactly that mechanically.)
- **Reconcile to source** for historicals and **re-implement independently** for
  the forecast — never trust the sheet's own numbers to validate the sheet. The
  skill's harness re-derives every number in plain Python and recalculates the
  workbook with a real engine before accepting it.

---

## 6. How this skill embodies the practices

| Practice | Where |
|---|---|
| One row / one formula (FAST) | model forecast block; `validate_model.py` asserts column-shift consistency |
| Blue-input / black-formula / green-link colour code | model sheet + Cover legend + Checks/Dashboard links |
| Custom number formats, signs, borders | `build_model.py` formatting spec; `MODEL_ANATOMY.md` §4–5 |
| White space, freeze panes, gridlines off, grouping, print setup | `build_model.py`, `build_workbook.py` (`polish_model_sheet`, `setup_prints`) |
| Cover + contents + colour legend + tab colours | `build_workbook.py` `build_cover` |
| Inputs identified & centralised; one-way flow | assumptions block; supporting sheets read-only from model |
| Named ranges for key outputs | `add_named_ranges` |
| Integrity checks aggregated to one master light | `build_workbook.py` `build_checks` (15 tie-outs) |
| Errors vs plausibility alerts separated | Checks sheet red/green vs amber |
| Limited, local IFERROR; no hidden hardcodes | template design; validator's formula-hygiene sweep |
| Independent re-derivation + real-engine recalc | `validate_model.py`, `validate_workbook.py` |
| Dashboards & advanced charts | `DASHBOARD.md` |
| Scenarios & sensitivity | `SENSITIVITY.md` |

### Honest limitations (state them, don't paper over them)

- The model is a **single sheet with an assumptions block**, not a separate live
  Input sheet; scenario switching is delivered as **precomputed** results plus a
  documented CHOOSE recipe, because programmatic `.xlsx` cannot embed live Excel
  Data Tables. Both trade-offs are documented on the sheet and in `SENSITIVITY.md`.
- The template is annual, single-entity, term-debt-only, no dividends/OCI/revolver
  — see `SKILL.md` "Scope and honesty rules".

---

## 7. Cheat-sheet

**Font colours** — input `#0000FF` · formula `#000000` · cross-sheet link
`#008000` · external/warning `#FF0000`.
**Status fills** — OK `#C6EFCE`/`#006100` · ERROR `#FFC7CE`/`#9C0006` · alert
`#FFEB9C`/`#9C6500`.
**Number formats** — money `_-* #,##0_-;\(#,##0\)_-;_-* "-"_-;_-@` · percent
`0.0%` · multiple `0.0"x"` · signed delta `+0.0%;-0.0%`.
**Tab colours** — cover/navy, model/blue, checks/green, dashboard/accent,
sensitivity/purple.
**The integrity checks a 3-statement model must carry** — balance sheet balances;
BS cash = CFS closing cash; BS PP&E/debt = schedule closings; IS D&A = schedule =
CFS add-back; IS interest = schedule; NWC = AR+Inv−AP; cash/PP&E/debt corkscrews
close; retained-earnings and equity rolls reconcile; SE = capital + retained
earnings — all aggregated to one master OK/ERROR light, with negative-cash /
negative-equity / loss-year / margin-out-of-range as separate amber alerts.
