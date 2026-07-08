# Sensitivity & Scenario Analysis

A model that produces a single number is a guess with decimal places. The value
of a three-statement model is showing **how the answer moves when the assumptions
move**. This reference explains the three standard techniques, how this skill
implements them, and how to make them live inside Excel.

The skill builds a **Sensitivity** sheet (see `scripts/build_workbook.py`,
`build_sensitivity`) with a scenario summary, two two-way tables, and a tornado
chart. Everything on it is computed deterministically at build time and
independently re-checked by `scripts/validate_workbook.py`.

---

## 1. The three techniques

| Technique | Question it answers | This skill |
|---|---|---|
| **Scenario analysis** | "What happens under Base / Upside / Downside — several drivers moving together?" | Scenario summary table |
| **Sensitivity (data) tables** | "How does one output respond to two drivers swept across a range?" | Two-way heat-mapped grids |
| **Tornado analysis** | "Which single driver moves the output most?" | Tornado chart |

They are complementary: scenarios tell a *story* (a coherent world), sensitivity
tables show the *response surface*, tornados **rank** the drivers so you know
which assumptions to defend hardest.

---

## 2. Scenario analysis

A scenario shifts **several** drivers together into a coherent narrative. The
skill ships three (`SCENARIOS` in `build_workbook.py`), each documented on the
sheet so the reader sees exactly what changed:

| Scenario | Drivers moved |
|---|---|
| Base case | assumptions as entered |
| Upside | revenue growth +2pp, COGS −2pp of revenue |
| Downside | revenue growth −2pp, COGS +3pp, interest +1pp |

Each row reports final-year **Revenue, EBITDA, Net earnings, Closing cash** and
the **minimum forecast cash** (the liquidity low-point — often the real risk).
Values are produced by re-running the *same* validated model economics
(`validate_model.simulate`) with the shifted drivers, so a scenario is a full
re-solve of all three statements, not a back-of-envelope tweak.

**Choosing scenario drivers.** Move the drivers that (a) matter — see the tornado
— and (b) co-move in reality. In a downturn, revenue growth falls *and* margins
compress *and* borrowing costs rise; a credible downside moves all three. Do not
build a downside by moving one driver you happen to have a number for.

---

## 3. Sensitivity (two-way data) tables

A two-way table sweeps **one output** against **two drivers** — the classic
"data table" of financial modelling. The skill builds two:

1. **Net earnings (final year)** vs **revenue growth** (rows, −4pp…+4pp) ×
   **COGS % of revenue** (columns, −4pp…+4pp).
2. **Minimum forecast cash** vs **revenue growth** × **interest rate**
   (−2pp…+2pp) — a liquidity-stress view.

The centre cell is the base case; a red-yellow-green **colour scale** (Excel
conditional formatting) turns the grid into a heat map so the reader sees the
gradient at a glance. Reading a two-way table well:

- **Diagonals** show combined moves (both drivers adverse ⇒ the corner).
- **Steeper axis = more sensitive driver** (this is the tornado, seen sideways).
- Watch for **sign flips** and **cliff edges** (e.g. the cell where minimum cash
  first goes negative — that is the covenant/financing trigger).

### Why the grids are computed, not live — and how to make them live

`openpyxl` (and any programmatic `.xlsx` writer) **cannot emit a native Excel
What-If Data Table** — the `{=TABLE(...)}` array is a feature of the Excel
calculation engine, not a storable formula. So the skill computes each grid cell
by re-solving the model in Python and writes the numbers, clearly labelled
"computed at build time." To convert a grid to a **live** Excel data table the
reader can recalculate:

1. Put the output formula in the grid's **top-left corner**, e.g.
   `='Three Statement Model'!<lastcol>38` (final-year net earnings).
2. List the **row-input values** down the left column, the **column-input
   values** across the top row.
3. Select the whole block (corner + both axes + the empty body).
4. **Data ▸ What-If Analysis ▸ Data Table**; set *Row input cell* and *Column
   input cell* to the two driver cells the axes represent.

Excel then fills the body live. (Caveat: a data table pointed at a single
driver cell varies *that year's* driver only; to sweep a driver across *all*
forecast years, first route every forecast year's driver through one master
cell, then point the data table at the master.)

---

## 4. Tornado analysis

A tornado **ranks drivers by impact**. For each driver the skill re-solves the
model at driver − Δ and driver + Δ, records the two resulting final-year net
earnings, and draws a horizontal bar from the low to the high outcome. Bars are
sorted longest-on-top, producing the funnel shape that names the chart. It is
built with the invisible-base stacked-bar technique (`_tornado`): one series for
`low − base` (extends left, red) and one for `high − base` (extends right,
green), so the bars pivot around the base case at zero.

Drivers swept (`TORNADO_DRIVERS`): revenue growth ±2pp, COGS ±2pp, salaries
±2pp, rent ±10%, D&A rate ±5pp, interest ±2pp, tax ±5pp. Read it top-down: the
top bar is the assumption whose error costs you the most — the one to research,
source, and defend first.

**One-at-a-time (OAT) caveat.** A tornado moves one driver while holding the
rest at base. It therefore ignores **interactions** (e.g. growth and margin
adverse together are worse than the sum of the two bars). The two-way tables
cover the most important interactions; for full interaction analysis you need
Monte Carlo (out of scope for this template — mentioned for completeness: it
would sample all drivers jointly from distributions and build an output
histogram).

---

## 5. Live scenario switching (CHOOSE pattern)

The skill keeps the model sheet's forecast assumptions as **hardcoded blue
inputs** (single scenario), because that is what the validated template
guarantees and what keeps the sheet free of circularity. To make scenarios
*switch live* inside Excel without breaking that discipline, add a thin control
layer on a separate inputs sheet:

1. A **switch cell** with a Base/Upside/Downside dropdown (Data ▸ Data
   Validation ▸ List), returning 1/2/3.
2. For each assumption, three columns (base, up, down) of hardcoded values.
3. The **active** assumption feeding the model =
   `CHOOSE($switch, base, up, down)`.
4. Point the model's driver cells at the active column.

This is the institutional "scenario manager without the Scenario Manager"
pattern: one cell flips the entire model between coherent worlds, every
assumption stays visible and colour-coded, and there is still no circularity.
Because it changes the model's blue inputs into links, apply it as an *extension*
of the delivered workbook, not inside the validated single sheet.

---

## 6. What the validator checks

`scripts/validate_workbook.py` independently re-computes every scenario value,
every grid cell, and every tornado endpoint from `simulate` and asserts the
sheet's numbers match to the cent, and ties the base-case column back to the
live model's final-year figures. If a driver perturbation or a grid axis is
mis-wired, the build fails rather than shipping a plausible-looking but wrong
sensitivity table.
