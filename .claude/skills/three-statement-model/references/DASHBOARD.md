# Visualisation & Dashboarding

Charts are how a model *communicates*. A three-statement model that only prints
tables makes the reader do the seeing; a good dashboard does the seeing for them.
This reference covers the visualisation principles the skill follows, the charts
it builds, how to build the ones Excel has no native button for (waterfalls),
and how to lay a dashboard out so it reads as one system.

The skill builds a **Dashboard** sheet (`scripts/build_workbook.py`,
`build_dashboard`) with nine KPI tiles and six charts, all driven by **live
cross-sheet links** into the model, so the dashboard recalculates whenever the
model's blue inputs change.

---

## 1. Principles (apply before choosing a chart type)

1. **One message per chart.** Each chart answers one question ("is growth
   profitable?", "where does cash go?"). If you need a legend paragraph to
   explain it, split it.
2. **Maximise data-ink, minimise chart-junk** (Tufte). No 3-D, no gradient
   fills, no drop shadows, no gridline clutter. Every pixel should carry data.
3. **The 5-second rule.** A reader should get the headline of a dashboard in
   five seconds. That is what KPI tiles and a strict visual hierarchy buy you:
   biggest/top-left = most important.
4. **Consistent, meaning-carrying colour.** Pick one accent and reuse it; let
   red/green mean bad/good *consistently* (a cost is red everywhere). Keep the
   same series colour for the same series across every chart.
5. **Accessibility.** ~8% of men have red-green colour deficiency, so never rely
   on red-vs-green *alone* — pair it with position (waterfall up/down bars sit at
   different heights) or labels. Keep text-to-fill contrast high.
6. **Label directly where you can.** A number on the bar beats a trip to the
   axis; a line labelled at its end beats a legend.
7. **Show history and forecast distinctly.** A vertical divider or a fill/shade
   change at the forecast boundary tells the reader which columns are actuals and
   which are projections — the single most important context on a financial chart.

---

## 2. Which chart for which message

| Message | Chart | In the skill |
|---|---|---|
| Level over time (revenue, cash) | Column / line | Revenue bars |
| Two series on different scales (revenue **and** margin %) | **Combo** bar + line on a secondary axis | "Revenue & gross margin" |
| Composition that nets to a total (CFO/CFI/CFF → Δcash) | Stacked column (+ total line) | "Cash flow composition" |
| A total broken into additive steps (open → +/− → close) | **Waterfall / bridge** | "Cash-flow waterfall", "Profit waterfall" |
| Ranking of driver impact | **Tornado** (horizontal bars, sorted) | Sensitivity sheet |
| Two-driver response surface | **Heat-mapped table** (colour scale) | Sensitivity grids |
| A single number vs a target/prior | **KPI tile / card** | 9 tiles |
| In-cell micro-trend | **Sparkline** | see §5 |
| Valuation range across methods | Football-field (floating bars) | (extension) |

Avoid: pie charts for anything with >3 slices or where comparison matters
(angles read poorly); dual-axis combos where the two scales invite a false
correlation; stacked areas with many bands.

---

## 3. The KPI tile row

Nine tiles across the top, each a small card (title / big value / sub-metric):
Revenue (+ revenue CAGR), Gross margin, EBITDA (+ margin), Net earnings (+
margin), Closing cash, Total assets, Return on equity, Current ratio, Net debt /
EBITDA. All are **cross-sheet formulas** on the final forecast year, e.g.
`='Three Statement Model'!<lastcol>38` for net earnings, with EBITDA rebuilt as
`EBT + interest + D&A`. Ratios are wrapped in `IFERROR(...,"n/a")` so a
zero-denominator company (no debt, no equity) shows "n/a" rather than `#DIV/0!`.

Tile design: left-aligned, generous whitespace, one strong navy number, a muted
grey sub-metric. Group them on one row so the eye scans left-to-right in order of
importance (top-line → profitability → liquidity → returns → leverage).

---

## 4. Building a waterfall without a native waterfall

Excel gained a native Waterfall chart in 2016, but programmatic writers
(`openpyxl`) can't emit it, and older Excel/LibreOffice/Google Sheets don't have
it. The portable, universally-rendering technique is the **invisible-base stacked
column** (`_waterfall_block` / `_waterfall_chart` in the skill):

For each step compute four series and plot them as a **stacked** column:

| Series | Formula | Fill |
|---|---|---|
| **Base** | `MIN(prev_cumulative, cumulative)` | **none (invisible)** — floats the bar |
| **Up** | `MAX(delta, 0)` | green |
| **Down** | `MAX(-delta, 0)` | red |
| **Total** | the anchor value (opening/closing) | navy |

with `cumulative = prev_cumulative + delta`. "Total" bars (Opening, Closing,
Revenue, Net earnings) set base 0 and total = value so they rise from the axis;
step bars set total 0 and float on the invisible base. Because the base series is
`noFill`, the coloured segment appears to hover between the running totals — a
waterfall. The `MIN`/`MAX` formulas make it work for **both** positive and
negative steps (financing can be an inflow or an outflow), and because they're
live formulas the waterfall updates with the model.

The skill draws two:
- **Cash-flow waterfall:** Opening cash → Operating → Investing → Financing →
  Closing cash (the final year's cash bridge).
- **Profit waterfall:** Revenue → −COGS → −Salaries → −Rent → −D&A → −Interest →
  −Tax → Net earnings (the P&L bridge).

---

## 5. Sparklines

Sparklines (word-sized in-cell trend lines) are excellent next to KPI tiles or on
a line-item table. `openpyxl` has no first-class sparkline API, so if you add
them, either (a) inject the `x14:sparklineGroups` extension XML into the saved
`.xlsx`, or (b) approximate with a very small line chart per row. Because they are
cosmetic (no cell value), they never affect validation. The skill's KPI tiles
carry a numeric sub-metric (CAGR, margin) instead; sparklines are a documented
extension.

---

## 6. Alignment, sizing and "empty areas"

A dashboard reads as *one system* only when its objects line up. Rules the skill
follows and you should preserve when editing:

- **Snap charts to the cell grid.** Anchor every chart at a cell (`A1`-style
  anchor) and give charts a **consistent width and height** (the skill uses a
  uniform `~15.5cm × 7.4cm`), so a 2-across grid of charts has aligned edges.
  Hold **Alt** while dragging in Excel to snap to gridlines.
- **A consistent gutter.** Equal spacing between tiles and between charts. White
  space is not wasted space — it groups related content and gives the eye rest.
  Deliberate **empty rows/columns** separate the KPI band from the chart band
  from the (hidden) data band.
- **Hide the plumbing.** The chart-data block and waterfall helper tables are
  linked-in helper rows; the skill puts them below the visuals and **groups +
  hides** them (row outline level 1) so the dashboard shows only the story, not
  the scaffolding. Group, don't delete — the charts still read from them.
- **Gridlines off, one zoom.** Turn worksheet gridlines off on presentation
  sheets; the cards and chart frames provide the structure.
- **Chart formatting:** thin or no axis lines, muted gridlines, no chart border,
  title left-aligned in the model's accent navy, legend only when >1 series and
  not otherwise labelled.

---

## 7. How the dashboard links to the model (prerequisite wiring)

The dashboard is an **output** sheet: it must only *read* from the model, never
feed back into it (one-directional flow — see `BEST_PRACTICES.md` §Architecture).
Concretely:

- KPI tiles and the chart-data block are `='Three Statement Model'!…` links
  (shown in **green**, the cross-sheet-link colour).
- The chart-data block restates the model rows the charts need *locally* on the
  Dashboard, so each chart's source range lives on the same sheet as the chart —
  more robust than a chart reaching across sheets, and easy to re-point.
- Waterfall helper tables derive their steps from the model's final-year cells
  via the same green links, then compute base/up/down with local `MIN`/`MAX`.

`scripts/validate_workbook.py` recalculates the workbook and asserts every KPI
tile equals the value an independent simulation of the model produces, and that
no dashboard cell evaluates to an Excel error — so the "live link" claim is
verified, not assumed.
