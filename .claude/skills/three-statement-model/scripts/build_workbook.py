#!/usr/bin/env python3
"""
Full best-practice workbook builder for the three-statement model.

Wraps the validated single-sheet model (scripts/build_model.py) in the
supporting sheets a professionally built model has, following recognised
modelling standards (FAST / ICAEW / SMART): a Cover + contents + colour
legend, a Checks/audit sheet that rolls every integrity tie-out into one
status light, an executive Dashboard (KPI tiles + advanced charts incl.
live cash-flow and profit waterfalls), and a Sensitivity sheet (scenario
summary, two-way heat-mapped tables, and a tornado of driver impacts).

Design constraints honoured here:
  * The "Three Statement Model" sheet is written by the SAME
    populate_model_sheet() the standalone builder uses, so it stays
    byte-for-byte identical and passes scripts/validate_model.py unchanged.
  * Everything is written in a SINGLE workbook and saved ONCE. openpyxl
    silently drops charts when it loads and re-saves a file, so an
    "enhance an existing .xlsx" approach would destroy the model's charts;
    we build all sheets before the first save instead.
  * Checks and Dashboard cells are LIVE cross-sheet formulas -> they update
    if the user edits the model's blue inputs. Sensitivity grids are
    computed deterministically at build time (openpyxl cannot emit native
    Excel What-If Data Tables); the sheet documents the live-data-table and
    CHOOSE-scenario-switch recipes for users who want them live.

Usage:
  python3 build_workbook.py --inputs inputs.json [--jurisdiction pack.json]
                            [--output model.xlsx] [--prepared-by NAME]

Exit code 0 on success; non-zero with a message if input validation fails.
Validate the result with scripts/validate_workbook.py.
"""
import argparse
import copy
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.marker import DataPoint
from openpyxl.drawing.line import LineProperties

from build_model import (populate_model_sheet, validate_inputs,
                         apply_pack_defaults, resolve_labels, NUM_FMT, PCT_FMT)
from validate_model import simulate

MODEL = "Three Statement Model"

# ---- palette (institutional convention; see references/BEST_PRACTICES.md) ----
NAVY = "FF000C3F"        # forecast band / totals
ACCENT = "FF3271D2"      # section headers
INPUT_BLUE = "FF0000FF"  # hardcoded inputs
LINK_GREEN = "FF008000"  # links to other sheets
GREY = "FF5B5B5B"
WHITE = "FFFFFFFF"
LIGHT = "FFE7F2FF"       # light-blue fill (historical band)
CARD = "FFF3F8FE"        # KPI card fill
RULE = "FFBFD3EE"        # hairline separators
OK_FILL, OK_FONT = "FFC6EFCE", "FF006100"
ERR_FILL, ERR_FONT = "FFFFC7CE", "FF9C0006"
WARN_FILL, WARN_FONT = "FFFFEB9C", "FF9C6500"
RISE = "FF2E7D32"        # waterfall increase (green)
FALL = "FFC0392B"        # waterfall decrease (red)
TOTAL = "FF1F3864"       # waterfall total (navy)

# ---- fonts ----
F_TITLE = Font(name="Open Sans", size=20, bold=True, color=NAVY)
F_SUB = Font(name="Open Sans", size=11, color=ACCENT)
F_H2 = Font(name="Open Sans", size=13, bold=True, color=ACCENT)
F_BODY = Font(name="Arial Narrow", size=12, color="FF000000")
F_BODY_B = Font(name="Arial Narrow", size=12, bold=True, color="FF000000")
F_LINK = Font(name="Arial Narrow", size=12, color=LINK_GREEN)
F_LINK_B = Font(name="Arial Narrow", size=12, bold=True, color=LINK_GREEN)
F_INPUT = Font(name="Arial Narrow", size=12, color=INPUT_BLUE)
F_NOTE = Font(name="Arial Narrow", size=10, italic=True, color=GREY)
F_KPI_T = Font(name="Arial Narrow", size=11, color=GREY)
F_KPI_V = Font(name="Open Sans", size=16, bold=True, color=NAVY)
F_KPI_S = Font(name="Arial Narrow", size=10, color=GREY)
F_MASTER = Font(name="Open Sans", size=16, bold=True)
F_NAV = Font(name="Arial Narrow", size=12, bold=True, color=LINK_GREEN, underline="single")

MULT_FMT = '0.0"x"'
RATIO_FMT = '0.00'
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def M(cell):
    """A reference into the model sheet, e.g. M('M38') -> 'Three Statement Model'!M38."""
    return f"'{MODEL}'!{cell}"


def put(ws, coord, value=None, font=None, fill=None, align=None, fmt=None, border=None):
    c = ws[coord]
    if value is not None:
        c.value = value
    if font:
        c.font = font
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if align:
        c.alignment = align
    if fmt:
        c.number_format = fmt
    if border:
        c.border = border
    return c


def nav(ws, coord, text, target_sheet):
    """A clickable internal hyperlink to another sheet (real hyperlink, not a
    formula, so it survives every recalc engine)."""
    c = ws[coord]
    c.value = text
    c.hyperlink = Hyperlink(ref=coord, location=f"'{target_sheet}'!A1", display=text)
    c.font = F_NAV
    return c


def hairline(ws, row, c0, c1, side="bottom"):
    b = Border(**{side: Side(style="thin", color=RULE)})
    for col in range(c0, c1 + 1):
        ws.cell(row=row, column=col).border = b


def _rgb(color):
    """Strip a leading FF alpha so 'FF2E7D32' -> '2E7D32' for chart fills."""
    return color[2:] if isinstance(color, str) and color.startswith("FF") and len(color) == 8 else color


def series_fill(s, color=None, no_fill=False, line_color=None, line_width=None):
    gp = GraphicalProperties()
    if no_fill:
        gp.noFill = True
    elif color:
        gp.solidFill = _rgb(color)
    if line_color or line_width:
        gp.line = LineProperties(solidFill=(_rgb(line_color) if line_color else None),
                                 w=line_width)
    s.graphicalProperties = gp


def point_colors(series, colors):
    """Colour each data point of a series individually (waterfall step colours)."""
    for i, color in enumerate(colors):
        dp = DataPoint(idx=i)
        dp.graphicalProperties = GraphicalProperties(solidFill=_rgb(color))
        series.data_points.append(dp)


# ======================================================================
#  Deterministic economics (for scenarios / sensitivity / tornado)
# ======================================================================
def kpis_from_sim(sim, nh, nf):
    n = nh + nf
    last = n - 1
    ebitda = [sim[35][i] + sim[33][i] + sim[32][i] for i in range(n)]
    return {
        "revenue": sim[26][last],
        "ebitda": ebitda[last],
        "net_earnings": sim[38][last],
        "closing_cash": sim[82][last],
        "min_fc_cash": min(sim[82][nh:]),
        "max_abs_check": max(abs(x) for x in sim[60]),
    }


def perturbed_inputs(inputs, pack, changes):
    """Deep-copy inputs, ensure the tax line is filled, then shift forecast
    driver arrays. changes: list of (key, delta, mode) with mode in {add,mul}."""
    inp = apply_pack_defaults(copy.deepcopy(inputs), pack)
    fa = inp["forecast_assumptions"]
    for key, delta, mode in changes:
        arr = fa[key]
        for i in range(len(arr)):
            arr[i] = arr[i] + delta if mode == "add" else arr[i] * (1 + delta)
    return inp


def run(inputs, pack, changes):
    return kpis_from_sim(simulate(perturbed_inputs(inputs, pack, changes)),
                         inputs["n_historical"], inputs["n_forecast"])


SCENARIOS = [
    ("Base case", []),
    ("Upside", [("revenue_growth_pct", 0.02, "add"),
                ("cogs_pct_revenue", -0.02, "add")]),
    ("Downside", [("revenue_growth_pct", -0.02, "add"),
                  ("cogs_pct_revenue", 0.03, "add"),
                  ("interest_pct_avg_debt", 0.01, "add")]),
]
SCENARIO_NOTE = {
    "Base case": "assumptions as entered",
    "Upside": "revenue growth +2pp, COGS −2pp of revenue",
    "Downside": "revenue growth −2pp, COGS +3pp, interest +1pp",
}

GRID_STEPS = [-0.04, -0.02, 0.0, 0.02, 0.04]
INT_STEPS = [-0.02, -0.01, 0.0, 0.01, 0.02]

TORNADO_DRIVERS = [
    ("Revenue growth", "revenue_growth_pct", 0.02, "add"),
    ("COGS % of revenue", "cogs_pct_revenue", 0.02, "add"),
    ("Salaries % of revenue", "salaries_pct_revenue", 0.02, "add"),
    ("Rent & overhead", "rent_and_overhead", 0.10, "mul"),
    ("D&A % of PP&E", "da_pct_opening_ppe", 0.05, "add"),
    ("Interest % of debt", "interest_pct_avg_debt", 0.02, "add"),
    ("Tax rate", "tax_pct_ebt", 0.05, "add"),
]


def compute_sensitivity(inputs, pack):
    base = run(inputs, pack, [])
    scenarios = [(name, run(inputs, pack, ch), SCENARIO_NOTE[name])
                 for name, ch in SCENARIOS]
    grid_ne = [[run(inputs, pack, [("revenue_growth_pct", g, "add"),
                                   ("cogs_pct_revenue", c, "add")])["net_earnings"]
                for c in GRID_STEPS] for g in GRID_STEPS]
    grid_cash = [[run(inputs, pack, [("revenue_growth_pct", g, "add"),
                                     ("interest_pct_avg_debt", it, "add")])["min_fc_cash"]
                  for it in INT_STEPS] for g in GRID_STEPS]
    tornado = []
    for label, key, delta, mode in TORNADO_DRIVERS:
        lo = run(inputs, pack, [(key, -delta, mode)])["net_earnings"]
        hi = run(inputs, pack, [(key, delta, mode)])["net_earnings"]
        tornado.append((label, min(lo, hi), max(lo, hi)))
    tornado.sort(key=lambda t: t[2] - t[1], reverse=True)
    return {"base": base, "scenarios": scenarios,
            "grid_ne": grid_ne, "grid_cash": grid_cash,
            "tornado": tornado, "base_ne": base["net_earnings"]}


# ======================================================================
#  Geometry of the model sheet (mirrors build_model.build)
# ======================================================================
class Geo:
    def __init__(self, payload):
        H, F = payload["n_historical"], payload["n_forecast"]
        self.H, self.F, self.N = H, F, H + F
        self.FIRST = 4
        self.LAST = self.FIRST + H + F - 1
        self.FC1 = self.FIRST + H
        self.cols = [L(c) for c in range(self.FIRST, self.LAST + 1)]
        self.lc = L(self.LAST)                 # last (final forecast) column letter
        self.fc1 = L(self.FC1)                 # first forecast column
        self.prev_fc1 = L(self.FC1 - 1)        # last historical column
        self.y0 = payload["first_historical_year"]
        self.last_year = self.y0 + H + F - 1


# ======================================================================
#  COVER
# ======================================================================
def build_cover(wb, payload, pack, geo, prepared_by, units_label):
    ws = wb.create_sheet("Cover")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    ws.column_dimensions["A"].width = 2.5
    ws.column_dimensions["B"].width = 30
    for col in "CDEF":
        ws.column_dimensions[col].width = 22

    company = payload.get("company", "Company")
    put(ws, "B2", company, F_TITLE)
    put(ws, "B3", "Three-Statement Financial Model", F_SUB)
    ws.row_dimensions[2].height = 28
    hairline(ws, 4, 2, 6)

    country = (pack or {}).get("country", "—")
    framework = (pack or {}).get("accounting_framework", "IFRS / generic")
    if len(framework) > 90:
        framework = framework[:87] + "…"
    info = [
        ("Reporting entity", company),
        ("Jurisdiction", country),
        ("Currency & units", payload.get("units", units_label)),
        ("Accounting framework", framework),
        ("Historical period", f"{geo.y0}–{geo.y0 + geo.H - 1} ({geo.H} yrs, actual)"),
        ("Forecast period", f"{geo.y0 + geo.H}–{geo.last_year} ({geo.F} yrs, projected)"),
        ("Model type", "Annual, single-entity, going-concern operating model"),
        ("Prepared by", prepared_by or "—"),
        ("Version", "1.0"),
    ]
    r = 6
    for k, v in info:
        put(ws, f"B{r}", k, F_BODY_B)
        put(ws, f"C{r}", v, F_BODY, align=LEFT)
        ws.merge_cells(f"C{r}:F{r}")
        r += 1

    # ---- overall status light (pulled live from the Checks sheet) ----
    r += 1
    put(ws, f"B{r}", "Model integrity", F_H2)
    r += 1
    put(ws, f"B{r}", "Balance & tie-out checks", F_BODY_B)
    sc = put(ws, f"C{r}", "='Checks'!C2", F_BODY_B, align=CENTER)
    ws.conditional_formatting.add(
        f"C{r}", CellIsRule(operator="equal", formula=['"OK"'],
                            fill=PatternFill("solid", fgColor=OK_FILL),
                            font=Font(color=OK_FONT, bold=True)))
    ws.conditional_formatting.add(
        f"C{r}", CellIsRule(operator="equal", formula=['"ERROR"'],
                            fill=PatternFill("solid", fgColor=ERR_FILL),
                            font=Font(color=ERR_FONT, bold=True)))
    r += 1
    put(ws, f"B{r}", "Plausibility alerts", F_BODY_B)
    put(ws, f"C{r}", "='Checks'!C3", F_BODY_B, align=CENTER)
    ws.conditional_formatting.add(
        f"C{r}", CellIsRule(operator="equal", formula=['"CLEAR"'],
                            fill=PatternFill("solid", fgColor=OK_FILL),
                            font=Font(color=OK_FONT, bold=True)))
    ws.conditional_formatting.add(
        f"C{r}", CellIsRule(operator="equal", formula=['"REVIEW"'],
                            fill=PatternFill("solid", fgColor=WARN_FILL),
                            font=Font(color=WARN_FONT, bold=True)))

    # ---- contents / navigation ----
    r += 2
    put(ws, f"B{r}", "Contents", F_H2)
    r += 1
    contents = [
        ("Three Statement Model", "Income statement, balance sheet, cash flow + schedules"),
        ("Checks", "Every integrity tie-out and plausibility alert, one status light"),
        ("Dashboard", "KPI tiles and charts (revenue, margins, cash-flow & profit waterfalls)"),
        ("Sensitivity", "Scenario summary, two-way tables, tornado of driver impacts"),
    ]
    for sheet, desc in contents:
        nav(ws, f"B{r}", sheet, sheet)
        put(ws, f"C{r}", desc, F_NOTE, align=LEFT)
        ws.merge_cells(f"C{r}:F{r}")
        r += 1

    # ---- colour legend ----
    r += 1
    put(ws, f"B{r}", "How to read this model", F_H2)
    r += 1
    legend = [
        ("1,234", F_INPUT, "Blue = a hardcoded input/assumption you can change"),
        ("=A1+B1", F_BODY, "Black = a formula calculated on the same sheet"),
        ("='Sheet'!A1", F_LINK, "Green = a link that pulls from another sheet"),
    ]
    for sample, fnt, meaning in legend:
        cell = put(ws, f"B{r}", sample, fnt, align=LEFT)
        # These are illustrative TEXT samples, not live formulas — force the
        # string type so a leading '=' isn't evaluated (='Sheet'!A1 -> #REF!).
        if isinstance(sample, str) and sample.startswith("="):
            cell.data_type = "s"
        put(ws, f"C{r}", meaning, F_BODY, align=LEFT)
        ws.merge_cells(f"C{r}:F{r}")
        r += 1
    r += 1
    put(ws, f"B{r}", "Signs & units", F_BODY_B)
    put(ws, f"C{r}", "Expenses shown positive and subtracted in subtotals; capex shown "
        "positive and subtracted in the net cash row; issuance +, repayment −. "
        f"All values in {payload.get('units', units_label)} unless noted.",
        F_NOTE, align=WRAP)
    ws.merge_cells(f"C{r}:F{r}")
    ws.row_dimensions[r].height = 42

    # ---- disclaimer ----
    r += 2
    put(ws, f"B{r}", "Basis & disclaimer", F_H2)
    r += 1
    disc = (pack or {}).get(
        "disclosure",
        "Figures are model-level projections built from the stated assumptions; "
        "they are not audited financial statements, forecasts of record, or tax advice.")
    disc += (" This workbook was generated by the three-statement-model skill; "
             "review every assumption before relying on it. Not tax, audit, or "
             "investment advice.")
    put(ws, f"B{r}", disc, F_NOTE, align=WRAP)
    ws.merge_cells(f"B{r}:F{r + 3}")
    for rr in range(r, r + 4):
        ws.row_dimensions[rr].height = 24
    return ws


# ======================================================================
#  CHECKS  (audit / integrity dashboard)
# ======================================================================
# Each integrity check: (name, basis, expr(col, prev) -> abs-difference formula
# body, scope). scope 'all' = every column, 'from2' = needs a prior column.
INTEGRITY = [
    ("Balance sheet balances", "Total L&E − Total Assets = 0",
     lambda c, p: f"ABS({M(c+'60')})", "all"),
    ("BS cash ties to CFS", "Balance-sheet cash = CFS closing cash",
     lambda c, p: f"ABS({M(c+'44')}-{M(c+'82')})", "all"),
    ("BS PP&E ties to schedule", "Balance-sheet PP&E = depreciation-schedule close",
     lambda c, p: f"ABS({M(c+'47')}-{M(c+'99')})", "all"),
    ("BS debt ties to schedule", "Balance-sheet debt = debt-schedule close",
     lambda c, p: f"ABS({M(c+'52')}-{M(c+'104')})", "all"),
    ("D&A ties (IS = schedule)", "IS D&A = depreciation-schedule depreciation",
     lambda c, p: f"ABS({M(c+'32')}-{M(c+'98')})", "all"),
    ("D&A ties (IS = CFS)", "IS D&A = CFS add-back",
     lambda c, p: f"ABS({M(c+'32')}-{M(c+'67')})", "all"),
    ("Interest ties (IS = schedule)", "IS interest = debt-schedule interest",
     lambda c, p: f"ABS({M(c+'33')}-{M(c+'105')})", "all"),
    ("Total assets = total L&E", "Independent tie of the two BS totals",
     lambda c, p: f"ABS({M(c+'48')}-{M(c+'58')})", "all"),
    ("NWC composition", "NWC = AR + Inventory − AP",
     lambda c, p: f"ABS({M(c+'92')}-({M(c+'45')}+{M(c+'46')}-{M(c+'51')}))", "all"),
    ("Cash corkscrew", "Closing cash = opening + net change",
     lambda c, p: f"ABS({M(c+'82')}-({M(c+'81')}+{M(c+'80')}))", "all"),
    ("PP&E corkscrew", "Closing = opening + capex − depreciation",
     lambda c, p: f"ABS({M(c+'99')}-({M(c+'96')}+{M(c+'97')}-{M(c+'98')}))", "all"),
    ("Debt corkscrew", "Closing = opening + issuance/(repayment)",
     lambda c, p: f"ABS({M(c+'104')}-({M(c+'102')}+{M(c+'103')}))", "all"),
    ("Retained-earnings roll", "RE = prior RE + net earnings",
     lambda c, p: f"ABS({M(c+'56')}-({M(p+'56')}+{M(c+'38')}))", "from2"),
    ("Equity capital roll", "Equity = prior equity + issuance",
     lambda c, p: f"ABS({M(c+'55')}-({M(p+'55')}+{M(c+'77')}))", "from2"),
    ("Shareholders' equity sums", "SE = equity capital + retained earnings",
     lambda c, p: f"ABS({M(c+'57')}-({M(c+'55')}+{M(c+'56')}))", "all"),
]

# Plausibility alerts (informational, amber): expr -> 1 when the year is flagged.
ALERTS = [
    ("Negative closing cash", "No revolver in this template: a negative cash "
     "balance means an unfunded financing gap to size externally",
     lambda c: f"IF({M(c+'82')}<0,1,0)"),
    ("Negative shareholders' equity", "Book insolvency",
     lambda c: f"IF({M(c+'57')}<0,1,0)"),
    ("Gross margin outside 0–100%", "Check COGS mapping / assumptions",
     lambda c: f'IFERROR(IF(OR({M(c+"28")}/{M(c+"26")}<0,{M(c+"28")}/{M(c+"26")}>1),1,0),1)'),
    ("Loss-making year", "Net earnings below zero (not an error, note it)",
     lambda c: f"IF({M(c+'38')}<0,1,0)"),
]

CHK_TOL = 1.0  # $-units tolerance, matching the model's own balance flag


def build_checks(wb, payload, geo):
    ws = wb.create_sheet("Checks")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "FF2E7D32"
    ws.freeze_panes = "A5"
    ncol = geo.N
    LBL, BASIS = 2, 3                     # columns B, C
    Y0COL = 4                             # first year column = D (aligns with model)
    MAXC = Y0COL + ncol                   # 'Max error' column
    STAT = MAXC + 1                       # 'Status' column
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 44
    for i in range(ncol):
        ws.column_dimensions[L(Y0COL + i)].width = 10
    ws.column_dimensions[L(MAXC)].width = 12
    ws.column_dimensions[L(STAT)].width = 10

    put(ws, "B1", f"{payload.get('company','')} — Model Checks", F_TITLE)
    put(ws, "B2", "Integrity status", F_H2)
    master = put(ws, "C2", None, F_MASTER, align=CENTER)
    put(ws, "B3", "Plausibility alerts", F_H2)
    alert_master = put(ws, "C3", None, F_MASTER, align=CENTER)
    put(ws, f"B4", "Every check is a live formula reading the model; green = pass.",
        F_NOTE)

    # header row for the grid
    hdr = 6
    put(ws, f"{L(LBL)}{hdr}", "Integrity check", F_BODY_B)
    put(ws, f"{L(BASIS)}{hdr}", "Basis (should be 0)", F_BODY_B)
    for i in range(ncol):
        c = ws.cell(row=hdr, column=Y0COL + i, value=f"={M(geo.cols[i] + '2')}")
        c.font = F_LINK_B
        c.alignment = CENTER
    put(ws, f"{L(MAXC)}{hdr}", "Max error", F_BODY_B, align=RIGHT)
    put(ws, f"{L(STAT)}{hdr}", "Status", F_BODY_B, align=CENTER)
    hairline(ws, hdr, LBL, STAT)

    integ_status_cells = []
    row = hdr + 1
    for name, basis, expr, scope in INTEGRITY:
        put(ws, f"{L(LBL)}{row}", name, F_BODY)
        put(ws, f"{L(BASIS)}{row}", basis, F_NOTE, align=LEFT)
        for i in range(ncol):
            col = geo.cols[i]
            prev = geo.cols[i - 1] if i > 0 else None
            if scope == "from2" and i == 0:
                ws.cell(row=row, column=Y0COL + i, value="—").alignment = CENTER
                continue
            cell = ws.cell(row=row, column=Y0COL + i, value=f"={expr(col, prev)}")
            cell.font = F_BODY
            cell.number_format = '0.00;[Red]-0.00'
            cell.alignment = RIGHT
        rng = f"{L(Y0COL)}{row}:{L(Y0COL + ncol - 1)}{row}"
        mx = ws.cell(row=row, column=MAXC,
                     value=f"=MAX({L(Y0COL)}{row}:{L(Y0COL + ncol - 1)}{row})")
        mx.number_format = '0.00'
        mx.alignment = RIGHT
        st = ws.cell(row=row, column=STAT,
                     value=f'=IF({L(MAXC)}{row}>{CHK_TOL},"ERROR","OK")')
        st.alignment = CENTER
        st.font = F_BODY_B
        integ_status_cells.append(f"{L(STAT)}{row}")
        row += 1

    integ_first, integ_last = hdr + 1, row - 1
    _status_cf(ws, f"{L(STAT)}{integ_first}:{L(STAT)}{integ_last}",
               [("OK", OK_FILL, OK_FONT), ("ERROR", ERR_FILL, ERR_FONT)])

    # ---- alerts block ----
    row += 1
    ahdr = row
    put(ws, f"{L(LBL)}{ahdr}", "Plausibility alert", F_BODY_B)
    put(ws, f"{L(BASIS)}{ahdr}", "Meaning", F_BODY_B)
    put(ws, f"{L(MAXC)}{ahdr}", "Years", F_BODY_B, align=RIGHT)
    put(ws, f"{L(STAT)}{ahdr}", "Flag", F_BODY_B, align=CENTER)
    hairline(ws, ahdr, LBL, STAT)
    row += 1
    alert_status_cells = []
    for name, meaning, expr in ALERTS:
        put(ws, f"{L(LBL)}{row}", name, F_BODY)
        put(ws, f"{L(BASIS)}{row}", meaning, F_NOTE, align=LEFT)
        for i in range(ncol):
            cell = ws.cell(row=row, column=Y0COL + i, value=f"={expr(geo.cols[i])}")
            cell.font = F_NOTE
            cell.alignment = CENTER
        cnt = ws.cell(row=row, column=MAXC,
                      value=f"=SUM({L(Y0COL)}{row}:{L(Y0COL + ncol - 1)}{row})")
        cnt.alignment = RIGHT
        st = ws.cell(row=row, column=STAT,
                     value=f'=IF({L(MAXC)}{row}>0,"REVIEW","CLEAR")')
        st.alignment = CENTER
        st.font = F_BODY_B
        alert_status_cells.append(f"{L(STAT)}{row}")
        row += 1
    _status_cf(ws, f"{L(STAT)}{ahdr + 1}:{L(STAT)}{row - 1}",
               [("CLEAR", OK_FILL, OK_FONT), ("REVIEW", WARN_FILL, WARN_FONT)])

    # ---- master roll-ups ----
    ir = f"{L(STAT)}{integ_first}:{L(STAT)}{integ_last}"
    ar = f"{L(STAT)}{ahdr + 1}:{L(STAT)}{row - 1}"
    master.value = f'=IF(COUNTIF({ir},"ERROR")>0,"ERROR","OK")'
    alert_master.value = f'=IF(COUNTIF({ar},"REVIEW")>0,"REVIEW","CLEAR")'
    for cell, good, bad, gf, bf in [
            (master, "OK", "ERROR", OK_FILL, ERR_FILL),
            (alert_master, "CLEAR", "REVIEW", OK_FILL, WARN_FILL)]:
        ws.conditional_formatting.add(
            cell.coordinate, CellIsRule(operator="equal", formula=[f'"{good}"'],
                                        fill=PatternFill("solid", fgColor=gf),
                                        font=Font(color=OK_FONT, bold=True, size=16)))
        ws.conditional_formatting.add(
            cell.coordinate, CellIsRule(operator="equal", formula=[f'"{bad}"'],
                                        fill=PatternFill("solid", fgColor=bf),
                                        font=Font(color=(ERR_FONT if bad == "ERROR"
                                                         else WARN_FONT),
                                                  bold=True, size=16)))
    return ws, {"master": "C2", "alert_master": "C3",
                "integrity_status": integ_status_cells,
                "alert_status": alert_status_cells}


def _status_cf(ws, rng, pairs):
    for text, fill, font in pairs:
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=[f'"{text}"'],
                            fill=PatternFill("solid", fgColor=fill),
                            font=Font(color=font, bold=True)))


# ======================================================================
#  DASHBOARD
# ======================================================================
def _card(ws, r, c, title, value_formula, fmt, sub=None, sub_fmt=None):
    """A 3-col × 3-row KPI card at (r,c): title / big value / sub, each line
    merged across the three columns so text can breathe."""
    box = Border(*[Side(style="thin", color=RULE)] * 4)
    for rr in range(r, r + 3):
        for cc in range(c, c + 3):
            cell = ws.cell(row=rr, column=cc)
            cell.fill = PatternFill("solid", fgColor=CARD)
            cell.border = box
    for rr in range(r, r + 3):
        ws.merge_cells(start_row=rr, start_column=c, end_row=rr, end_column=c + 2)
    tc = ws.cell(row=r, column=c, value=title)
    tc.font = F_KPI_T
    tc.alignment = Alignment(horizontal="left", vertical="top")
    vc = ws.cell(row=r + 1, column=c, value=value_formula)
    vc.font = F_KPI_V
    vc.number_format = fmt
    vc.alignment = Alignment(horizontal="left", vertical="center")
    if sub is not None:
        sc = ws.cell(row=r + 2, column=c, value=sub)
        sc.font = F_KPI_S
        sc.number_format = sub_fmt or "General"
        sc.alignment = Alignment(horizontal="left", vertical="bottom")
    return vc


def build_dashboard(wb, payload, geo, units_label):
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = ACCENT
    ws.sheet_view.zoomScale = 100
    for i in range(1, 22):
        ws.column_dimensions[L(i)].width = 10.5
    ws.column_dimensions["A"].width = 2.5

    lc = geo.lc
    put(ws, "B2", f"{payload.get('company','')} — Executive Dashboard", F_TITLE)
    put(ws, "B3", f"Final forecast year {geo.last_year} · all values in "
        f"{payload.get('units', units_label)} unless labelled", F_SUB)

    def m(row):
        return M(lc + str(row))

    ebitda = f"({m(35)}+{m(33)}+{m(32)})"
    cards = [
        ("Revenue", f"={m(26)}", NUM_FMT,
         f"=IFERROR((({m(26)})/({M(geo.prev_fc1 + '26')}))^(1/{geo.F})-1,\"n/a\")",
         "\"CAGR \"0.0%"),
        ("Gross margin", f"=IFERROR({m(28)}/{m(26)},\"n/a\")", "0.0%", None, None),
        ("EBITDA", f"={ebitda}", NUM_FMT,
         f"=IFERROR({ebitda}/{m(26)},\"n/a\")", "\"margin \"0.0%"),
        ("Net earnings", f"={m(38)}", NUM_FMT,
         f"=IFERROR({m(38)}/{m(26)},\"n/a\")", "\"margin \"0.0%"),
        ("Closing cash", f"={m(82)}", NUM_FMT, None, None),
        ("Total assets", f"={m(48)}", NUM_FMT, None, None),
        ("Return on equity", f"=IFERROR({m(38)}/{m(57)},\"n/a\")", "0.0%", None, None),
        ("Current ratio", f"=IFERROR(({m(44)}+{m(45)}+{m(46)})/{m(51)},\"n/a\")",
         MULT_FMT, None, None),
        ("Net debt / EBITDA", f"=IFERROR(({m(52)}-{m(44)})/{ebitda},\"n/a\")",
         MULT_FMT, None, None),
    ]
    kpi_cells = {}
    r0, c0, per = 5, 2, 3
    for i, (title, vf, fmt, sub, sfmt) in enumerate(cards):
        rr = r0 + (i // per) * 4
        cc = c0 + (i % per) * 3
        cell = _card(ws, rr, cc, title, vf, fmt, sub, sfmt)
        kpi_cells[title] = cell.coordinate
    for rr in range(r0, r0 + ((len(cards) - 1) // per + 1) * 4):
        ws.row_dimensions[rr].height = 18

    # ---- chart data block (live links into the model), grouped/hidden ----
    data_top = 40
    put(ws, f"B{data_top - 1}", "Chart data (linked to the model)", F_NOTE)
    year_row = data_top
    put(ws, f"A{year_row}", "Year", F_NOTE)
    for i, col in enumerate(geo.cols):
        cell = ws.cell(row=year_row, column=2 + i, value=f"={M(col + '2')}")
        cell.font = F_LINK
        cell.number_format = "0"
    series_defs = [
        ("Revenue", lambda c: f"={M(c + '26')}", NUM_FMT),
        ("Gross margin %", lambda c: f"=IFERROR({M(c + '28')}/{M(c + '26')},0)", "0.0%"),
        ("EBITDA", lambda c: f"={M(c + '35')}+{M(c + '33')}+{M(c + '32')}", NUM_FMT),
        ("Net earnings", lambda c: f"={M(c + '38')}", NUM_FMT),
        ("Operating cash flow", lambda c: f"={M(c + '69')}", NUM_FMT),
        ("Investing cash flow", lambda c: f"=-{M(c + '73')}", NUM_FMT),
        ("Financing cash flow", lambda c: f"={M(c + '78')}", NUM_FMT),
        ("Closing cash", lambda c: f"={M(c + '82')}", NUM_FMT),
        ("Accounts receivable", lambda c: f"={M(c + '45')}", NUM_FMT),
        ("Inventory", lambda c: f"={M(c + '46')}", NUM_FMT),
        ("Accounts payable", lambda c: f"={M(c + '51')}", NUM_FMT),
        ("Net working capital", lambda c: f"={M(c + '92')}", NUM_FMT),
        ("Debt", lambda c: f"={M(c + '52')}", NUM_FMT),
        ("Equity", lambda c: f"={M(c + '57')}", NUM_FMT),
        ("Interest coverage", lambda c: f"=IFERROR(({M(c + '35')}+{M(c + '33')})/{M(c + '33')},0)", MULT_FMT),
    ]
    srow = {}
    for j, (name, fn, fmt) in enumerate(series_defs):
        rr = year_row + 1 + j
        srow[name] = rr
        put(ws, f"A{rr}", name, F_NOTE)
        for i, col in enumerate(geo.cols):
            cell = ws.cell(row=rr, column=2 + i, value=fn(col))
            cell.font = F_LINK
            cell.number_format = fmt

    cats = Reference(ws, min_col=2, max_col=1 + geo.N, min_row=year_row, max_row=year_row)

    def rowref(name):
        return Reference(ws, min_col=2, max_col=1 + geo.N,
                         min_row=srow[name], max_row=srow[name])

    # Chart 1 — Revenue & gross margin (combo)
    ch1 = BarChart()
    ch1.type = "col"; ch1.grouping = "clustered"
    ch1.title = "Revenue & gross margin"
    ch1.height, ch1.width = 7.4, 15.5
    ch1.add_data(rowref("Revenue"), titles_from_data=False, from_rows=True)
    ch1.set_categories(cats)
    ch1.series[0].tx = SeriesLabel(strRef=StrRef(f"'Dashboard'!$A${srow['Revenue']}"))
    series_fill(ch1.series[0], ACCENT)
    ln = LineChart()
    ln.add_data(rowref("Gross margin %"), titles_from_data=False, from_rows=True)
    ln.set_categories(cats)
    ln.series[0].tx = SeriesLabel(strRef=StrRef(f"'Dashboard'!$A${srow['Gross margin %']}"))
    ln.y_axis.axId = 200; ln.y_axis.numFmt = "0%"; ln.y_axis.majorGridlines = None
    ch1.y_axis.crosses = "max"
    ch1 += ln
    ws.add_chart(ch1, "B26")

    # Chart 2 — Cash-flow composition (stacked) + closing cash line
    ch2 = BarChart()
    ch2.type = "col"; ch2.grouping = "stacked"; ch2.overlap = 100
    ch2.title = "Cash flow composition"
    ch2.height, ch2.width = 7.4, 15.5
    for name, color in [("Operating cash flow", RISE),
                        ("Investing cash flow", FALL),
                        ("Financing cash flow", ACCENT)]:
        ch2.add_data(rowref(name), titles_from_data=False, from_rows=True)
        s = ch2.series[-1]
        s.tx = SeriesLabel(strRef=StrRef(f"'Dashboard'!$A${srow[name]}"))
        series_fill(s, color)
    ln2 = LineChart()
    ln2.add_data(rowref("Closing cash"), titles_from_data=False, from_rows=True)
    ln2.set_categories(cats)
    ln2.series[0].tx = SeriesLabel(strRef=StrRef(f"'Dashboard'!$A${srow['Closing cash']}"))
    series_fill(ln2.series[0], NAVY, line_color=NAVY, line_width=28000)
    ch2.set_categories(cats)
    ch2 += ln2
    ws.add_chart(ch2, "K26")

    # Chart 3 — Working capital
    ch3 = BarChart()
    ch3.type = "col"; ch3.grouping = "clustered"
    ch3.title = "Working capital"
    ch3.height, ch3.width = 7.4, 15.5
    for name, color in [("Accounts receivable", ACCENT),
                        ("Inventory", "FF7FA6DC"),
                        ("Accounts payable", FALL)]:
        ch3.add_data(rowref(name), titles_from_data=False, from_rows=True)
        s = ch3.series[-1]
        s.tx = SeriesLabel(strRef=StrRef(f"'Dashboard'!$A${srow[name]}"))
        series_fill(s, color)
    ln3 = LineChart()
    ln3.add_data(rowref("Net working capital"), titles_from_data=False, from_rows=True)
    ln3.set_categories(cats)
    ln3.series[0].tx = SeriesLabel(strRef=StrRef(f"'Dashboard'!$A${srow['Net working capital']}"))
    series_fill(ln3.series[0], NAVY, line_color=NAVY, line_width=28000)
    ch3.set_categories(cats)
    ch3 += ln3
    ws.add_chart(ch3, "B42")

    # Chart 4 — Leverage & coverage
    ch4 = BarChart()
    ch4.type = "col"; ch4.grouping = "clustered"
    ch4.title = "Capital structure & interest cover"
    ch4.height, ch4.width = 7.4, 15.5
    for name, color in [("Debt", FALL), ("Equity", ACCENT)]:
        ch4.add_data(rowref(name), titles_from_data=False, from_rows=True)
        s = ch4.series[-1]
        s.tx = SeriesLabel(strRef=StrRef(f"'Dashboard'!$A${srow[name]}"))
        series_fill(s, color)
    ln4 = LineChart()
    ln4.add_data(rowref("Interest coverage"), titles_from_data=False, from_rows=True)
    ln4.set_categories(cats)
    ln4.series[0].tx = SeriesLabel(strRef=StrRef(f"'Dashboard'!$A${srow['Interest coverage']}"))
    ln4.y_axis.axId = 210; ln4.y_axis.numFmt = '0.0"x"'
    series_fill(ln4.series[0], NAVY, line_color=NAVY, line_width=28000)
    ch4.y_axis.crosses = "max"
    ch4.set_categories(cats)
    ch4 += ln4
    ws.add_chart(ch4, "K42")

    # ---- waterfalls (live, via invisible-base stacked bars) ----
    wf_top = year_row + len(series_defs) + 3
    cf_rows = _waterfall_block(
        ws, wf_top, "Cash-flow waterfall ({})".format(geo.last_year),
        [("Opening cash", None, M(lc + "81"), None),
         ("Operating", M(lc + "69"), None, "up"),
         ("Investing", "-" + M(lc + "73"), None, "down"),
         ("Financing", M(lc + "78"), None, "up"),
         ("Closing cash", None, M(lc + "82"), None)])
    _waterfall_chart(ws, "B58", "Cash-flow waterfall ({})".format(geo.last_year), cf_rows)

    pf_top = wf_top + 10
    pf_rows = _waterfall_block(
        ws, pf_top, "Profit waterfall ({})".format(geo.last_year),
        [("Revenue", None, M(lc + "26"), None),
         ("COGS", "-" + M(lc + "27"), None, "down"),
         ("Salaries", "-" + M(lc + "30"), None, "down"),
         ("Rent & OH", "-" + M(lc + "31"), None, "down"),
         ("D&A", "-" + M(lc + "32"), None, "down"),
         ("Interest", "-" + M(lc + "33"), None, "down"),
         ("Tax", "-" + M(lc + "37"), None, "down"),
         ("Net earnings", None, M(lc + "38"), None)])
    _waterfall_chart(ws, "K58", "Profit waterfall ({})".format(geo.last_year), pf_rows)

    # group the data + waterfall helper rows so the dashboard stays clean
    for rr in range(data_top - 1, pf_rows["last_row"] + 1):
        ws.row_dimensions[rr].outlineLevel = 1
        ws.row_dimensions[rr].hidden = True
    ws.sheet_properties.outlinePr.summaryBelow = False
    return ws, kpi_cells, {"cf": cf_rows, "pf": pf_rows}


def _waterfall_block(ws, top, title, steps):
    """Write a live waterfall helper table that renders correctly even when the
    running total crosses zero (loss years / negative cash) — the plain
    invisible-base trick breaks there because a stacked column draws negative and
    positive segments independently around the axis. Each floating bar [lo, hi]
    is split at the axis into two stacks: base_above (invisible) + above (visible)
    rise from 0 to hi, base_below (invisible) + below (visible) fall from 0 to lo,
    so a bar spanning negative→positive shows as two same-coloured segments
    meeting at zero. Columns A=label B=cumulative C=lo D=hi E=base_above F=above
    G=base_below H=below. Returns row bookkeeping + per-step colours."""
    put(ws, f"A{top}", title, F_NOTE)
    hdr = top + 1
    for j, h in enumerate(["Step", "Cumul.", "Lo", "Hi", "BaseUp", "Up",
                           "BaseDn", "Dn"]):
        put(ws, f"{L(1 + j)}{hdr}", h, F_NOTE)
    first = hdr + 1
    rows, colors = [], []
    for k, (label, delta, total, direction) in enumerate(steps):
        r = first + k
        rows.append(r)
        put(ws, f"A{r}", label, F_NOTE)
        if total is not None:                        # anchored total bar (0..V)
            put(ws, f"B{r}", f"={total}", F_LINK, fmt=NUM_FMT)
            put(ws, f"C{r}", f"=MIN(0,B{r})", F_NOTE, fmt=NUM_FMT)
            put(ws, f"D{r}", f"=MAX(0,B{r})", F_NOTE, fmt=NUM_FMT)
            colors.append(TOTAL)
        else:                                        # floating step bar
            p = f"B{r - 1}"
            put(ws, f"B{r}", f"={p}+({delta})", F_NOTE, fmt=NUM_FMT)
            put(ws, f"C{r}", f"=MIN({p},B{r})", F_NOTE, fmt=NUM_FMT)
            put(ws, f"D{r}", f"=MAX({p},B{r})", F_NOTE, fmt=NUM_FMT)
            colors.append(RISE if direction == "up" else FALL)
        put(ws, f"E{r}", f"=MAX(C{r},0)", F_NOTE, fmt=NUM_FMT)              # base_above
        put(ws, f"F{r}", f"=MAX(D{r},0)-MAX(C{r},0)", F_NOTE, fmt=NUM_FMT)  # above
        put(ws, f"G{r}", f"=MIN(D{r},0)", F_NOTE, fmt=NUM_FMT)              # base_below
        put(ws, f"H{r}", f"=MIN(C{r},0)-MIN(D{r},0)", F_NOTE, fmt=NUM_FMT)  # below
    return {"first": first, "last": rows[-1], "last_row": rows[-1],
            "cat_col": 1, "base_above": 5, "above": 6, "base_below": 7,
            "below": 8, "colors": colors}


def _waterfall_chart(ws, anchor, title, rr):
    ch = BarChart()
    ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100
    ch.title = title
    ch.height, ch.width = 7.4, 15.5
    cats = Reference(ws, min_col=rr["cat_col"], max_col=rr["cat_col"],
                     min_row=rr["first"], max_row=rr["last"])
    # invisible base then visible on each side of the axis; the two visible
    # series carry per-step colours so each bar reads as one colour.
    specs = [(rr["base_above"], True), (rr["above"], False),
             (rr["base_below"], True), (rr["below"], False)]
    for col, invisible in specs:
        ref = Reference(ws, min_col=col, max_col=col, min_row=rr["first"], max_row=rr["last"])
        ch.add_data(ref, titles_from_data=False, from_rows=False)
        s = ch.series[-1]
        if invisible:
            series_fill(s, no_fill=True)
        else:
            point_colors(s, rr["colors"])
    ch.set_categories(cats)
    ch.legend = None
    ws.add_chart(ch, anchor)


# ======================================================================
#  SENSITIVITY
# ======================================================================
def build_sensitivity(wb, payload, geo, sens, units_label):
    ws = wb.create_sheet("Sensitivity")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "FF7030A0"
    ws.column_dimensions["A"].width = 2.5
    ws.column_dimensions["B"].width = 24
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 13

    put(ws, "B2", f"{payload.get('company','')} — Sensitivity & Scenarios", F_TITLE)
    put(ws, "B3", f"Outputs for the final forecast year ({geo.last_year}); "
        f"{payload.get('units', units_label)}. Computed at build time — see the "
        "live-model recipe below.", F_SUB)

    written = {"scenarios": {}, "grid_ne": [], "grid_cash": [], "tornado": []}
    placements = {}  # coord -> written literal, for the validator

    # ---- scenario summary ----
    r = 5
    put(ws, f"B{r}", "Scenario summary", F_H2)
    r += 1
    heads = ["Scenario", "Revenue", "EBITDA", "Net earnings", "Closing cash",
             "Min FC cash"]
    for j, h in enumerate(heads):
        put(ws, f"{L(2 + j)}{r}", h, F_BODY_B, align=(LEFT if j == 0 else RIGHT))
    hairline(ws, r, 2, 7)
    r += 1
    for name, k, note in sens["scenarios"]:
        put(ws, f"B{r}", name, F_BODY_B if name == "Base case" else F_BODY)
        vals = [k["revenue"], k["ebitda"], k["net_earnings"],
                k["closing_cash"], k["min_fc_cash"]]
        for j, v in enumerate(vals):
            coord = f"{L(3 + j)}{r}"
            put(ws, coord, round(v, 2), F_BODY, fmt=NUM_FMT, align=RIGHT)
            placements[coord] = round(v, 2)
        written["scenarios"][name] = vals
        r += 1
    for i, (name, _, note) in enumerate(sens["scenarios"]):
        put(ws, f"B{r + i}", f"{name}: {note}", F_NOTE)
        ws.merge_cells(f"B{r + i}:H{r + i}")
    r += len(sens["scenarios"]) + 1

    # ---- two-way grid 1: Net earnings vs growth × COGS ----
    r += 1
    put(ws, f"B{r}", f"Net earnings ({geo.last_year}) — revenue growth × COGS % of revenue",
        F_H2)
    r += 1
    written["grid_ne"] = _grid(ws, r, sens["grid_ne"], GRID_STEPS, GRID_STEPS,
                               "Growth / COGS", placements)
    r += len(GRID_STEPS) + 3

    # ---- two-way grid 2: Min FC cash vs growth × interest ----
    put(ws, f"B{r}", f"Minimum forecast cash — revenue growth × interest rate", F_H2)
    r += 1
    written["grid_cash"] = _grid(ws, r, sens["grid_cash"], GRID_STEPS, INT_STEPS,
                                 "Growth / Interest", placements)
    r += len(GRID_STEPS) + 3

    # ---- tornado ----
    put(ws, f"B{r}", f"Tornado — driver impact on net earnings ({geo.last_year})", F_H2)
    r += 1
    tb = _tornado(ws, r, sens["tornado"], sens["base_ne"], placements)
    written["tornado"] = [(lab, lo, hi) for lab, lo, hi in sens["tornado"]]
    r = tb + 2

    # ---- live-model recipe ----
    put(ws, f"B{r}", "Make it live in Excel", F_H2)
    r += 1
    recipe = (
        "These grids are computed deterministically at build time because .xlsx "
        "written programmatically cannot embed native What-If Data Tables. To make "
        "them recalculate inside Excel: (1) put the output formula (e.g. "
        "='Three Statement Model'!" + geo.lc + "38) in the top-left corner of a grid, "
        "row inputs down the left and column inputs across the top; select the block "
        "and use Data ▸ What-If Analysis ▸ Data Table, pointing Row/Column input cells "
        "at the driver cells. (2) For scenarios, add a switch cell with a "
        "Base/Upside/Downside dropdown (Data ▸ Data Validation) and drive each "
        "assumption with =CHOOSE(switch, base, up, down) on an inputs sheet, keeping "
        "the model's blue cells pointed at the active scenario column.")
    put(ws, f"B{r}", recipe, F_NOTE, align=WRAP)
    ws.merge_cells(f"B{r}:H{r + 4}")
    for rr in range(r, r + 5):
        ws.row_dimensions[rr].height = 22
    written["placements"] = placements
    return ws, written


def _grid(ws, top, grid, row_steps, col_steps, corner, placements, pct_axis=True):
    put(ws, f"B{top}", corner, F_NOTE, align=CENTER, fill=LIGHT)
    for j, cs in enumerate(col_steps):
        put(ws, f"{L(3 + j)}{top}", cs, F_BODY_B, fmt="+0.0%;-0.0%", align=CENTER,
            fill=LIGHT)
    written = []
    for i, rs in enumerate(row_steps):
        rr = top + 1 + i
        put(ws, f"B{rr}", rs, F_BODY_B, fmt="+0.0%;-0.0%", align=CENTER, fill=LIGHT)
        rowvals = []
        for j, _ in enumerate(col_steps):
            v = round(grid[i][j], 2)
            centre = (row_steps[i] == 0.0 and col_steps[j] == 0.0)
            coord = f"{L(3 + j)}{rr}"
            put(ws, coord, v, F_BODY_B if centre else F_BODY, fmt=NUM_FMT, align=RIGHT)
            placements[coord] = v
            rowvals.append(v)
        written.append(rowvals)
    rng = f"{L(3)}{top + 1}:{L(2 + len(col_steps))}{top + len(row_steps)}"
    ws.conditional_formatting.add(rng, ColorScaleRule(
        start_type="min", start_color="FFF8696B",
        mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
        end_type="max", end_color="FF63BE7B"))
    return written


def _tornado(ws, top, tornado, base, placements):
    # helper table: label | low-base (neg) | high-base (pos).
    # A horizontal bar chart renders category 0 at the BOTTOM, so write the
    # widest-swing driver LAST to put it on top (the classic tornado funnel).
    tornado = list(reversed(tornado))
    put(ws, f"A{top}", "Driver", F_NOTE)
    put(ws, f"B{top}", "Down", F_NOTE)
    put(ws, f"C{top}", "Up", F_NOTE)
    first = top + 1
    for k, (label, lo, hi) in enumerate(tornado):
        rr = first + k
        put(ws, f"A{rr}", label, F_NOTE)
        put(ws, f"B{rr}", round(lo - base, 2), F_NOTE, fmt=NUM_FMT)
        put(ws, f"C{rr}", round(hi - base, 2), F_NOTE, fmt=NUM_FMT)
        placements[f"B{rr}"] = round(lo - base, 2)
        placements[f"C{rr}"] = round(hi - base, 2)
    last = first + len(tornado) - 1
    ch = BarChart()
    ch.type = "bar"; ch.grouping = "stacked"; ch.overlap = 100
    ch.title = f"Net-earnings swing vs base ({round(base):,})"
    ch.height, ch.width = 8.0, 16.0
    cats = Reference(ws, min_col=1, max_col=1, min_row=first, max_row=last)
    for col, color in [(2, FALL), (3, RISE)]:
        ref = Reference(ws, min_col=col, max_col=col, min_row=first, max_row=last)
        ch.add_data(ref, titles_from_data=False, from_rows=False)
        series_fill(ch.series[-1], color)
    ch.set_categories(cats)
    ch.legend = None
    ws.add_chart(ch, f"E{top}")
    # group helper rows
    for rr in range(top, last + 1):
        ws.row_dimensions[rr].outlineLevel = 1
        ws.row_dimensions[rr].hidden = True
    return last + 14  # leave room under the chart


# ======================================================================
#  Model-sheet polish (safe, non-structural) + named ranges + print
# ======================================================================
def polish_model_sheet(ws, geo):
    # collapse the supporting-schedule and chart-feed blocks
    for rng in ((87, 106), (110, 116)):
        for rr in range(rng[0], rng[1] + 1):
            ws.row_dimensions[rr].outlineLevel = 1
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.print_options.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr = None
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = None
    ws.print_area = f"A1:{geo.lc}116"


def add_named_ranges(wb, geo):
    lc = geo.lc
    names = {
        "FinalYearRevenue": f"'{MODEL}'!${lc}$26",
        "FinalYearNetEarnings": f"'{MODEL}'!${lc}$38",
        "FinalYearClosingCash": f"'{MODEL}'!${lc}$82",
        "ModelBalanceChecks": f"'{MODEL}'!$D$3:${lc}$3",
    }
    for name, ref in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)


def setup_prints(wb):
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = None
        ws.oddFooter.left.text = "&F"
        ws.oddFooter.center.text = "&A"
        ws.oddFooter.right.text = "Page &P of &N"
        ws.oddHeader.right.text = "&D"


# ======================================================================
#  Orchestration
# ======================================================================
def build_full(payload, pack, out_path, prepared_by=None):
    geo = Geo(payload)
    units_label = (payload.get("units_label")
                   or (pack or {}).get("units_label", "$000's"))

    wb = Workbook()
    # model sheet first so charts survive the single save
    model_ws = wb.active
    model_ws.title = MODEL
    wb.properties.title = payload.get("company", MODEL)
    wb.properties.creator = "three-statement-model skill"
    populate_model_sheet(model_ws, payload, pack)
    polish_model_sheet(model_ws, geo)
    model_ws.sheet_properties.tabColor = ACCENT

    sens = compute_sensitivity(payload, pack)

    cover = build_cover(wb, payload, pack, geo, prepared_by, units_label)
    _, checks_map = build_checks(wb, payload, geo)
    _, kpi_cells, wf = build_dashboard(wb, payload, geo, units_label)
    _, sens_written = build_sensitivity(wb, payload, geo, sens, units_label)

    add_named_ranges(wb, geo)
    setup_prints(wb)

    # sheet order: Cover, Model, Checks, Dashboard, Sensitivity
    order = ["Cover", MODEL, "Checks", "Dashboard", "Sensitivity"]
    wb._sheets.sort(key=lambda s: order.index(s.title))
    wb.active = 0
    wb.save(out_path)
    return {"geo": geo, "kpi_cells": kpi_cells, "checks_map": checks_map,
            "waterfalls": wf, "sensitivity": sens_written, "sens_values": sens}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--inputs", required=True)
    ap.add_argument("--jurisdiction", help="jurisdiction pack JSON (optional)")
    ap.add_argument("--output", default="model.xlsx")
    ap.add_argument("--prepared-by", default=None)
    args = ap.parse_args()

    payload = json.load(open(args.inputs))
    pack = json.load(open(args.jurisdiction)) if args.jurisdiction else None
    payload = apply_pack_defaults(payload, pack)

    errors, warnings = validate_inputs(payload)
    for w in warnings:
        print("  ⚠", w)
    if errors:
        print("INPUT VALIDATION FAILED:")
        for e in errors:
            print("  ✗", e)
        sys.exit(2)

    build_full(payload, pack, args.output, args.prepared_by)
    print(f"OK: {args.output} — Cover · {MODEL} · Checks · Dashboard · Sensitivity"
          + (f" (jurisdiction: {pack.get('country')})" if pack else ""))


if __name__ == "__main__":
    main()
