#!/usr/bin/env python3
"""
Acceptance-test harness for single-sheet 3-statement models built to the reference template.

Validates a generated .xlsx against the specification in
references/TEMPLATE_SPEC.md, in three layers:

  1. STRUCTURE  - sheet name, row labels, year headers, sections.
  2. FORMULAS   - forecast block must be formulas (no hardcoded plugs),
                  canonical formula patterns on key rows, and column-to-column
                  consistency (each forecast formula must be the translation of
                  the previous column's formula).
  3. VALUES     - the workbook is recalculated (LibreOffice headless, cached
                  values, or the `formulas` package - whichever is available)
                  and every computed line is compared against an independent
                  pure-Python simulation of the model economics driven only by
                  the inputs JSON. Balance-sheet tie-out and cash tie-out are
                  asserted for every column.

Usage:
  python3 validate_model.py MODEL.xlsx --inputs inputs.json
        [--jurisdiction pack.json] [--reference reference_values.json]
        [--tolerance 0.01] [--sheet NAME]

With --jurisdiction, the pack's units_label and label_overrides are applied to
the expected row labels before the structure check, mirroring build_model.py.

Exit code 0 = all checks pass, 1 = failures (report printed to stdout).
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

import openpyxl
from openpyxl.utils import get_column_letter

# Column lists are derived from the inputs JSON in main(); these module-level
# defaults cover the canonical 5-historical + 5-forecast case.
COLS_ALL = list("DEFGHIJKLM")
COLS_HIST = list("DEFGH")
COLS_FCST = list("IJKLM")


def set_columns(n_hist, n_fcst):
    """Derive data-column letters: data column k (1-based) = sheet column 3+k."""
    global COLS_ALL, COLS_HIST, COLS_FCST
    COLS_ALL = [get_column_letter(4 + i) for i in range(n_hist + n_fcst)]
    COLS_HIST = COLS_ALL[:n_hist]
    COLS_FCST = COLS_ALL[n_hist:]

# Row map: row -> (label, section). Kept in sync with the prompt's cell map.
ROW_LABELS = {
    3: "Balance Sheet Check",
    5: "Assumptions", 7: "Key Assumptions",
    8: "Revenue Growth (% Change)",
    9: "Cost of Goods Sold (% of Revenue)",
    10: "Salaries and Benefits (% of Revenue)",
    11: "Rent and Overhead ($000's)",
    12: "Depreciation & Amortization (% of PP&E)",
    13: "Interest (% of Debt)",
    14: "Tax Rate (% of Earnings Before Tax)",
    15: "Accounts Receivable (Days)",
    16: "Inventory (Days)",
    17: "Accounts Payable (Days)",
    18: "Capital Expenditures ($000's)",
    19: "Debt Issuance (Repayment) ($000's)",
    20: "Equity Issued (Repaid) ($000's)",
    21: "Days in Period",
    24: "Income Statement",
    26: "Revenue", 27: "Cost of Goods Sold (COGS)", 28: "Gross Profit",
    29: "Expenses", 30: "Salaries and Benefits", 31: "Rent and Overhead",
    32: "Depreciation & Amortization", 33: "Interest", 34: "Total Expenses",
    35: "Earnings Before Tax", 37: "Taxes", 38: "Net Earnings",
    41: "Balance Sheet", 43: "Assets",
    44: "Cash", 45: "Accounts Receivable", 46: "Inventory",
    47: "Property & Equipment", 48: "Total Assets",
    50: "Liabilities", 51: "Accounts Payable", 52: "Debt",
    53: "Total Liabilities", 54: "Shareholder's Equity",
    55: "Equity Capital", 56: "Retained Earnings",
    57: "Shareholder's Equity",
    58: "Total Liabilities & Shareholder's Equity", 60: "Check",
    63: "Cash Flow Statement", 65: "Operating Cash Flow",
    66: "Net Earnings", 67: "Plus: Depreciation & Amortization",
    68: "Less: Changes in Working Capital", 69: "Cash from Operations",
    71: "Investing Cash Flow", 72: "Investments in Property & Equipment",
    73: "Cash from Investing", 75: "Financing Cash Flow",
    76: "Issuance (repayment) of debt", 77: "Issuance (repayment) of equity",
    78: "Cash from Financing", 80: "Net Increase (decrease) in Cash",
    81: "Opening Cash Balance", 82: "Closing Cash Balance",
    86: "Supporting Schedules", 88: "Working Capital Schedule",
    89: "Accounts Receivable", 90: "Inventory", 91: "Accounts Payable",
    92: "Net Working Capital (NWC)", 93: "Change in NWC",
    95: "Depreciation Schedule", 96: "PPE Opening", 97: "Plus Capex",
    98: "Less Depreciation", 99: "PPE Closing",
    101: "Debt & Interest Schedule", 102: "Debt Opening",
    103: "Issuance (repayment)", 104: "Debt Closing", 105: "Interest Expense",
}

# Canonical formulas for the FIRST FORECAST column (I). {c}=I, {p}=H (prior col).
CANONICAL_FCST = {
    26: "={p}26*(1+{c}8)",
    27: "={c}26*{c}9",
    28: "={c}26-{c}27",
    30: "={c}10*{c}26",
    31: "={c}11",
    32: "={c}67",
    33: "={c}105",
    34: "=SUM({c}30:{c}33)",
    35: "={c}28-{c}34",
    37: "={c}35*{c}14",
    38: "={c}35-{c}37",
    44: "={c}82",
    45: "={c}26*{c}15/{c}21",
    46: "={c}27*{c}16/{c}21",
    47: "={c}99",
    48: "=SUM({c}44:{c}47)",
    51: "={c}27*{c}17/{c}21",
    52: "={c}104",
    53: "=SUM({c}51:{c}52)",
    55: "={p}55+{c}77",
    56: "={p}56+{c}38",
    57: "=SUM({c}55:{c}56)",
    58: "={c}53+{c}57",
    60: "={c}58-{c}48",
    66: "={c}38",
    67: "={c}98",
    68: "={c}93",
    69: "={c}66+{c}67-{c}68",
    72: "={c}97",
    73: "=SUM({c}72)",
    76: "={c}103",
    77: "={c}20",
    78: "=SUM({c}76:{c}77)",
    80: "={c}69-{c}73+{c}78",
    81: "={p}82",
    82: "=SUM({c}80:{c}81)",
    89: "={c}45",
    90: "={c}46",
    91: "={c}51",
    92: "={c}89+{c}90-{c}91",
    93: "={c}92-{p}92",
    96: "={p}99",
    97: "={c}18",
    98: "={c}96*{c}12",
    99: "={c}96+{c}97-{c}98",
    102: "={p}104",
    103: "={c}19",
    104: "={c}102+{c}103",
    105: "=AVERAGE({c}102,{c}104)*{c}13",
}

# Canonical formulas for HISTORICAL columns from the SECOND one onward
# ({c}=current, {p}=prior). The first historical column D differs: row 8 blank,
# rows 81/96/102 typed opening balances.
CANONICAL_HIST = {
    8: "={c}26/{p}26-1",
    9: "={c}27/{c}26",
    10: "={c}30/{c}26",
    11: "={c}31",
    12: "={c}32/{c}47",
    13: "={c}33/{c}52",
    14: "={c}37/{c}35",
    15: "={c}45/{c}26*365",
    16: "={c}46/{c}27*365",
    17: "={c}51/{c}27*365",
    18: "={c}72",
    19: "={c}76",
    20: "={c}77",
    28: "={c}26-{c}27",
    34: "=SUM({c}30:{c}33)",
    35: "={c}28-{c}34",
    38: "={c}35-{c}37",
    48: "=SUM({c}44:{c}47)",
    53: "=SUM({c}51:{c}52)",
    57: "=SUM({c}55:{c}56)",
    58: "={c}53+{c}57",
    60: "={c}58-{c}48",
    69: "={c}66+{c}67-{c}68",
    73: "=SUM({c}72)",
    78: "=SUM({c}76:{c}77)",
    80: "={c}69-{c}73+{c}78",
    81: "={p}82",
    82: "=SUM({c}80:{c}81)",
    89: "={c}45",
    90: "={c}46",
    91: "={c}51",
    92: "={c}89+{c}90-{c}91",
    93: "={c}92-{p}92",
    96: "={p}99",
    97: "={c}72",
    98: "={c}67",
    99: "={c}96+{c}97-{c}98",
    102: "={p}104",
    103: "={c}76",
    104: "={c}102+{c}103",
    105: "={c}33",
}

# Rows whose FORECAST cells must be blue-input constants, not formulas.
FCST_INPUT_ROWS = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]


def norm(f):
    """Normalize a formula for comparison: uppercase, strip spaces, $, and
    leading '+' after '=' (Lotus-style '=+D2+1' == '=D2+1')."""
    if not isinstance(f, str):
        return f
    f = f.upper().replace(" ", "").replace("$", "")
    f = re.sub(r"^=\+", "=", f)
    return f


def shift_cols(formula, delta):
    """Translate a formula's column letters (A-Z single-letter refs) by delta."""
    def repl(m):
        col, row = m.group(1), m.group(2)
        return chr(ord(col) + delta) + row
    return re.sub(r"\b([A-Z])(\d+)\b", repl, formula)


def simulate(inputs):
    """Independent re-implementation of the model economics. Returns
    {row -> [10 values for D..M]} for every computed line. Historical columns
    reproduce input hardcodes + subtotal math; forecast columns run the
    recurrences. This never looks at the workbook being tested."""
    h = inputs["historical"]
    fa = inputs["forecast_assumptions"]
    nh, nf = inputs["n_historical"], inputs["n_forecast"]
    n = nh + nf
    dip = inputs.get("days_in_period", 365)

    def series(key):
        return list(h[key])

    rev = series("revenue"); cogs = series("cogs")
    sal = series("salaries_and_benefits"); rent = series("rent_and_overhead")
    depr = series("depreciation_amortization"); interest = series("interest_expense")
    tax = series("taxes"); cash = series("cash"); ar = series("accounts_receivable")
    inv = series("inventory"); ppe = series("ppe_closing"); ap = series("accounts_payable")
    debt = series("debt_closing"); eq = series("equity_capital"); re_ = series("retained_earnings")
    capex = series("capex"); dt_iss = series("debt_issuance_repayment")
    eq_iss = series("equity_issuance_repayment")

    ppe_open = [inputs["historical"]["ppe_opening_first_year"]] + ppe[:-1]
    debt_open = [inputs["historical"]["debt_opening_first_year"]] + debt[:-1]
    cash_open = [inputs["historical"]["opening_cash_first_year"]] + cash[:-1]

    for t in range(nf):
        rev.append(rev[-1] * (1 + fa["revenue_growth_pct"][t]))
        cogs.append(rev[-1] * fa["cogs_pct_revenue"][t])
        sal.append(rev[-1] * fa["salaries_pct_revenue"][t])
        rent.append(fa["rent_and_overhead"][t])
        ppe_open.append(ppe[-1])
        depr.append(ppe_open[-1] * fa["da_pct_opening_ppe"][t])
        ppe.append(ppe_open[-1] + fa["capex"][t] - depr[-1])
        capex.append(fa["capex"][t])
        debt_open.append(debt[-1])
        debt.append(debt_open[-1] + fa["debt_issuance_repayment"][t])
        dt_iss.append(fa["debt_issuance_repayment"][t])
        eq_iss.append(fa["equity_issuance_repayment"][t])
        interest.append((debt_open[-1] + debt[-1]) / 2 * fa["interest_pct_avg_debt"][t])
        gp_t = rev[-1] - cogs[-1]
        texp_t = sal[-1] + rent[-1] + depr[-1] + interest[-1]
        ebt_t = gp_t - texp_t
        tax.append(ebt_t * fa["tax_pct_ebt"][t])
        ar.append(rev[-1] * fa["ar_days"][t] / dip)
        inv.append(cogs[-1] * fa["inventory_days"][t] / dip)
        ap.append(cogs[-1] * fa["ap_days"][t] / dip)
        eq.append(eq[-1] + eq_iss[-1])

    gp = [rev[i] - cogs[i] for i in range(n)]
    texp = [sal[i] + rent[i] + depr[i] + interest[i] for i in range(n)]
    ebt = [gp[i] - texp[i] for i in range(n)]
    ne = [ebt[i] - tax[i] for i in range(n)]

    nwc = [ar[i] + inv[i] - ap[i] for i in range(n)]
    dnwc = [nwc[0]] + [nwc[i] - nwc[i - 1] for i in range(1, n)]
    # Historical CFS lines are inputs (as reported); forecast lines are computed.
    ne_cf = series("net_earnings_cf") + ne[nh:]
    depr_cf = series("depreciation_cf") + depr[nh:]
    dnwc_cf = series("change_in_nwc_cf") + dnwc[nh:]
    cfo = [ne_cf[i] + depr_cf[i] - dnwc_cf[i] for i in range(n)]
    cfi = list(capex)
    cff = [dt_iss[i] + eq_iss[i] for i in range(n)]
    dcash = [cfo[i] - cfi[i] + cff[i] for i in range(n)]
    for t in range(nf):
        cash_open.append(cash[-1])
        cash.append(cash_open[-1] + dcash[nh + t])
        re_.append(re_[-1] + ne[nh + t])

    ta = [cash[i] + ar[i] + inv[i] + ppe[i] for i in range(n)]
    tl = [ap[i] + debt[i] for i in range(n)]
    se = [eq[i] + re_[i] for i in range(n)]
    tle = [tl[i] + se[i] for i in range(n)]

    return {
        26: rev, 27: cogs, 28: gp, 30: sal, 31: rent, 32: depr, 33: interest,
        34: texp, 35: ebt, 37: tax, 38: ne,
        44: cash, 45: ar, 46: inv, 47: ppe, 48: ta,
        51: ap, 52: debt, 53: tl, 55: eq, 56: re_, 57: se, 58: tle,
        60: [tle[i] - ta[i] for i in range(n)],
        66: ne_cf, 67: depr_cf, 68: dnwc_cf, 69: cfo,
        72: cfi, 73: cfi, 76: dt_iss, 77: eq_iss, 78: cff,
        80: dcash, 81: cash_open, 82: cash,
        89: ar, 90: inv, 91: ap, 92: nwc, 93: dnwc,
        96: ppe_open, 97: capex, 98: depr, 99: ppe,
        102: debt_open, 103: dt_iss, 104: debt, 105: interest,
    }


def recalc_workbook(path, sheet_name):
    """Return {cell_ref -> numeric value} for the model block, recalculated.
    Strategy: LibreOffice headless round-trip (computes uncached formula
    results), else cached values, else the `formulas` package."""
    values = {}

    # Know which cells are formulas up front: with data_only=True a
    # never-recalculated formula cell reads as None, indistinguishable from an
    # empty cell unless we check the formula view.
    wb_formulas = openpyxl.load_workbook(path, data_only=False)
    ws_formulas = (wb_formulas[sheet_name] if sheet_name in wb_formulas.sheetnames
                   else wb_formulas.active)
    formula_cells = set()
    for row in range(2, 117):
        for col in COLS_ALL:
            v = ws_formulas[f"{col}{row}"].value
            if isinstance(v, str) and v.startswith("="):
                formula_cells.add(f"{col}{row}")

    def harvest(wb):
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        out, missing = {}, 0
        for row in range(2, 117):
            for col in COLS_ALL:
                ref = f"{col}{row}"
                v = ws[ref].value
                if v is None:
                    if ref in formula_cells:
                        missing += 1
                    continue
                if isinstance(v, str) and v.startswith("="):
                    missing += 1
                else:
                    out[ref] = v
        return out, missing

    lo = shutil.which("libreoffice") or shutil.which("soffice")
    if lo:
        with tempfile.TemporaryDirectory() as td:
            try:
                subprocess.run(
                    [lo, "--headless", "--convert-to", "xlsx", "--outdir", td, path],
                    check=True, capture_output=True, timeout=180,
                    env={**os.environ, "HOME": td},
                )
                out_path = os.path.join(
                    td, os.path.splitext(os.path.basename(path))[0] + ".xlsx")
                wb = openpyxl.load_workbook(out_path, data_only=True)
                values, missing = harvest(wb)
                if values and missing == 0:
                    return values, "libreoffice"
            except Exception:
                pass

    wb = openpyxl.load_workbook(path, data_only=True)
    values, missing = harvest(wb)
    if values and missing == 0:
        return values, "cached"

    # Cached values are absent/incomplete (freshly built workbook, no Excel/
    # LibreOffice save). Evaluate with the pure-Python `formulas` package.
    try:
        import formulas  # noqa
        xl = formulas.ExcelModel().loads(path).finish()
        sol = xl.calculate()
        base = os.path.basename(path).upper()
        target = sheet_name.upper()
        for k, v in sol.items():
            m = re.match(rf"'\[{re.escape(base)}\]([^']+)'!([A-Z]+\d+)", k.upper())
            # Filter to the target sheet: keying by bare coordinate would let
            # a same-address cell on another sheet (dashboard/checks) clobber a
            # model cell. Harmless when the workbook is single-sheet.
            if m and m.group(1) == target:
                try:
                    values[m.group(2)] = v.value[0, 0]
                except Exception:
                    pass
        if values:
            return values, "formulas-pkg"
    except Exception as e:
        print(f"  (formulas fallback failed: {e})")

    print("WARNING: no recalculation engine succeeded (LibreOffice unavailable "
          "and the 'formulas' package failed) - computed cells cannot be "
          "verified. Install LibreOffice or `pip install formulas`.")
    return values, "cached-partial"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("model")
    ap.add_argument("--inputs", required=True)
    ap.add_argument("--jurisdiction", help="jurisdiction pack JSON (optional)")
    ap.add_argument("--reference")
    ap.add_argument("--tolerance", type=float, default=0.01)
    ap.add_argument("--sheet", default="Three Statement Model")
    args = ap.parse_args()

    inputs = json.load(open(args.inputs))
    nh = inputs["n_historical"]
    y0 = inputs["first_historical_year"]
    set_columns(nh, inputs["n_forecast"])
    pack = json.load(open(args.jurisdiction)) if args.jurisdiction else None
    # Mirror build_model.resolve_labels: payload units_label > pack > default;
    # row 37 falls back to the pack's default_tax_line.label; overrides win.
    units = (inputs.get("units_label")
             or (pack or {}).get("units_label", "$000's"))
    for r in list(ROW_LABELS):
        ROW_LABELS[r] = ROW_LABELS[r].replace("($000's)", f"({units})")
    if pack:
        tax_label = (pack.get("default_tax_line") or {}).get("label")
        if tax_label:
            ROW_LABELS[37] = tax_label
        for r, label in (pack.get("label_overrides") or {}).items():
            ROW_LABELS[int(r)] = label
        # Pack may default the tax line; mirror build_model.py so the
        # simulation uses the same rates the workbook was built with.
        rate = (pack.get("default_tax_line") or {}).get("rate")
        fa = inputs["forecast_assumptions"]
        if rate is not None and not fa.get("tax_pct_ebt"):
            fa["tax_pct_ebt"] = [rate] * inputs["n_forecast"]
    if not inputs["forecast_assumptions"].get("tax_pct_ebt"):
        sys.exit("tax_pct_ebt is missing/empty: pass the same --jurisdiction pack "
                 "used at build time, or set tax_pct_ebt in the inputs JSON")
    failures, passes = [], 0

    def check(ok, msg):
        nonlocal passes
        if ok:
            passes += 1
        else:
            failures.append(msg)

    wb = openpyxl.load_workbook(args.model, data_only=False)
    check(args.sheet in wb.sheetnames,
          f"STRUCTURE: sheet '{args.sheet}' missing (found {wb.sheetnames})")
    ws = wb[args.sheet] if args.sheet in wb.sheetnames else wb.active

    # --- 1. STRUCTURE ---
    for row, label in ROW_LABELS.items():
        got = ws.cell(row=row, column=1).value
        check(isinstance(got, str) and got.strip().lower() == label.strip().lower(),
              f"STRUCTURE: A{row} expected '{label}', got '{got}'")
    for i, col in enumerate(COLS_ALL):
        cell = ws[f"{col}2"]
        if i == 0:
            check(cell.value == y0, f"STRUCTURE: D2 must be input year {y0}, got {cell.value!r}")
        else:
            check(isinstance(cell.value, str) and cell.value.startswith("="),
                  f"STRUCTURE: {col}2 must be a formula (prior year + 1), got {cell.value!r}")
    for col in COLS_ALL:
        f3 = ws[f"{col}3"].value
        check(isinstance(f3, str) and "IFERROR" in f3.upper() and "ABS" in f3.upper(),
              f"STRUCTURE: {col}3 balance-check flag formula missing/wrong: {f3!r}")

    # --- 2. FORMULAS ---
    # 2a. Forecast assumptions must be constants (inputs), not formulas.
    for row in FCST_INPUT_ROWS:
        for col in COLS_FCST:
            v = ws[f"{col}{row}"].value
            check(v is not None and not (isinstance(v, str) and v.startswith("=")),
                  f"FORMULA: {col}{row} (forecast assumption) must be a hardcoded input, got {v!r}")
    # 2b. Historical statement cells must be constants.
    for row in [26, 27, 30, 31, 32, 33, 37, 44, 45, 46, 47, 51, 52, 55, 56,
                66, 67, 68, 72, 76, 77]:
        for col in COLS_HIST:
            v = ws[f"{col}{row}"].value
            check(not (isinstance(v, str) and v.startswith("=")),
                  f"FORMULA: {col}{row} (historical actual) must be a hardcode, got formula {v!r}")
    # 2c. Canonical first-forecast-column formulas.
    first_fc = COLS_FCST[0]
    prior = chr(ord(first_fc) - 1)
    for row, tpl in CANONICAL_FCST.items():
        want = norm(tpl.format(c=first_fc, p=prior))
        got = ws[f"{first_fc}{row}"].value
        check(norm(got) == want,
              f"FORMULA: {first_fc}{row} expected {want}, got {norm(got) if isinstance(got, str) else got!r}")
    # 2c-bis. Canonical historical formulas (2nd historical column onward).
    def strip_iferror(f):
        """Builder wraps zero-denominator historical back-calcs in
        IFERROR(expr,"") - compare against the inner expression."""
        if isinstance(f, str):
            m = re.match(r'^=IFERROR\((.+),""\)$', f)
            if m:
                return "=" + m.group(1)
        return f

    for col in COLS_HIST[1:]:
        pcol = get_column_letter(openpyxl.utils.column_index_from_string(col) - 1)
        for row, tpl in CANONICAL_HIST.items():
            want = norm(tpl.format(c=col, p=pcol))
            got = strip_iferror(ws[f"{col}{row}"].value)
            check(norm(got) == want,
                  f"FORMULA: {col}{row} expected {want}, got {norm(got) if isinstance(got, str) else got!r}")
    # 2d. Column consistency: each subsequent forecast formula must be the
    # one-column-right translation of its neighbor (FAST: one formula per row).
    for row in CANONICAL_FCST:
        for a, b in zip(COLS_FCST, COLS_FCST[1:]):
            fa_, fb = ws[f"{a}{row}"].value, ws[f"{b}{row}"].value
            if isinstance(fa_, str) and isinstance(fb, str):
                check(norm(shift_cols(norm(fa_), 1)) == norm(fb),
                      f"FORMULA: {b}{row} is not the column-shifted copy of {a}{row} "
                      f"({norm(fb)} vs expected {norm(shift_cols(norm(fa_), 1))})")

    # --- 3. VALUES ---
    values, method = recalc_workbook(args.model, args.sheet)
    print(f"Recalculation method: {method} ({len(values)} cells)")
    sim = simulate(inputs)
    tol = args.tolerance
    for row, series in sim.items():
        for i, col in enumerate(COLS_ALL):
            ref = f"{col}{row}"
            expected = series[i]
            got = values.get(ref)
            if got is None:
                check(False, f"VALUE: {ref} ({ROW_LABELS.get(row)}) has no computed value")
                continue
            try:
                ok = abs(float(got) - expected) <= max(tol, abs(expected) * 1e-6)
            except (TypeError, ValueError):
                ok = False
            check(ok, f"VALUE: {ref} {ROW_LABELS.get(row)}: expected {expected:.4f}, got {got}")
    # Tie-outs (on recalculated values).
    for col in COLS_ALL:
        chk = values.get(f"{col}60")
        try:
            ok = chk is not None and abs(float(chk)) <= 1
        except (TypeError, ValueError):
            ok = False
        check(ok, f"TIE-OUT: balance check {col}60 = {chk} (|x|>1)")
        bs_cash, cf_cash = values.get(f"{col}44"), values.get(f"{col}82")
        try:
            ok = (bs_cash is not None and cf_cash is not None
                  and abs(float(bs_cash) - float(cf_cash)) <= tol)
        except (TypeError, ValueError):
            ok = False
        check(ok, f"TIE-OUT: BS cash {col}44={bs_cash} != CFS closing cash {col}82={cf_cash}")

    # No Excel error string may survive anywhere in the model block.
    for ref, v in values.items():
        check(not (isinstance(v, str) and v.startswith("#")),
              f"ERROR-CELL: {ref} evaluates to {v}")

    # --- 4. Optional regression vs reference workbook ---
    if args.reference:
        ref = json.load(open(args.reference))
        for cell, entry in ref["cells"].items():
            exp = entry.get("value")
            if not isinstance(exp, (int, float)):
                continue
            got = values.get(cell)
            ok = got is not None and abs(float(got) - exp) <= max(tol, abs(exp) * 1e-6)
            check(ok, f"REGRESSION: {cell} expected {exp}, got {got}")

    print(f"\n{'=' * 60}\nPASS: {passes}   FAIL: {len(failures)}")
    for f in failures[:80]:
        print("  ✗", f)
    if len(failures) > 80:
        print(f"  ... and {len(failures) - 80} more")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
