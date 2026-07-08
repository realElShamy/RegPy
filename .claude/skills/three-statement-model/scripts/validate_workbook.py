#!/usr/bin/env python3
"""
Acceptance harness for the FULL best-practice workbook (build_workbook.py).

Layers:
  0. CORE   - runs scripts/validate_model.py on the "Three Statement Model"
              sheet as a subprocess; exit 0 required (structure, formula canon,
              value regression, tie-outs, no error cells). The core sheet is
              byte-identical to the standalone build, so this is the same
              acceptance gate PR #1 established.
  1. SHAPE  - the four supporting sheets exist in order with tab colours; the
              cover carries navigation hyperlinks to each; the model defined
              names and the expected chart counts are present.
  2. CHECKS - the Checks sheet recalculates to master "OK", carries no "ERROR",
              and its plausibility master agrees with an independent simulation.
  3. DASH   - every KPI tile recalculates to the value an independent Python
              simulation of the model economics produces; no dashboard cell
              evaluates to an Excel error.
  4. SENS   - the scenario table, both two-way grids and the tornado endpoints
              on the Sensitivity sheet equal the deterministic economics, and a
              few values are tied independently back to the base-case simulation.

Recalculation uses LibreOffice if available, else the `formulas` package (the
same strategy as validate_model.py). Cross-sheet formulas are supported by both.

Usage:
  python3 validate_workbook.py MODEL.xlsx --inputs inputs.json
        [--jurisdiction pack.json] [--tolerance 0.01]
Exit code 0 = all layers pass, 1 = failures (report printed).
"""
import argparse
import json
import os
import re
import subprocess
import sys
import tempfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_model import apply_pack_defaults
from build_workbook import build_full, Geo, MODEL
from validate_model import simulate

HERE = os.path.dirname(os.path.abspath(__file__))
EXPECT_ORDER = ["Cover", MODEL, "Checks", "Dashboard", "Sensitivity"]
EXPECT_CHARTS = {MODEL: 2, "Dashboard": 6, "Sensitivity": 1}


def recalc_all(path):
    """{(SHEET_UPPER, CELLREF) -> value} for the whole workbook."""
    import formulas
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()
    base = os.path.basename(path).upper()
    pat = re.compile(rf"'\[{re.escape(base)}\]([^']+)'!([A-Z]+\d+)$")
    out = {}
    for k, v in sol.items():
        m = pat.match(k.upper())
        if not m:
            continue
        try:
            out[(m.group(1), m.group(2))] = v.value[0, 0]
        except Exception:
            out[(m.group(1), m.group(2))] = None
    return out


def is_err(v):
    return isinstance(v, str) and v.startswith("#") or \
        (v is not None and str(v).startswith("#"))


def expected_kpis(sim, n):
    i = n - 1
    ebitda = sim[35][i] + sim[33][i] + sim[32][i]
    return {
        "Revenue": sim[26][i],
        "Gross margin": sim[28][i] / sim[26][i] if sim[26][i] else None,
        "EBITDA": ebitda,
        "Net earnings": sim[38][i],
        "Closing cash": sim[82][i],
        "Total assets": sim[48][i],
        "Return on equity": sim[38][i] / sim[57][i] if sim[57][i] else None,
        "Current ratio": (sim[44][i] + sim[45][i] + sim[46][i]) / sim[51][i]
        if sim[51][i] else None,
        "Net debt / EBITDA": (sim[52][i] - sim[44][i]) / ebitda if ebitda else None,
    }


def expected_alert(sim, n):
    for i in range(n):
        if sim[82][i] < 0 or sim[57][i] < 0 or sim[38][i] < 0:
            return "REVIEW"
        if not sim[26][i]:
            return "REVIEW"
        gm = sim[28][i] / sim[26][i]
        if gm < 0 or gm > 1:
            return "REVIEW"
    return "CLEAR"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("model")
    ap.add_argument("--inputs", required=True)
    ap.add_argument("--jurisdiction")
    ap.add_argument("--tolerance", type=float, default=0.01)
    args = ap.parse_args()

    inputs = json.load(open(args.inputs))
    pack = json.load(open(args.jurisdiction)) if args.jurisdiction else None
    inputs = apply_pack_defaults(inputs, pack)
    nh, nf = inputs["n_historical"], inputs["n_forecast"]
    n = nh + nf
    geo = Geo(inputs)
    tol = args.tolerance

    failures, passes = [], 0

    def check(ok, msg):
        nonlocal passes
        if ok:
            passes += 1
        else:
            failures.append(msg)

    def close(a, b):
        try:
            return abs(float(a) - float(b)) <= max(tol, abs(float(b)) * 1e-4)
        except (TypeError, ValueError):
            return False

    # ---------- 0. CORE ----------
    cmd = [sys.executable, os.path.join(HERE, "validate_model.py"), args.model,
           "--inputs", args.inputs]
    if args.jurisdiction:
        cmd += ["--jurisdiction", args.jurisdiction]
    core = subprocess.run(cmd, capture_output=True, text=True)
    check(core.returncode == 0,
          "CORE: validate_model.py did not pass on the model sheet:\n"
          + "\n".join(core.stdout.splitlines()[-6:]))
    m = re.search(r"Recalculation method: (\S+)", core.stdout)
    core_method = m.group(1) if m else "?"

    # ---------- rebuild a reference to recover deterministic coordinates ----------
    with tempfile.TemporaryDirectory() as td:
        ref_path = os.path.join(td, "ref.xlsx")
        info = build_full(json.loads(json.dumps(inputs)), pack, ref_path)
    kpi_cells = info["kpi_cells"]
    placements = info["sensitivity"]["placements"]

    # ---------- load target ----------
    wb = openpyxl.load_workbook(args.model, data_only=False)

    # ---------- 1. SHAPE ----------
    check(wb.sheetnames == EXPECT_ORDER,
          f"SHAPE: sheet order {wb.sheetnames} != {EXPECT_ORDER}")
    for s in EXPECT_ORDER[1:]:
        if s in wb.sheetnames and s != MODEL:
            tc = wb[s].sheet_properties.tabColor
            check(tc is not None, f"SHAPE: sheet '{s}' has no tab colour")
    for s, want in EXPECT_CHARTS.items():
        got = len(wb[s]._charts) if s in wb.sheetnames else -1
        check(got == want, f"SHAPE: sheet '{s}' has {got} charts, expected {want}")
    for name in ("FinalYearRevenue", "FinalYearNetEarnings",
                 "FinalYearClosingCash", "ModelBalanceChecks"):
        check(name in wb.defined_names, f"SHAPE: defined name '{name}' missing")
    # cover navigation hyperlinks to each other sheet
    cov = wb["Cover"]
    targets = set()
    for row in cov.iter_rows():
        for cell in row:
            if cell.hyperlink and cell.hyperlink.location:
                targets.add(cell.hyperlink.location.split("!")[0].strip("'"))
    for s in ("Three Statement Model", "Checks", "Dashboard", "Sensitivity"):
        check(s in targets, f"SHAPE: cover has no hyperlink to '{s}'")

    # ---------- recalc for computed layers ----------
    vals = recalc_all(args.model)

    def g(sheet, cell):
        return vals.get((sheet.upper(), cell))

    sim = simulate(inputs)

    # ---------- 2. CHECKS ----------
    check(g("Checks", "C2") == "OK",
          f"CHECKS: integrity master C2 = {g('Checks', 'C2')!r} (expected OK)")
    exp_alert = expected_alert(sim, n)
    check(g("Checks", "C3") == exp_alert,
          f"CHECKS: alert master C3 = {g('Checks', 'C3')!r} (expected {exp_alert})")
    err_on_checks = [c for (s, c), v in vals.items()
                     if s == "CHECKS" and (v == "ERROR" or is_err(v))]
    check(not err_on_checks,
          f"CHECKS: {len(err_on_checks)} cell(s) flag ERROR/error: {err_on_checks[:8]}")

    # ---------- 3. DASHBOARD ----------
    ekpi = expected_kpis(sim, n)
    for name, coord in kpi_cells.items():
        exp = ekpi.get(name)
        got = g("Dashboard", coord)
        if exp is None:
            continue
        check(got is not None and close(got, exp),
              f"DASH: KPI '{name}' at {coord} = {got!r}, expected {exp:.4f}")
    err_on_dash = [c for (s, c), v in vals.items() if s == "DASHBOARD" and is_err(v)]
    check(not err_on_dash,
          f"DASH: {len(err_on_dash)} cell(s) evaluate to an error: {err_on_dash[:8]}")

    # ---------- 4. SENSITIVITY ----------
    ws_s = wb["Sensitivity"]
    for coord, want in placements.items():
        got = ws_s[coord].value
        check(got is not None and close(got, want),
              f"SENS: {coord} = {got!r}, expected {want}")
    # independent ties back to the base-case simulation
    base_ne = sim[38][n - 1]
    # scenario table: Base-case row is the first data row (row 7); its columns
    # C..G are revenue, ebitda, net earnings, closing cash, min-fc-cash.
    check(close(ws_s["E7"].value, base_ne),
          f"SENS: base-case net earnings E7 = {ws_s['E7'].value!r}, model = {base_ne:.2f}")
    check(close(ws_s["C7"].value, sim[26][n - 1]),
          f"SENS: base-case revenue C7 = {ws_s['C7'].value!r}, model = {sim[26][n-1]:.2f}")
    err_on_sens = [c for (s, c), v in vals.items() if s == "SENSITIVITY" and is_err(v)]
    check(not err_on_sens,
          f"SENS: {len(err_on_sens)} cell(s) evaluate to an error: {err_on_sens[:8]}")

    # ---------- report ----------
    print(f"Core recalc method: {core_method}   Enhanced-sheet recalc: formulas/LO")
    print(f"{'=' * 60}\nPASS: {passes}   FAIL: {len(failures)}")
    for f in failures[:60]:
        print("  ✗", f)
    if len(failures) > 60:
        print(f"  ... and {len(failures) - 60} more")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
