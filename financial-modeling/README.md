# Three-Statement Financial Model — Reverse Engineering & Prompt Harness

This folder reverse engineers the **CFI Case Study Three Statement Model** workbook and
packages what was learned into a production-grade, harness-engineered prompt that makes
an LLM agent build equivalent models end-to-end for any company.

## Contents

| File | What it is |
|---|---|
| [`REVERSE_ENGINEERING.md`](REVERSE_ENGINEERING.md) | Complete tear-down of the original workbook: row map, every formula, the dependency graph, formatting system, sign conventions, checks, and the deliberate design choices (no-circularity, corkscrews, cash as the balancing item) |
| [`PROMPT_THREE_STATEMENT_MODEL.md`](PROMPT_THREE_STATEMENT_MODEL.md) | **The main deliverable.** An end-to-end system prompt (§0–§9): role, input contract, canonical cell map, dependency-ordered build phases with exact formula templates, formatting spec, known failure modes, and a mandatory self-validation protocol |
| [`harness/validate_model.py`](harness/validate_model.py) | Deterministic acceptance test. Validates any generated .xlsx in three layers: structure, formula integrity (incl. column-consistency), and values against an independent Python re-implementation of the model economics. Recalculates via LibreOffice headless / cached values / `formulas` |
| [`harness/case_study_inputs.json`](harness/case_study_inputs.json) | The CFI case study reshaped as the prompt's parameterization payload (5 historical years + 5 years of forecast assumptions) |
| [`harness/reference_values.json`](harness/reference_values.json) | Ground truth extracted from the original workbook — every populated cell's formula and cached value — for regression testing |

## Verification status

- The harness run against the **original CFI workbook**: **1,720 checks, 0 failures** —
  proving the reverse-engineered formula canon and the independent simulation exactly
  reproduce the source model.
- The prompt's acceptance gate is a **blind-build test**: an agent given only the prompt
  text and the inputs JSON (never the original workbook) must produce a workbook that
  passes the full harness, including recalculation and value regression to the cent.
  (Latest run status is recorded in the PR conversation.)

## Usage

```bash
# validate any generated model
python3 harness/validate_model.py MODEL.xlsx \
    --inputs harness/case_study_inputs.json \
    [--reference harness/reference_values.json]   # regression vs the CFI original
# exit code 0 = accepted
```

To build a model for a different company: keep the prompt unchanged, supply a new
inputs JSON conforming to the Input Contract (§1 of the prompt), and gate the output
with `validate_model.py --inputs <your inputs>` (omit `--reference`, which is specific
to the CFI case data).

## Model architecture in one paragraph

Single sheet, years across columns (5 historical hardcoded, 5 forecast formula-driven),
six sections down the rows: Assumptions → Income Statement → Balance Sheet → Cash Flow
Statement → Supporting Schedules (working capital, PP&E corkscrew, debt & interest
corkscrew) → chart feeds. Forecast statements pull from assumption inputs; historical
assumption rows back-calculate the same drivers from the hardcoded actuals. Cash is the
balancing item: the CFS's closing cash feeds the balance sheet, which therefore balances
by construction — proven by a live `L&E − Assets` check row with a conditional-format
OK/ERROR flag. Interest accrues on the average debt balance without circularity because
debt movements are exogenous inputs (term debt, no revolver).
