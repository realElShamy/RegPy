# Reverse Engineering: Case Study — Three Statement Model

A complete tear-down of the uploaded reference case-study workbook. Every fact below was extracted programmatically from the workbook
(formulas, cached values, styles, comments, charts, conditional formatting, raw OOXML).
This document is the ground truth behind the harness prompt in
[`PROMPT_THREE_STATEMENT_MODEL.md`](PROMPT_THREE_STATEMENT_MODEL.md).

---

## 1. Workbook at a glance

| Property | Value |
|---|---|
| Sheets | 1 visible sheet: **`Three Statement Model`** |
| Time axis | Columns **D–M** = fiscal years **2020–2029** |
| Historical vs forecast | **D–H** (2020–2024) hardcoded actuals; **I–M** (2025–2029) formula-driven forecast |
| Units | USD thousands (`$000's`) |
| Periodicity | Annual, 365 days per period (row 21) |
| View | Gridlines off, zoom 120%, top 3 rows frozen (`ySplit=3`) |
| Merged cells | `D1:H1` "Historical Results", `I1:M1` " Forecast Period" |
| Column widths | A=12.83, B=12.33, C=11.16, D=11.66, N=9.16 (rest default) |
| Row heights | 16 pt default; section-header rows 20 pt; year row (2) 24 pt |
| Circularity | **None** — deliberate design choice (see §7) |

The model is a **single-sheet, top-to-bottom** design in six sections, each introduced by a
styled section banner in column A:

| Row | Section |
|---|---|
| 5 | **Assumptions** |
| 24 | **Income Statement** |
| 41 | **Balance Sheet** |
| 63 | **Cash Flow Statement** |
| 86 | **Supporting Schedules** |
| 109 | **Charts and Graphs** |

---

## 2. Complete row map

Labels live in column A. Data occupies D–M. Column B/C are spacer columns (labels overflow
into them visually).

### Header block (rows 1–3)
| Row | Label / content | Formula pattern |
|---|---|---|
| 1 | `Historical Results` (D1:H1 merged), ` Forecast Period` (I1:M1 merged) | text |
| 2 | Fiscal years | `D2 = 2020` (input); `E2:M2 = +D2+1` chained |
| 3 | `Balance Sheet Check` | `=IFERROR(IF(ABS(D60)>1,"ERROR","OK"),"OK")` per column |

### Assumptions (rows 7–21) — the driver panel
Historical columns **compute** each ratio *from* the statements below (audit view);
forecast columns are **hardcoded blue inputs** that *drive* the statements.

| Row | Driver | Historical formula (D–H) | Forecast values (I–M, blue inputs) |
|---|---|---|---|
| 8 | Revenue Growth (% Change) | `=E26/D26-1` (D8 blank — no prior yr) | 10%, 10%, 10%, 10%, 10% |
| 9 | Cost of Goods Sold (% of Revenue) | `=D27/D26` | 42%, 47%, 50%, 36%, 35% |
| 10 | Salaries and Benefits (% of Revenue) | `=D30/D26` | 17% flat |
| 11 | Rent and Overhead ($000's) | `=D31` (link) | 15,000 flat |
| 12 | Depreciation & Amortization (% of PP&E) | `=D32/D47` | 35% flat — *cell comment: "% of opening PP&E"* |
| 13 | Interest (% of Debt) | `=D33/D52` | 10% flat — *cell comment: "Average balance"* |
| 14 | Tax Rate (% of Earnings Before Tax) | `=D37/D35` | 28% flat |
| 15 | Accounts Receivable (Days) | `=D45/D26*365` | 18 flat |
| 16 | Inventory (Days) | `=D46/D27*365` | 80, 90, 100, 100, 100 |
| 17 | Accounts Payable (Days) | `=D51/D27*365` | 37 flat |
| 18 | Capital Expenditures ($000's) | `=D72` (link) | 15,000 flat |
| 19 | Debt Issuance (Repayment) ($000's) | `=D76` (link) | 0, 0, −20,000, 0, 0 |
| 20 | Equity Issued (Repaid) ($000's) | `=D77` (link) | 0 flat |
| 21 | Days in Period | 365 (blue input, all 10 columns) | 365 |

### Income Statement (rows 26–38)
| Row | Line item | Historical (D–H) | Forecast (I–M) |
|---|---|---|---|
| 26 | Revenue | hardcode (blue, bold) | `=H26*(1+I8)` |
| 27 | Cost of Goods Sold (COGS) | hardcode | `=I26*I9` |
| 28 | **Gross Profit** | `=D26-D27` | same |
| 29 | *Expenses* (label only) | — | — |
| 30 | Salaries and Benefits | hardcode | `=I10*I26` |
| 31 | Rent and Overhead | hardcode | `=I11` |
| 32 | Depreciation & Amortization | hardcode | `=I67` ← CFS row (← depreciation schedule) |
| 33 | Interest | hardcode | `=I105` ← debt schedule |
| 34 | Total Expenses | `=SUM(D30:D33)` | same |
| 35 | **Earnings Before Tax** | `=D28-D34` | same |
| 37 | Taxes | hardcode | `=I35*I14` |
| 38 | **Net Earnings** | `=D35-D37` | same |

### Balance Sheet (rows 43–60)
| Row | Line item | Historical (D–H) | Forecast (I–M) |
|---|---|---|---|
| 44 | Cash | hardcode | `=I82` ← CFS closing cash |
| 45 | Accounts Receivable | hardcode | `=I26*I15/I21` (Revenue × AR days ÷ 365) |
| 46 | Inventory | hardcode | `=I27*I16/I21` (COGS × Inv days ÷ 365) |
| 47 | Property & Equipment | hardcode | `=I99` ← depreciation schedule closing PP&E |
| 48 | **Total Assets** | `=SUM(D44:D47)` | same |
| 51 | Accounts Payable | hardcode | `=I27*I17/I21` (COGS × AP days ÷ 365) |
| 52 | Debt | hardcode | `=I104` ← debt schedule closing |
| 53 | **Total Liabilities** | `=SUM(D51:D52)` | same |
| 55 | Equity Capital | hardcode | `=H55+I77` (prior + issuance from CFS) |
| 56 | Retained Earnings | hardcode | `=H56+I38` (prior + Net Earnings; **no dividends**) |
| 57 | **Shareholder's Equity** | `=SUM(D55:D56)` | same |
| 58 | **Total Liabilities & Shareholder's Equity** | `=D53+D57` | same |
| 60 | Check (italic) | `=D58-D48` | same — feeds row 3 flag |

### Cash Flow Statement (rows 65–82)
| Row | Line item | Historical (D–H) | Forecast (I–M) |
|---|---|---|---|
| 66 | Net Earnings | hardcode | `=I38` |
| 67 | Plus: Depreciation & Amortization | hardcode | `=I98` ← depreciation schedule |
| 68 | Less: Changes in Working Capital | hardcode | `=I93` ← WC schedule |
| 69 | **Cash from Operations** | `=D66+D67-D68` | same |
| 72 | Investments in Property & Equipment | hardcode (positive!) | `=I97` ← depreciation schedule |
| 73 | **Cash from Investing** | `=SUM(D72)` (positive!) | same |
| 76 | Issuance (repayment) of debt | hardcode | `=I103` ← debt schedule |
| 77 | Issuance (repayment) of equity | hardcode | `=I20` ← assumptions |
| 78 | **Cash from Financing** | `=SUM(D76:D77)` | same |
| 80 | **Net Increase (decrease) in Cash** | `=D69-D73+D78` (investing **subtracted**) | same |
| 81 | Opening Cash Balance | `D81 = 0` (blue input) | `=H82` (prior closing) |
| 82 | **Closing Cash Balance** | `=SUM(D80:D81)` | same → feeds BS row 44 |

### Supporting Schedules (rows 88–105)
**Working Capital Schedule (88–93)** — all columns are formulas:
| Row | Item | Formula (all columns) |
|---|---|---|
| 89 | Accounts Receivable | `=D45` (link from BS) |
| 90 | Inventory | `=D46` |
| 91 | Accounts Payable | `=D51` |
| 92 | Net Working Capital (NWC) | `=D89+D90-D91` |
| 93 | Change in NWC | `=D92-C92` (first column references empty C92 ⇒ ΔNWC = full NWC in year 1) |

**Depreciation Schedule (95–99)** — corkscrew:
| Row | Item | Historical | Forecast |
|---|---|---|---|
| 96 | PPE Opening | `D96 = 50,000` (blue input); `E96 = D99` | `=H99` (prior closing) |
| 97 | Plus Capex | `=D72` (link from CFS) | `=I18` (from assumptions) |
| 98 | Less Depreciation | `=D67` (link from CFS) | `=I96*I12` (**opening** PP&E × rate) |
| 99 | PPE Closing | `=D96+D97-D98` | same → feeds BS row 47 |

**Debt & Interest Schedule (101–105)** — corkscrew:
| Row | Item | Historical | Forecast |
|---|---|---|---|
| 102 | Debt Opening | `D102 = 50,000` (blue input); `E102 = D104` | `=H104` (prior closing) |
| 103 | Issuance (repayment) | `=D76` (link from CFS) | `=I19` (from assumptions) |
| 104 | Debt Closing | `=D102+D103` | same → feeds BS row 52 |
| 105 | Interest Expense | `=D33` (link from IS) | `=AVERAGE(I102,I104)*I13` → feeds IS row 33 |

*Cell comment on A102 (template author):* "We are using term debt/bond in this example so there's no compounding interest — revolver mechanics and compound interest are deferred to advanced material."

### Charts and Graphs data block (rows 111–116)
| Row | Series | Formula | Number format |
|---|---|---|---|
| 111 | Revenue | `=D26` | thousands |
| 112 | Gross Profit Margin (%) | `=D28/D26` | `0%` |
| 114 | Operating Cash Flow | `=D69` | thousands |
| 115 | Investing Cash Flow | `=-D73` (**negated** for display) | thousands |
| 116 | Financing Cash Flow | `=D78` | thousands |

**Embedded charts** (both use categories `D2:M2`):
1. **"Revenue & Gross Profit Margin"** — combo: clustered **bar** (Revenue, row 111) +
   **line** (GPM %, row 112, secondary axis).
2. **"Cash Flow"** — **stacked bar** (`overlap=100`) of rows 114/115/116.

---

## 3. The wiring diagram (dependency graph)

```
                          ┌────────────────────────────┐
                          │  ASSUMPTIONS (rows 8–21)   │
                          │  forecast cols = blue      │
                          └──────┬─────────────────────┘
        growth/margins/tax ▼             ▼ days drivers        ▼ capex/debt/equity $
┌──────────────────────┐   ┌──────────────────────┐   ┌─────────────────────────────┐
│ INCOME STATEMENT     │   │ BALANCE SHEET        │   │ SUPPORTING SCHEDULES        │
│ Rev→COGS→GP          │   │ AR  = Rev×days/365   │   │ Depreciation corkscrew      │
│ Salaries, Rent       │   │ Inv = COGS×days/365  │   │  open→+capex→−depr→close ───┼─→ BS PP&E (47)
│ D&A    ←──────────────────┼──────────────────────┼───┤  depr = open PP&E × rate ───┼─→ IS D&A (32) & CFS (67)
│ Interest ←────────────────┼──────────────────────┼───┤ Debt corkscrew              │
│ →EBT→Tax→Net Earnings │   │ AP  = COGS×days/365  │   │  open→±issuance→close ──────┼─→ BS Debt (52)
└──────────┬───────────┘   │ Debt ← schedule      │   │  interest=AVG(open,close)×r─┼─→ IS Interest (33)
           │               │ Equity roll-forward  │   │ Working-capital schedule    │
  Net Earnings (38)        │ RE roll-forward      │   │  NWC = AR+Inv−AP, ΔNWC ─────┼─→ CFS (68)
           ▼               │ Cash ← CFS (82) ◄────┼─┐ └─────────────────────────────┘
┌──────────────────────┐   └──────────┬───────────┘ │
│ CASH FLOW STATEMENT  │              ▼             │
│ NE + D&A − ΔNWC = CFO│   Check row 60 = (L&E − A) │
│ − Capex (investing)    │   Row 3 flag OK/ERROR      │
│ + Debt/Equity (CFF)  │                            │
│ → Net Δ cash         │                            │
│ open + Δ = closing ──┴────────────────────────────┘
└──────────────────────┘
```

**Cash is the balancing item.** Closing cash from the CFS is the only path into BS cash;
the balance sheet then ties by construction, and row 60 (`L&E − Assets`) proves it.

**The "pull-down / push-up" pattern.** In historical columns the assumption rows *pull*
ratios out of hardcoded statements (e.g. `D9 = D27/D26`); in forecast columns the
statements *pull* from the assumption inputs (e.g. `I27 = I26*I9`). The same inversion
happens on capex/debt/equity: historical assumption rows link *down* to the CFS
(`D18 = D72`), while forecast schedules link *up* to assumptions (`I97 = I18`).

---

## 4. Formatting system (exact spec)

| Element | Spec |
|---|---|
| Body font | Arial Narrow 12 pt |
| Section banners (A5, A24, A41, A63, A86, A109) | Open Sans 14 pt bold, font `#3271D2`, fill `#E7F2FF` |
| Year header D2:H2 (historical) | Arial Narrow 14 bold, fill `#E7F2FF` |
| Year header I2:M2 (forecast) | Arial Narrow 14 bold, **white** font, fill `#000C3F` (dark navy) |
| Banner row 1 | matching fills; centered bold |
| **Inputs (hardcodes)** | font **blue `#0000FF`** — all historical statement values, all forecast assumptions, Days in Period, opening balances (D81, D96, D102), year 2020 (D2) |
| **Formulas** | black font |
| Subtotals (Gross Profit 28, EBT 35, Total Liabilities 53, NWC 92) | bold + thin top border |
| Grand totals (Net Earnings 38, Total Assets 48, Total L&E 58) | bold + thin top border + **double bottom border** |
| Shareholder's Equity (57) | bold + thin top + thin bottom border |
| Interest row 33 | thin bottom border (closes the SUM block above Total Expenses) |
| Revenue rows (26, and label A26) | bold |
| Check row 60 | italic, format `0.0000_ ;\-0.0000\ ` |
| Check flag row 3 | italic 10 pt, right-aligned; conditional format: text "OK" → green font `#006100`; text "ERROR" → red font `#9C0006` on pink fill `#FFC7CE` |
| Standard number format | `_-* #,##0_-;\(#,##0\)_-;_-* "-"_-;_-@` (thousands separator, parentheses negatives, dash for zero) |
| Percent drivers | `0.0%`; chart GPM row `0%` |

The blue-input / black-formula convention is the industry standard; this workbook applies
it strictly — every blue cell is a typed constant, every black data cell is a formula.

---

## 5. Sign conventions (critical, and slightly idiosyncratic)

1. **Capex is displayed positive.** Row 72 "Investments in PP&E" = +15,000 and row 73
   "Cash from Investing" = `SUM(D72)` is **positive**; the outflow happens in row 80:
   `Net Δ Cash = CFO − Cash from Investing + CFF`. The chart block then re-negates it (`=-D73`) so the
   stacked chart shows investing below zero.
2. **ΔNWC is subtracted** ("Less: Changes in Working Capital"): CFO = NE + D&A − ΔNWC.
   An *increase* in NWC consumes cash.
3. **Debt issuance sign**: positive = issuance, negative = repayment (e.g. −20,000 in
   2022/2027). One row handles both.
4. **All expenses positive** on the IS; subtraction happens in subtotal rows
   (`GP = Rev − COGS`, `EBT = GP − Total Expenses`, `NE = EBT − Taxes`).

## 6. Check / integrity system

- Row 60: `= Total L&E − Total Assets` per column (should be 0.0000).
- Row 3: `=IFERROR(IF(ABS(D60)>1,"ERROR","OK"),"OK")` — tolerance of 1 ($1,000 since
  units are $000s), wrapped in IFERROR so a broken model shows a flag rather than `#REF!`.
- Conditional formatting turns the flag green (OK) or red-on-pink (ERROR).
- Frozen top 3 rows keep years + the check flag visible while scrolling — the checks are
  a *dashboard*, always in view.

## 7. Deliberate design choices (why the model is shaped this way)

1. **No circular reference.** Interest is charged on the *average* of opening and closing
   debt, which usually creates circularity (interest → net income → cash → revolver →
   debt → interest). Here debt movements are **exogenous inputs** (a term-debt schedule,
   no revolver, no cash sweep), so the loop never closes. The template author's comment on A102 makes
   this explicit. Cash simply accumulates.
2. **Corkscrew (roll-forward) schedules** for PP&E, debt, equity capital, retained
   earnings, and cash: `closing = opening + additions − reductions`, with each year's
   opening = prior year's closing. This is the backbone pattern of the whole model.
3. **Days-based working capital** (365 convention): AR from revenue; inventory and AP
   from COGS.
4. **D&A quirk**: the historical ratio row computes D&A ÷ *closing* PP&E (`D32/D47`),
   but the forecast applies the rate to *opening* PP&E (`I96*I12`, confirmed by the cell
   comment). The historical ratio is therefore only indicative — an intentional
   simplification in the reference case.
5. **First forecast column (I) is the only special column**: it reaches back into
   hardcoded historical column H (`I26=H26*(1+I8)`, `I81=H82`, `I96=H99`…). Columns J–M
   are pure copy-right of column I. This makes the model "one formula per row" compliant
   (FAST principle) within the forecast block.
6. **Retained earnings**: `RE(t) = RE(t−1) + Net Earnings(t)` — no dividend line in this
   case (a dividends row would subtract here and appear in CFF).

## 8. How the design squares with published methodology

Web research (the template author's own guides plus industry modeling standards) corroborates every
structural choice observed in the file:

- **Build order.** The provider's published sequence — historicals → historical ratios → forecast
  assumptions → income statement down to EBITDA → supporting schedules → wire D&A and
  interest back into the IS → balance sheet except cash → cash flow statement → closing
  cash into the BS — is exactly the dependency order embedded in this workbook's wiring
  (the author's published "3-Statement Model" guide).
- **Single sheet on purpose.** The provider "strongly recommends" the single-worksheet layout so
  every period lives in one column across all stacked sections, reducing mis-linking —
  this file follows that recommendation literally.
- **Color code.** The blue-inputs/black-formulas font convention is the universal Wall
  Street standard (blue hardcodes, black same-sheet formulas, green cross-sheet links —
  unused here since there is one sheet).
- **One row, one formula.** The FAST standard's core consistency rule (a single unique
  formula per row, filled across) holds for every forecast row; only the first forecast
  column differs by design, since it bridges into hardcoded history.
- **Checks.** A dedicated assets-vs-L&E tie-out with a red/green conditionally formatted
  flag row matches standard checks-dashboard practice. Note the template wraps only the
  *flag* in `IFERROR` (display safety), not calculation cells — blanket `IFERROR` around
  calculations is widely considered bad practice because it hides real errors.
- **Interest convention.** Interest on the *average* debt balance is the more accurate
  convention but normally risks circularity; industry teaching materials discuss the
  opening-balance alternative precisely for that reason. This template keeps average-
  balance interest *and* avoids circularity by making debt movements exogenous inputs —
  the beginner-friendly resolution of that trade-off (the revolver + circularity switch
  is deferred to advanced curricula).

## 9. Reference data extracted from the workbook

- [`harness/reference_values.json`](harness/reference_values.json) — every populated cell
  D2:M116: formula (or input) + cached value, plus row labels. Ground truth for the
  validation harness.
- [`harness/case_study_inputs.json`](harness/case_study_inputs.json) — the pure inputs
  (blue cells) reshaped as a parameterization payload: 5 years of historicals + 5 years
  of forecast assumptions.

Spot-check of forecast outputs (cached values, $000s):

| Metric | 2025 | 2026 | 2027 | 2028 | 2029 |
|---|---|---|---|---|---|
| Revenue | 165,849 | 182,434 | 200,678 | 220,745 | 242,820 |
| Net Earnings | 26,543 | 24,401 | 25,209 | 52,749 | 61,839 |
| Cash from Operations | 36,501 | 33,125 | 33,652 | 69,849 | 74,464 |
| Closing Cash | 161,050 | 179,174 | 177,827 | 232,676 | 292,140 |
| Total Assets = Total L&E | 223,884 | 249,916 | 256,605 | 307,239 | 369,637 |
| Balance check | 0 | 0 | 0 | 0 | 0 |
