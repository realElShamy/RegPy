# Jurisdiction Guide — Egypt & GCC accounting and tax packs

How the packs in `../assets/jurisdictions/` were built, what they encode, when to
override them, and what to disclose. All facts verified against official and Big-4
sources (PwC Worldwide Tax Summaries, EY/KPMG alerts, regulator portals) with an
adversarial fact-check pass, **as of July 2026** (`as_of` field in each pack).
Rates change: if today is materially later than `as_of`, sanity-check the headline
rate before relying on a pack, and say you did.

## How packs plug into the model

The template carries ONE tax line: `Taxes = EBT × tax_pct_ebt`. Packs adapt it three ways:
1. **Default rate** — `default_tax_line.rate` fills `tax_pct_ebt` when the payload
   leaves it empty. Packs with ownership-dependent regimes (Qatar, Kuwait) set
   `rate: null`, which forces an explicit choice — the build fails loudly rather
   than guess.
2. **Labels/units** — `label_overrides` renames rows (e.g. Saudi "Zakat and Income
   Tax"); `units_label` rewrites the `($000's)` suffixes (e.g. `EGP '000`).
3. **Recipes + disclosures** — `tax_recipes` are the override formulas you apply for
   non-default situations; `disclosure` is the text that must accompany delivery.

Everything else about a jurisdiction (VAT, tax depreciation, filing calendar) is
context you disclose, not mechanics the annual model computes.

## Egypt (`eg`)

- **Standards:** Egyptian Accounting Standards (EAS) — an Arabic-language IFRS
  adaptation, not IFRS itself (Decree 110/2015; EAS 47/48/49 ≈ IFRS 9/15/16 from
  2019/2020; revaluation permitted from 2023 by PM Decree 883/2023; EAS 50 ≈ IFRS 17
  from July 2024). Statutory filings are Arabic. Two EAS quirks that can surprise an
  IFRS reader: employees' profit-sharing is an **equity distribution**, not a P&L
  expense; and 2016/2022 FX-devaluation losses were allowed to be deferred, so
  historical P&L may understate them.
- **Tax line:** CIT **22.5%** (Law 91/2005) → pack default `0.225`. Sector rates:
  40.55% oil & gas E&P, 40% Suez Canal Authority/EGPC/CBE — set explicitly if
  applicable. **Small-business option (Law 6/2025):** turnover ≤ EGP 20m may elect a
  turnover tax (0.4%–1.5% of *revenue* by band) replacing CIT — the recipe converts
  it to an EBT-equivalent rate per year. **Solidarity levy:** 0.25% of gross revenue
  (UHI Law 2/2018) — add `0.0025 × revenue/EBT` for thin-margin companies.
  No Zakat. No Pillar Two as of mid-2026. Note: a "Second Tax Facilitation Package"
  passed in June 2026 (listed-securities CGT → stamp duty, solidarity levy made
  deductible) — headline CIT unchanged.
- **VAT 14%** (5% machinery schedule) — disclosure note only.
- **Units/FY:** `EGP '000` (EGP m for large companies); calendar year default,
  July–June common for state-linked entities.

## Saudi Arabia (`sa`)

- **Standards:** IFRS as endorsed by SOCPA (all companies since FY2018; IFRS for
  SMEs available). Arabic statutory filings.
- **Tax line — the key GCC subtlety:** liability splits by ownership.
  *Saudi/GCC share* → **Zakat**: statutorily 2.5% of the **Zakat base** (a
  net-assets formula under MR 1007/2024, floored around adjusted net profit; the
  Hijri-vs-Gregorian proration ≈2.5778% applies to the non-profit base component
  only). The pack's `0.025 × EBT` default is a floor-level *proxy* for a fully
  Saudi-owned company — flag that asset-heavy/leveraged entities can owe materially
  more, because the real base is equity-driven.
  *Foreign share* → **CIT 20%** of its share of taxable income (oil/hydrocarbon
  50–85% tiered by capital).
  **Mixed ownership recipe:** `tax_pct_ebt = 0.20 × foreign_share + 0.025 ×
  saudi_gcc_share`. Label overrides rename rows 14/37 to "Zakat and Income Tax"
  (Tadawul presentation convention). No enacted Pillar Two as of mid-2026 (new
  Income Tax Law still draft).
- **VAT 15%**; SAR pegged 3.75/USD; return due 120 days after year end.

## United Arab Emirates (`ae`)

- **Standards:** IFRS (CCL FDL 32/2021; MD 114/2023 permits IFRS for SMEs below
  AED 50m revenue). English filings standard.
- **Tax line:** Federal CT **9%** above AED 375,000 taxable income; 0% below
  (FDL 47/2022, from FY starting June 2023) → pack default `0.09` (fine above
  ~AED 2m EBT; use the `band_exact` recipe near the band). Overrides:
  **Small Business Relief** (revenue ≤ AED 3m, periods ending ≤ 31 Dec 2026) → 0%;
  **Qualifying Free Zone Person** → 0% on qualifying income (MD 229/2025 activity
  list; de minimis lower of 5% of revenue / AED 5m) — blend and disclose;
  **DMTT 15%** (CD 142/2024) from FY2025 for ≥ EUR 750m groups.
- **VAT 5%**; AED pegged 3.6725/USD; CT return due 9 months after period end.

## Qatar (`qa`)

- **Standards:** IFRS (Islamic banks: AAOIFI FAS primary).
- **Tax line:** `rate: null` — ownership decides. CIT **10%** flat (Law 24/2018)
  on the **foreign-owned share** of profits; Qatari/GCC-owned share exempt (filing
  still required); oil operations ≥ 35%. Recipes: fully Qatari/GCC → 0; fully
  foreign → 0.10; mixed → `0.10 × foreign_share`. **Pillar Two is live**: 15%
  DMTT + IIR (Law 22/2024; Cabinet Resolution 2/2026) from FY2025 for ≥ EUR 750m
  groups.
- **No VAT** as of mid-2026 (e-invoicing draft law approved May 2026 — a likely
  precursor; note it in long forecasts). QAR pegged 3.64/USD; Dhareeba filing
  4 months after year end.

## Kuwait (`kw`)

- **Standards:** IFRS mandatory (MR 18/1990; CBK ECL modification for banks).
  Arabic filings; MOCI FS due ~3 months after year end.
- **Tax line:** `rate: null` — ownership *and* listing decide. **15%** income tax
  applies only to foreign (non-GCC) corporate bodies. Kuwaiti/GCC companies: no
  general CIT as of mid-2026, but quasi-taxes — NLST 2.5% (Boursa-listed), Zakat 1%
  (Kuwaiti KSCs), KFAS 1% — on differing bases; the `kuwaiti_listed` recipe uses
  ~4.5% of EBT as an aggregate proxy. **Watch item:** a broad 15% Business Profits
  Tax (KWD 1.5m exemption) is drafted for ~2027 but UNENACTED — flag it in any
  forecast reaching 2027. DMTT 15% live from FY2025 for ≥ EUR 750m groups.
- **No VAT.** KWD floats (no USD peg) and quotes to 3 decimals.

## Bahrain (`bh`)

- **Standards:** IFRS sole GAAP (CBB: AAOIFI FAS for Islamic banks).
- **Tax line:** default `0.0` — **no general CIT as of mid-2026**. The 46% regime is
  hydrocarbon-extraction only. **DMTT 15%** live from FY2025 for ≥ EUR 750m groups.
  **Watch item:** a draft **10% CIT** (BHD 1m revenue / BHD 200k profit thresholds,
  10% above BHD 200k) targeted at 1 Jan 2027, unenacted as of June 2026 — any
  forecast crossing 2027 should surface this and let the user pick a scenario.
- **VAT 10%** (registration BHD 37,500). BHD pegged 0.376/USD, 3-decimal fils.

## Oman (`om`)

- **Standards:** IFRS (IFRS for SMEs permitted since Feb 2025 FSA decision).
- **Tax line:** CIT **15%** → pack default `0.15`. **SME 3%** where registered
  capital ≤ OMR 60,000, gross income ≤ OMR 150,000, average employees ≤ 25
  (post-RD 118/2020 thresholds — note the older 100k/15 figures are obsolete).
  DMTT 15% from FY2025 for ≥ EUR 750m groups (already at the default rate).
- **VAT 5%**; OMR pegged at USD 2.6008 per OMR (~0.3845/USD), 3-decimal baisa.

## International fallback (`intl`)

IFRS assumed; `rate: null` — there is no defensible generic tax rate, so the payload
must set `tax_pct_ebt` explicitly (statutory or recent effective rate).

## Cross-cutting rules

1. **Never let a null-rate pack silently default.** Qatar/Kuwait/intl builds fail at
   input validation until the rate is set — that is intentional; resolve ownership
   with the user and apply the recipe.
2. **Pillar Two check (all GCC + Egypt):** if the entity belongs to a consolidated
   group with revenue ≥ EUR 750m in 2 of the last 4 years, DMTT (15%) likely
   overrides the pack default in AE/QA/KW/BH/OM from FY2025 — ask before defaulting.
3. **Zakat is not income tax.** Where Zakat applies (KSA always for the Saudi share;
   Kuwait 1% for KSCs), the %-of-EBT line is a proxy for an equity/net-assets-based
   levy — always include the pack's disclosure sentence.
4. **VAT never enters the P&L** in this template; it appears only as the AR/AP
   gross-of-VAT caveat in the delivery note.
5. **Historical taxes are as-reported** — packs shape only the *forecast* tax line;
   never restate historical tax charges to match the pack.
6. **Fiscal years:** the template is annual/365-day; align columns to the entity's
   actual fiscal year and label year headers accordingly (Egypt July–June state
   entities: label "FY2025/26" style in the delivery note, not in the sheet).
7. **Arabic filings:** for EG/SA/KW the statutory statements are Arabic; this model
   is a management/analysis artifact in English — say so when it matters.
