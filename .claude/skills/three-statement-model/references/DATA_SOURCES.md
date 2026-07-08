# Data-Source Playbook — filling the input contract from whatever is connected

The skill adapts to the session's MCP connections. Probe with ToolSearch before
assuming anything is or isn't available; never tell the user a source is unsupported
without searching first. Priority order: (1) a connected accounting system, (2) files
the user uploaded or pointed to, (3) cloud drives, (4) manual entry.

Whatever the source, the target is always the same payload
([`../assets/inputs.schema.json`](../assets/inputs.schema.json)) and the same §1
validation identities. Record every mapping judgement you make (what went into COGS
vs Salaries vs Rent, what counts as debt) — the build report must disclose them.

## 1. Accounting-system MCP connectors

### Zoho Books (tools named `mcp__*Zoho*Books*` / `ZohoBooks_*`)
- `list_organizations` / `get_organization` → company name, **country (drives the
  jurisdiction pack)**, base currency, fiscal year start.
- `list_chart_of_accounts` → map account types: income → revenue; cost_of_goods_sold →
  COGS; expense accounts containing salary/wage/payroll → salaries_and_benefits; rent,
  utilities, overheads → rent_and_overhead; depreciation → D&A; interest → interest.
- P&L / BS reports where exposed; otherwise reconstruct annual figures from
  `list_invoices` (revenue, AR), `list_bills` (purchases, AP), `list_expenses`,
  `list_bank_accounts` (cash), fixed-asset accounts (PP&E).
- Working-capital days: compute from year-end AR/Inventory/AP balances and the
  derived revenue/COGS (the historical assumption rows will re-derive them in-sheet).

### Other ERPs (QuickBooks, Xero, Odoo, ERPNext, SAP-like connectors)
Same shape: find the organization/company endpoint (country, currency), the trial
balance or P&L/BS report endpoints, then map to the contract. Prefer report endpoints
over transaction reconstruction when both exist.

### CRM-only connectors (Zoho CRM, Apollo, etc.)
Not a books source. They can support revenue forecast assumptions (pipeline-based
growth) but never historicals — say so rather than improvising.

## 2. Files (uploads, Google Drive, SharePoint, Zoom docs)
- Uploaded .xlsx/.csv financial statements: parse with openpyxl/pandas. If the file
  is itself a statement export, map lines to the contract; if it's a prior model,
  extract the historical columns only.
- Google Drive (`mcp__Google_Drive__search_files` / `read_file_content`) and
  SharePoint (`mcp__ms365__sharepoint_*`): search for "financial statements",
  "trial balance", "management accounts", audit-report exports; download and parse.
- PDFs of audited statements: read the IS/BS/CFS pages; transcribe the needed lines.

## 3. Manual entry (always available)
Ask the user for, per historical year (2+ years, oldest first): revenue, COGS,
salaries, rent/overhead, D&A, interest, tax charge; year-end cash, AR, inventory,
net PP&E, AP, interest-bearing debt, share capital, retained earnings; capex,
debt/equity movements; plus the first year's opening cash/PP&E/debt. Then forecast
targets. Offer [`../assets/example_inputs.json`](../assets/example_inputs.json) as
a template they can edit.

## Mapping rules that keep §1 identities intact
The template's world is deliberately small: 4 asset lines, 2 liability lines,
2 equity lines. Real trial balances are bigger. Fold as follows and disclose it:
- "Cash" = cash + cash equivalents + short-term deposits.
- "Accounts Receivable" = trade receivables (+ net other current monetary assets if
  needed to make the BS identity close; disclose).
- "Inventory" = all inventories.
- "Property & Equipment" = net PP&E + intangibles + right-of-use assets (disclose).
- "Accounts Payable" = trade payables + accruals + other current liabilities that
  are not interest-bearing (disclose).
- "Debt" = ALL interest-bearing liabilities, current + non-current, incl. lease
  liabilities if material (disclose).
- "Equity Capital" = share capital + premium + reserves other than retained earnings.
- "Retained Earnings" = retained earnings/accumulated losses. If the audited RE
  movement includes dividends or OCI, the RE roll identity will fail — either fold
  those years' dividends into `equity_issuance_repayment` (negative) with disclosure,
  or adjust and disclose. Never force the identity silently.
- CFS lines (`*_cf`): take them as reported; the ±1 tolerance absorbs rounding. If
  the reported CFS categorizes differently (e.g. interest paid in financing),
  re-derive the three lines from the identity instead, and disclose.

## Delivery (after build + validation)
- Always: send the .xlsx to the user (SendUserFile or the harness equivalent).
- If a cloud drive is connected and the user wants it filed: upload/create there.
- Summarize: forecast net earnings + closing cash by year, the drivers used, the
  jurisdiction pack applied (tax basis + caveats), and all mapping disclosures.
