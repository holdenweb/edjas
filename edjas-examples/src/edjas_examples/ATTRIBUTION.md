# Source data and licences

All seven bundled workbooks are UK public-sector statistics, Crown copyright, and
redistributed here **unmodified** under the **Open Government Licence v3.0** purely as
realistic examples for the EDJAS reporting demo:

> Contains public sector information licensed under the
> [Open Government Licence v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/).

Each entry below records the exact URL the bundled copy was downloaded from, so the
provenance of every workbook can be checked. All seven were verified byte-for-byte
against their published originals.

## `retail.xlsx` — a flat statistics release

**Retail Sales Index — Summary and Quality Tables**, Retail Sales Index, Great Britain,
May 2025, Office for National Statistics. Published 20 June 2025.
More at [ons.gov.uk](http://www.ons.gov.uk/businessindustryandtrade/retailindustry).

A "[Releasing statistics in spreadsheets](https://analysisfunction.civilservice.gov.uk/policy-store/releasing-statistics-in-spreadsheets/)"
good-practice release: a cover sheet, a contents sheet, and data laid out as a **named
Excel Table** (`RevisionTriangles_Table1`).

Downloaded from [ons.gov.uk](https://www.ons.gov.uk/file?uri=/businessindustryandtrade/retailindustry/datasets/retailsalessummarytables/current/previous/v105/qualityandsummarytablesmay2025.xlsx).

## `slgfs.xlsx` — a local-government revenue ledger

**Scottish Local Government Finance Statistics 2024-25 — LA Level, Net Revenue
Expenditure by Subservice**, Scottish Government. A revenue account per council: subservice
line items grouped into categories, each closed by a `Total …` subtotal, down to
`All Services (GF + HRA)`. Every council sheet is a named Excel Table.
Collection: [gov.scot](https://www.gov.scot/collections/scottish-local-government-finance-statistics/).

Downloaded from [gov.scot](https://www.gov.scot/binaries/content/documents/govscot/publications/statistics/2026/02/scottish-local-government-finance-statistics-2024-25/documents/la-level---net-revenue-expenditure-by-subservice/la-level---net-revenue-expenditure-by-subservice/govscot%3Adocument/Scottish%2BLocal%2BGovernment%2BFinance%2BStatistics%2B%2528SLGFS%2529%2B2024-25%2B-%2BLA%2BLevel%2B-%2BNet%2BRevenue%2BExpenditure%2Bby%2BSubservice%2B-%2B2026-01-09.xlsx).

## `pesa.xlsx` — a central-government expenditure ledger

**Public Expenditure Statistical Analyses 2025 — Chapter 1 (Departmental budgets)**,
HM Treasury (Accredited Official Statistics). Table 1.1, Total Managed Expenditure, is a
three-tier ledger — `CURRENT`/`CAPITAL EXPENDITURE` sections, budget categories, line
items, and `Total …` roll-ups — over six years.
Collection: [GOV.UK](https://www.gov.uk/government/collections/public-expenditure-statistical-analyses-pesa).

Downloaded from [GOV.UK](https://assets.publishing.service.gov.uk/media/6874fa8f92691289bdb7d394/PESA_2025_CP_Chapter_1_tables.xlsx).

## `dwp.xlsx` — a departmental spending ledger

**Public spending and administration budget 2021 to 2026**, Department for Work and
Pensions — the data-tables companion to the DWP Annual Report and Accounts 2024-25.
Budget categories (Resource/Capital DEL and AME) head runs of `Section A/B/C…` estimate
lines, each closed by a `Total …` subtotal with `Of which:` breakdowns. Its sheet is
named `Table 1`, so the spec quotes it as Excel does.
Publication: [GOV.UK](https://www.gov.uk/government/publications/dwp-annual-report-and-accounts-2024-to-2025).

Downloaded from [GOV.UK](https://assets.publishing.service.gov.uk/media/686e85c6a08d3a3ca3b67966/dwp-spending-and-budget-2021-to-2026.xlsx).

Note: these tables are compiled on the HM Treasury OSCAR budgetary basis and, as their own
footnotes say, are not reported on the same basis as the audited financial statements.

## `wales.xlsx` — a revenue account with GSS metadata

**Local authority revenue and capital outturn expenditure: April 2024 to March 2025**,
Welsh Government. A GSS-format release — cover sheet, table of contents, notes, then one
table per sheet. Table 1 is a full revenue account running from service line items through
`Gross`/`Net revenue expenditure` and `Budget requirement` to `Council tax requirement`.
Its value columns deliberately mix £ million, change, percentage and £-per-head, which the
ledger template formats to differing precision.
More at [gov.wales](https://www.gov.wales/local-authority-revenue-and-capital-outturn-expenditure-april-2024-march-2025).

Downloaded from [gov.wales](https://gov.wales/sites/default/files/statistics-and-research/2025-10/local-authority-revenue-and-capital-outturn-expenditure-april-2024-march-2025-267.xlsx).

## `cra.xlsx` — a COFOG-classified expenditure ledger

**Country and Regional Analysis 2024, Chapter B tables**, HM Treasury. Table B.1 classifies
England's identifiable expenditure by COFOG function: ten numbered function headings, each
over numbered sub-function line items (some with nested `of which:` detail) and closed by a
`Total <function>` subtotal, ending at `Total Expenditure on Services in England`. Nil cells
hold `-`, which the template passes through rather than formatting as a number.
Publication: [GOV.UK](https://www.gov.uk/government/statistics/country-and-regional-analysis-2024).

Downloaded from [GOV.UK](https://assets.publishing.service.gov.uk/media/673c6cc79a48a5ab14acc372/CRA_2024_Chapter_B_tables.xlsx).

## `nbs.xlsx` — a large multi-sheet reference workbook

**The UK national balance sheet estimates**, Office for National Statistics — the full
reference-tables workbook: 19 sheets, including ~50-column time-series matrices per
institutional sector, carrying 216 auto-generated Excel Table fragments. The spec
deliberately extracts only the headline statement from `Table A` (net worth at the start
of 2024, by sector and asset class, down to the `Total economy` roll-up), demonstrating
how a small spec carves a readable brief out of a workbook built for expert reuse.
Releases: [ons.gov.uk](https://www.ons.gov.uk/economy/nationalaccounts/uksectoraccounts/bulletins/nationalbalancesheet/previousReleases).

Downloaded from [ons.gov.uk](https://www.ons.gov.uk/file?uri=/economy/nationalaccounts/uksectoraccounts/datasets/thenationalbalancesheetestimates/current/nbsreferencetables2025.xlsx).
