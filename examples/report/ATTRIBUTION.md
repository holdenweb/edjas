# Source data and licence

## Bundled workbook — `retail.xlsx`

`retail.xlsx` is a copy of a data workbook published by the UK **Office for National
Statistics (ONS)**:

> **Retail Sales Index — Summary and Quality Tables**, Retail Sales Index, Great
> Britain, May 2025. Published 20 June 2025.
> More at <http://www.ons.gov.uk/businessindustryandtrade/retailindustry>.

It is redistributed here **unmodified**, purely as a realistic example for the EDJAS
reporting demo. ONS statistical outputs are Crown copyright and made available under the
**Open Government Licence v3.0**:

> Source: Office for National Statistics licensed under the
> [Open Government Licence v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/).

The workbook is a good example of the "[Releasing statistics in
spreadsheets](https://analysisfunction.civilservice.gov.uk/policy-store/releasing-statistics-in-spreadsheets/)"
good practice: a cover sheet of metadata, a contents sheet, and the actual data laid out
as a **named Excel Table** (`RevisionTriangles_Table1`) — exactly the constructs the
EDJAS spec (`retail.toml`) targets.

## Larger example — not bundled

For a bigger, multi-table workbook (too large to ship in the repository), the ONS
**UK National Balance Sheet** dataset is a good target. Download the reference-tables
workbook from the ONS national accounts pages and point a spec at its named tables the
same way; the same OGL v3 terms apply.
