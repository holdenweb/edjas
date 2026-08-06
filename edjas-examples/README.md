# Seven spreadsheets, one method: the EDJAS demonstrations

*A guide for people who work with spreadsheets, on what these examples show and why it
might matter to you.*

## The problem these demonstrations address

If you produce or consume spreadsheets for a living — budgets, outturns, statistical
releases, management accounts — you will recognise the pattern. The workbook is the
master copy of something important. The numbers in it are needed elsewhere: in a report,
on a web page, in a briefing, in another system. And the way they usually get there is by
copying and pasting, by re-keying, by screenshotting a range into a document, or by
someone writing a fragile macro that breaks the next time a column moves.

The deeper problem is that a spreadsheet entangles three things that ought to be
separate: the **data**, the **layout** the author chose for it, and every downstream
**presentation** of it. Change any one and the others wobble. EDJAS — *Extract Data in
JSON from Any Spreadsheet* — is a small tool built on a simple discipline: leave the
workbook completely alone, and describe what you want out of it in a separate,
human-readable specification file. The spreadsheet is never edited, annotated, or
marked up. It is only ever read.

These seven demonstrations, all built on real, unmodified UK government workbooks
published under the Open Government Licence, show what that discipline buys you.

## What a specification looks like

An EDJAS spec is a short text file. Each line names a piece of output and says where in
the workbook it comes from:

```toml
[extract]
title     = "'Cover Sheet'!A1"
contents  = "['Contents'!A2:B5 | records]"
revisions = "[RevisionTriangles_Table1 | records | round 3]"
```

That is the entire extraction logic for the first demonstration. Three things are worth
noticing. References can use whatever the workbook already offers — a named Excel
Table, a named range, or a plain cell address, on any sheet. Small transformations
(reshape a table into records, round the figures, coerce text to numbers) are written as
a readable pipeline, not code. And because the spec is a file of its own, it can be
version-controlled, reviewed, diffed and shared — the extraction becomes a documented,
repeatable artefact rather than a ritual someone performs.

Run `edjas workbook.xlsx spec.toml` and you get JSON: the universal currency of modern
data exchange, readable by practically every programming language, database, dashboard
and web service in existence.

## The seven demonstrations

**`retail`** — the ONS Retail Sales Index summary tables — is the gentle introduction: a
statistics release with a cover sheet of metadata, a table of contents, and a data table
published, per government good practice, as a named Excel Table. The spec pulls the
title, publication dates and contact details from the cover sheet and the data by its
table name. If the ONS moves the table next month, the extraction still works, because
it never depended on cell coordinates.

The other five are **ledgers** — the shape this project set out to prove, because it is
the shape of so much working finance: repeated category headings, runs of line items,
and `Total …` roll-ups.

**`slgfs`** (Scottish local government finance) is a revenue account per council:
subservice lines rolling up through `Total Education`, `Total Social Work` and their
peers to `All Services`. **`pesa`** (HM Treasury) is a three-tier statement — capital and
current sections, budget categories, line items, subtotals — across six years.
**`dwp`** is a departmental spending table whose lines are parliamentary estimate
sections. **`wales`** is a full revenue-financing cascade whose columns deliberately mix
pounds, percentages and per-head figures. **`cra`** classifies all identifiable
government spending by international (COFOG) function, complete with `-` markers where
there was no spend.

Between them they also exercise the awkward realities of workbooks in the wild: sheet
names containing spaces, six-line column headings, padding cells, mixed precision, and
nil markers — each handled in the spec or the template, never by touching the source.

## The seventh: extraction as summarisation

The final demonstration, **`nbs`** — the ONS **UK National Balance Sheet** reference
tables — makes a different point: what to do when the workbook is simply *too much*.

This is a half-megabyte workbook of nineteen sheets. Each sector sheet is a matrix of
around fifty columns — every asset class the international framework recognises, coded
and cross-referenced, running back decades. It is an admirable publication, but it is
built for expert *reuse*, not for reading: nobody opens it to find out how the country
is doing. It also shows where workbook structure stops helping. Its two hundred and
sixteen Excel Tables are auto-generated fragments named `Table1`, `Table2`, and so on —
names exist, but they mean nothing, so the spec falls back on the workbook's other
handles: quoted sheet names and plain cell ranges.

The spec's answer to all that abundance is six lines. It takes the title and source
from the contents sheet, the list of reference tables, and exactly one statement: `Table
A`, net worth at the start of 2024, by institutional sector and asset class, twelve rows
ending at `Total economy`. The same ledger template that renders the spending statements
renders this balance sheet — a roll-up row is a roll-up row, whether it totals
expenditure or the nation's wealth.

Here extraction is doing the work of *summarisation*: the published page is not the
workbook made prettier, it is the workbook made **answerable** — one question, one
screen. The other eighteen sheets are still there for whoever needs them; the spec simply
declines to inflict them on readers who do not.

## From extraction to publication

Extraction alone earns its keep, but the demonstrations go one step further: each JSON
result is poured into an HTML template and emerges as a styled, readable web page.

The interesting part is how little template it takes. All six statement-shaped reports —
six different publishers, six different layouts, five spending ledgers and one balance
sheet — share **one** template. It looks at each extracted
row and classifies it: a row with a label but no figures is a section heading; a row
whose label begins `Total` or `All` is a subtotal; everything else is a line item.
Headings become full-width bands, subtotals are emphasised and ruled, line items are
indented, and figures are right-aligned with thousands separators. The dense, scrolling
grid of the workbook becomes a document a reader can actually follow — on a screen of
any size, in light or dark mode, with none of the squinting a pasted screenshot demands.

Because presentation choices live in the template rather than the data, they are made
once and apply everywhere. A small example: two ONS column headings run to
twenty-plus words each. The demonstration offers an optional renaming step —
off by default, so the faithful wording is preserved unless you ask otherwise. That is
the general pattern: fidelity by default, presentation by explicit choice.

## What you might do with this

The demonstrations use public statistics because anyone can verify them, but the method
transfers directly to private work:

- **Recurring reports.** If you produce a monthly workbook in a stable shape, one spec
  turns every future edition into a web page, a JSON feed, or both — the moment the
  workbook is saved. The copy-paste stage of the monthly cycle simply disappears.
- **Many audiences, one workbook.** Different specs can extract different slices of the
  same master workbook — a summary for the board, detail for the service manager,
  raw numbers for the analyst — without anyone maintaining parallel copies.
- **Feeding other systems.** JSON drops straight into dashboards, databases, budgeting
  tools and web applications. The spreadsheet remains the tool your team actually uses;
  EDJAS becomes the adaptor that lets everything else consume it.
- **Publication and accessibility.** A well-structured HTML page is searchable,
  linkable, responsive and far friendlier to assistive technology than a workbook
  attachment or a PDF print of one.
- **Audit and continuity.** A spec is a precise, reviewable statement of which figures a
  report uses and where they come from. When the author of the workbook moves on, the
  knowledge does not leave with them.

## The quiet point underneath

Nothing in these demonstrations required the six publishing bodies to change anything.
No macros were added, no columns inserted, no "please restructure your spreadsheet"
emails sent. The workbooks are byte-for-byte as published — one of the test suite's
checks is precisely that rendering leaves the source file unchanged.

That is the proposition in a sentence: keep working the way you work, and let a small,
readable text file turn the spreadsheets you already make into data anyone can use and
documents anyone can read.

*To build the reports yourself: install the demo extra (`pip install edjas[demo]`) and
run `edjas-examples` — the reports are written to your current directory. Sources and
licences are listed in
[ATTRIBUTION.md](https://github.com/holdenweb/edjas/blob/main/edjas-examples/src/edjas_examples/ATTRIBUTION.md).*
