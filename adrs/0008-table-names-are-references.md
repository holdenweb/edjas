# ADR-0008: Excel Table names are references, with defined names taking precedence

**Status:** Accepted

**Date:** 2026-08-26

## Context

The project's advice is to reference named ranges rather than cell coordinates, because
names survive layout changes. Surveying real government spreadsheets for worked examples
showed that advice colliding with practice: the files that follow the Analysis Function's
*Releasing statistics in spreadsheets* guidance name their data as **Excel Tables**, not
defined names. The ONS retail release has zero defined names and one named table. EDJAS
resolved defined names and knew nothing about tables, so against the best-produced files
available it had nothing to offer but coordinates.

Tables are a different mechanism: they live per-sheet in `ws.tables` rather than
workbook-wide, and their range covers the header row and the data — which is exactly
what the `records` function wants.

Sheet names in these files also routinely contain spaces — `Cover Sheet`,
`Revisions triangles` — which Excel writes wrapped in apostrophes.

## Decision

A reference may be a defined name, the name of an Excel Table, or an A1-style range
optionally qualified with a sheet. A bare identifier that is not a defined name is
looked for as a table across every sheet, case-insensitively, and resolves to that
table's range. Precedence is defined name, then table name, then a raw reference. Sheet
names containing spaces or punctuation are written single-quoted exactly as Excel writes
them, `'Cover Sheet'!A1`, with a doubled apostrophe standing for one.

## Consequences

The good-practice files EDJAS is aimed at can be referenced the way their authors
intended, and the recommendation to prefer names over coordinates becomes true advice
rather than aspiration.

Because resolution happens in one place, all three expression forms gained table support
at once, and the change was small. Precedence rarely matters in practice: Excel keeps
defined names and table names in one namespace and forbids collisions, so
defined-name-first was chosen only because it leaves existing behaviour untouched.

Supporting quoted sheet names cost more than the tables did. It forced the tokeniser to
understand two quoting conventions at once — double quotes delimiting a string argument,
single quotes delimiting a sheet name, `''` escaping an apostrophe — and that dual rule
is now the most intricate part of the grammar. Before it, an inline reference to a
spaced sheet had never worked; it only appeared to, because such references had always
arrived via defined names, which are resolved before the tokeniser sees them.

Not every workbook plays along. The ONS national balance sheet carries 216
auto-generated tables named `Table1`, `Table2` and so on, useless as references, so its
specification falls back to coordinates.

## Alternatives considered

**Defined names only**, telling users to add a name for anything they want to reference.
Fails [ADR-0001](0001-the-workbook-is-never-modified.md): published files cannot be
edited, and those are the files that matter most here.

**Tables taking precedence over defined names.** No practical difference given Excel
forbids collisions, and it would have changed existing behaviour for no gain.
