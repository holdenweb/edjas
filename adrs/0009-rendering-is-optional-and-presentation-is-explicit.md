# ADR-0009: Rendering is optional, and presentation choices are explicit

**Status:** Accepted

**Date:** 2026-08-26

## Context

Extraction produces a data structure. Turning that into something a person reads is a
separate job, and a tempting one to take on: the whole argument for EDJAS is that
spreadsheet data should reach people in more assimilable forms, and JSON on stdout is
not that.

The risk was to the dependency footprint. EDJAS installs `openpyxl` and nothing else,
and that smallness is a genuine asset — it makes the tool easy to justify adding to an
environment. A templating engine is a reasonable dependency for someone producing
reports and an unreasonable one for someone piping JSON into a database.

A second question surfaced once real government data was flowing through: some column
headings run past twenty words, and reports built from them are unreadable. Shortening
them is obviously desirable and obviously a presentation decision — the workbook says
what it says.

## Decision

`render_report` feeds extracted data into a Jinja2 template, and Jinja2 is an optional
dependency: the core install never imports it, and calling `render_report` without it
raises an error naming the extra to install. Presentation adjustments are explicit and
default to off — `headings={old: new}` renames columns for display, and with no mapping
supplied the workbook's own wording is reproduced verbatim.

## Consequences

`pip install edjas` remains a one-dependency install and the published wheel is a few
kilobytes of code. Someone who wants reports opts in with `edjas[demo]`.

The default of fidelity matters more than it looks. The honest thing happens unless
somebody chooses otherwise, and the choice is recorded in the code that made it. The
same principle later settled an unrelated question: an ODS duration of 26 hours is
returned as written rather than wrapped into a plausible, wrong `02:30`.

A rename that collides with an existing heading raises rather than silently dropping a
column — found by review, not by design.

The awkward consequence is that the reporting templates began accumulating knowledge
that is really about the data: which rows are subtotals, which markers mean nil, which
padding to ignore. That is presentation logic doing data work, and it is the strongest
argument for the document-level transform deferred in
[ADR-0014](0014-transform-and-load-are-deferred.md).

## Alternatives considered

**Jinja2 as a core dependency.** Simpler to explain, at the cost of imposing a
templating engine on every user including those who only want JSON.

**A built-in report format** with no templating. Less flexible, and it would put EDJAS
in the business of designing report layouts — a much larger and more opinionated job
than rendering someone else's template.

**Shortening long headings by default.** Rejected: it silently alters what the publisher
wrote, and a reader comparing the report against the source would find them disagreeing
with no indication why.
