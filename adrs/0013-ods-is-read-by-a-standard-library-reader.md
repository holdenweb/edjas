# ADR-0013: OpenDocument files are read by a standard-library reader, presented as an openpyxl workbook

**Status:** Accepted

**Date:** 2026-08-26

## Context

Searching for worked examples kept producing files EDJAS could not open. The MHCLG
local-authority revenue outturn publication is eleven ODS files and no xlsx at all;
several of the Analysis Function's own good-practice examples are ODS. The tool aimed
at people stuck with government spreadsheets was answering a large part of its audience
with "no", not because it lacked a feature but because it could not read the file.

Weighed against a proposal to grow the specification language, this was the larger
gap: a reader who cannot open the file is not helped by a richer language.

Two questions followed. What to parse ODS with, and how to fit a second format into an
engine written entirely against openpyxl objects — references, named ranges, sheet
lookup and table resolution all assume that API.

Measurement settled the first. `odfpy` is the obvious library and costs about 670 KB
across two packages. But an ODS file is a zip containing XML, and `zipfile` and
`xml.etree` are both standard library, so the dependency buys convenience rather than
capability.

## Decision

ODS files are parsed by `edjas.ods` using only the standard library, and the result is
presented as an in-memory `openpyxl.Workbook`. `read_spec` dispatches on the file
extension; everything downstream — expressions, named ranges, pipelines, functions —
is unchanged and does not know which format it was given. ODS support is therefore
unconditional: it costs no dependency, so there is no extra to opt into.

## Consequences

The same specifications read both formats, and the project's name stopped
overpromising. The published wheel grew by about four kilobytes and gained no
dependency.

Presenting the result as an openpyxl workbook was the decision that made this small.
The alternative — teaching every part of the engine a second format — would have
touched reference resolution, table lookup and every function; instead the new format
is translated into the shape the engine already understands, and the existing tests
continue to cover it.

Writing a parser means owning its correctness, and review found that expensive. ODS
compresses runs of cells and rows with repeat counts, so a real file claims 1,048,576
rows and pads every row to 16,384 columns; mishandling that shifts a whole sheet.
Rows also nest inside wrappers for print titles and outline grouping, which a naive
walk drops — moving every row below them up, silently. Both are fixed and tested, but
they are a fair warning about what remains unknown: the format has corners this reader
has not met.

Fidelity is bounded in specific ways, each chosen over guessing. A range spanning two
sheets is declined rather than read off the first. A duration of 26 hours is returned
as written rather than wrapped to `02:30`. A file demanding more cells than a sane
ceiling is refused rather than exhausting memory. Sheet-scoped named ranges are not
hoisted to the workbook, where equal names would overwrite each other.

## Alternatives considered

**`odfpy`.** The mature, correct choice, and the one that would not have needed a
review to find repeat-count bugs. Rejected on weight: 670 KB and two packages for
something the standard library can do, in a project whose smallness is a stated asset.
This is the alternative to revisit if the reader's edge cases become a burden.

**`pandas` with its ODF engine.** Vastly heavier, and it would have imposed a
dataframe model on a tool that deliberately returns plain data.

**A second engine alongside openpyxl**, with the expression language taught to address
both. Rejected: it doubles the surface of every reference and every function for no
user-visible gain.

**Converting ODS to xlsx on disk first.** Would require a converter and would write a
file, which sits badly beside [ADR-0001](0001-the-workbook-is-never-modified.md).
