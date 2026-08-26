# ADR-0002: Extraction is described beside the spreadsheet, not inside it

**Status:** Accepted

**Date:** 2026-08-26

## Context

EDJAS originally read its instructions from the workbook. A `Parameters` named range
held cells whose values were markup — `[Sales | records]` and the like — and the tool
read that range to discover what to extract.

It worked, and it was wrong in three separate ways. It contradicted the promise in
[ADR-0001](0001-the-workbook-is-never-modified.md): the file had to be edited before
EDJAS could do anything with it. It put a second little language inside a document
that already contains a language, so every construct EDJAS invented had to be checked
against what Excel might one day do with the same characters. And it entangled the
description of the extraction with the data being extracted, so the two could not be
versioned, reviewed or reused separately — the exact confusion of data with
presentation that the project's own argument is against.

The practical trigger was the observation that one workbook usually has several
audiences, and each wants a different slice. With markup in the sheet there is one
extraction per workbook, and adding a second means editing the file again.

## Decision

An extraction is described in a separate TOML file, and the workbook is not consulted
about it at all. The specification names output keys and the expressions that produce
them; the workbook supplies only data. One specification can be applied to many
workbooks that share a shape, and one workbook can have many specifications, each
serving a different reader. `read_spec(spreadsheet, spec)` takes both paths and
returns a plain dictionary.

## Consequences

The specification becomes a first-class artefact: a text file that can be committed,
diffed, reviewed in a pull request and reasoned about without opening a spreadsheet.
That is the single largest gain, and it is what makes an extraction repeatable rather
than a ritual somebody performs.

The cost is indirection. A reader now needs two files open to understand what is
happening, and a specification that drifts from the workbook it describes fails at
run time rather than being visibly wrong in the sheet. Because references are often
pinned by coordinate, a republished workbook with two extra rows can silently shift
what gets read — a hazard recorded, unresolved, in
[ADR-0014](0014-transform-and-load-are-deferred.md).

The in-cell mechanism was removed outright rather than kept as a second mode, so the
change was breaking. It landed on 2026-07-22, when there were no external users to
break.

## Alternatives considered

**Keep both modes**, reading in-cell markup when present and a spec file otherwise.
Rejected: two ways to describe one thing doubles the surface to document and test,
and the in-cell mode would have kept the collision risk with Excel's own syntax alive
for no benefit.

**A sidecar sheet inside the workbook** — instructions on their own tab. Still
requires writing to the file, so it fails [ADR-0001](0001-the-workbook-is-never-modified.md)
for the same reason, and published statistical releases cannot carry one.
