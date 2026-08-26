# ADR-0007: References with no single rectangle are refused, not guessed

**Status:** Accepted

**Date:** 2026-08-26

## Context

Excel has three reference operators. The colon makes a range, `A1:B5`. The comma makes
a *union* of disjoint blocks, `A1:A5,C1:C5`, which is one reference describing several
rectangles. A space makes an *intersection*, `A1:A10 A5:C5`, the cells the two overlap.
Unions in particular arrive without anyone meaning them to: a user Ctrl-clicks several
blocks, names the selection, and the defined name is stored as a comma-joined union.

EDJAS reads one rectangle and turns it into a scalar, a list or an object. A union has
no single shape, so there is no honest mapping. Should `A1:A5,C1:C5` become one flat
list, two lists, a table with a gap? Every answer is a guess, and a guess here returns
plausible, wrong data — the worst failure available to a tool whose output feeds reports.

Left alone the failure was also obscure: the union text reached openpyxl and produced
`too many values to unpack`, which tells the user nothing.

## Decision

A reference that cannot be read as one rectangle is refused with a message saying so.
A comma anywhere in a resolved reference — written inline or arriving via a defined
name — raises an error naming the reference. The same principle governs anything else
with no single-rectangle reading: a range spanning two sheets is declined rather than
silently read as the same shape on the first of them.

## Consequences

A user with a union range gets an error instead of wrong numbers, and the error implies
the remedy: give each block its own key, which produces better-named output than a
merged blob would anyway.

Nothing supports unions, and nothing will without reversing this. That was considered
and settled — multi-area support is YAGNI, and the clear error is the feature.

Intersections are handled less tidily. They are rejected on both paths, but by accident
rather than design: inline, the pipeline tokeniser splits on the space and complains
that the expression is not a single reference; through a defined name, openpyxl reports
an invalid coordinate. Both are refusals, neither is a good message. A purpose-built
guard was considered and rejected as not worth it — a named intersection range is
vanishingly rare, and a naive check for a space would false-positive on the quoted
sheet names of [ADR-0008](0008-table-names-are-references.md).

## Alternatives considered

**Read the first area and ignore the rest.** The most tempting and the most dangerous:
it always returns something, and what it returns is wrong in a way nobody notices.

**Return a list of areas**, one entry per rectangle. Coherent, and the natural
extension if this is revisited. Deferred because no real workbook has required it, and
adding a shape to the data model for a hypothetical case is how a small tool stops
being small.
