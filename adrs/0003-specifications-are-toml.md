# ADR-0003: Specifications are TOML

**Status:** Accepted

**Date:** 2026-08-26

## Context

Having decided in [ADR-0002](0002-extraction-is-described-beside-the-spreadsheet.md)
that extractions live in their own file, the format had to be chosen. The people
expected to write these files are analysts and spreadsheet owners, not necessarily
programmers, so the format needed to be readable and forgiving of hand-editing. The
file also needed to carry short strings full of brackets, pipes, quotes and
apostrophes without ceremony, because that is what an extraction expression looks like.

Weighing on the other side: EDJAS had one dependency and wanted to keep it that way.
A format requiring a parser from PyPI would have doubled the install footprint of a
tool whose entire value is being small and boring.

## Decision

A specification is a TOML file containing an `[extract]` table that maps output keys
to extraction expressions. TOML is parsed with `tomllib` from the standard library.
Other top-level tables are reserved for future use and ignored today.

## Consequences

No new dependency, on any supported Python: `tomllib` has been stdlib since 3.11 and
EDJAS requires 3.12. The `[extract]` heading also gives the file an obvious extension
point, which is what a future `[transform]` or `[load]` would occupy —
see [ADR-0014](0014-transform-and-load-are-deferred.md).

TOML's string rules turned out to matter more than expected. Expressions contain
double quotes for string arguments and single quotes for Excel sheet names, so a
naive single-quoted TOML literal cannot hold every valid expression; the test helper
that writes specifications had to move to triple-quoted literals for exactly this
reason. Anyone writing specifications by hand will meet the same edge.

TOML is also flat by nature, which suits a mapping of keys to expressions and would
suit a richer structure less well. If the specification language ever needs genuine
nesting, this decision will be under pressure.

## Alternatives considered

**YAML.** More familiar to many, and comfortable with nesting, but not in the
standard library, significantly more complex to parse, and carries well-known
surprises around implicit typing that would bite people writing expressions like
`NO` or `12:30`.

**JSON.** Stdlib and unambiguous, but hostile to hand-editing: no comments, and the
worked examples lean heavily on comment headers to explain why a specification reads
the ranges it does. Losing comments would have cost more than the syntax saved.

**A bespoke DSL.** Considered seriously, and deferred rather than refused. It would
fit the expression language better than TOML's key-value shape does, but it means
writing and documenting a parser, and nothing yet requires it.
