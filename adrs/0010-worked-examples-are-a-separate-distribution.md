# ADR-0010: The worked examples ship as a separate distribution

**Status:** Accepted

**Date:** 2026-08-26

## Context

Seven real government workbooks, their specifications and the templates that render
them are the most persuasive thing the project has: they are evidence rather than
description. They are also about a megabyte of binary data, and they have no business
inside a library that reads spreadsheets.

The obvious move — put them behind the existing `demo` extra — does not work, and the
reason is worth recording because it is not obvious. A Python extra can only name
*dependencies*; it cannot conditionally include *files*. Data either ships in the wheel
for every install or it does not ship at all. So "examples, but only if you ask for
them" cannot be expressed within one distribution.

## Decision

The examples live in a second distribution, `edjas-examples`, developed in the same
repository as a uv workspace member. It carries the seven workbooks, their
specifications, the shared templates, an `edjas-examples` command, and its own copy of
the licence. The core `demo` extra depends on it, so `pip install edjas[demo]` brings
Jinja2 and the examples together, while `pip install edjas` brings neither.

## Consequences

The core wheel stays around fifteen kilobytes of pure code and the example data is
downloaded only by people who asked for it. Both goals are met, which no
single-distribution arrangement could have achieved.

The cost is a release-ordering obligation that did not exist before, and it is sharp:
**`edjas-examples` must reach PyPI before, or together with, any `edjas` release
carrying the `demo` extra**, or `pip install edjas[demo]` fails to resolve for
everybody. A plain `uv build` at the workspace root builds only the root package; the
companion needs `uv build --package edjas-examples`. This is written into
`pyproject.toml` beside the extra because it is the kind of thing that is forgotten
exactly once.

Two distributions also means two version numbers, two sets of metadata and a
dependency from the companion back to the core, which pins a minimum version.

The examples package depends on `edjas`, and the documentation build depends on the
examples package, so the documentation cannot be built in a core-only environment.
The tests account for that by skipping rather than failing.

## Alternatives considered

**Ship the examples inside the core wheel.** Simplest, and it puts a megabyte of
government spreadsheets into every install of a library most users want for its own
sake. Rejected on footprint.

**Keep the examples in the repository only**, not packaged at all. This was the
arrangement before, and it means the examples are available only to people who clone
the repository — not to the `pip install` audience the documentation points at them.

**A separate repository.** Cleanest packaging story, worst maintenance story: the
examples exercise the library's behaviour and are verified against it by tests, and
splitting them would let the two drift with nothing to catch it.
