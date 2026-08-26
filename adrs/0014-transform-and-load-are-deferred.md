# ADR-0014: A document-level `[transform]` and a `[load]` section are deferred

**Status:** Deferred

**Date:** 2026-08-26

## Context

The `[extract]` table invites two obvious companions: `[transform]` to reshape the
assembled document, and `[load]` to send it somewhere. Together they would make a
specification a complete, runnable job rather than a description of half of one, and
they answer real limitations.

The case for `[transform]` is the strongest, and it is not a feature request but a
defect report. The reporting templates currently decide which rows are section
headings, which are subtotals, which markers mean nil and which padding to ignore.
That is data semantics living in the presentation layer — precisely the entanglement
the project's own argument is against. Extract the CRA example to JSON today and the
consumer receives `"-"` strings mixed with numbers and no indication which rows are
roll-ups, having to re-derive what the template already worked out.
Pipelines cannot fix this: they see one key at a time and cannot reshape across keys.

The case for `[load]` is weaker under examination. Its stated beneficiary is the
non-programmer who should not have to write a shell pipeline — but someone who cannot
write `edjas x.xlsx s.toml > out.json` will not write `--bind out=sqlite:///stats.db`
either. It would also trade an absolute for a conditional: today the security claim of
[ADR-0005](0005-a-specification-cannot-execute-code.md) is one sentence with no
caveats, and a specification that can write paths or reach the network turns it into
four conditions. And each destination drags in dependency churn from ecosystems the
project does not control, against a core of one stable dependency.

There is also a scope question neither section answers on its own: accepting the frame
"pipeline runner" invites incremental runs, change detection, retries and scheduling.

## Decision

Neither section is implemented. `[extract]` remains the only table a specification
carries, and the reserved space for others in
[ADR-0003](0003-specifications-are-toml.md) stays reserved. Reshaping beyond a
per-key pipeline is done by the calling program; delivery is done by the shell.

## Consequences

The tool stays describable in a sentence and stays in a category it can win: narrowness
is its advantage against dbt, pandas and the Singer ecosystem, all of which do
transformation and loading better.

The known defect stays. Templates continue to hold data semantics, the JSON output is
less useful to a non-HTML consumer than it should be, and a second renderer would have
to re-derive the same classifications. That cost is accepted, not overlooked.

Anyone needing more reshaping writes Python around `read_spec`, which is the intended
extension point under [ADR-0005](0005-a-specification-cannot-execute-code.md).

## The triggers to watch for

This is a decision not to decide yet. Revisit when:

- **A second consumer of the extracted data needs the row classification** the
  templates perform. One renderer holding that knowledge is a smell; two would be
  duplication, and the argument for a narrow `[transform]` becomes conclusive.
- **A template hack becomes genuinely annoying to maintain.** The narrow transform is
  a refactor of logic that already exists, so it can wait for a real itch.
- **Someone asks for `[load]` who cannot already use a shell pipeline**, which would
  falsify the reasoning above.

If `[transform]` is built, it should be a fixed registry of declarative verbs —
`classify_ledger`, `nulls`, `group_by` — and not an expression language. Conditionals
and arithmetic in TOML are how the guarantee in
[ADR-0005](0005-a-specification-cannot-execute-code.md) dies.

## Alternatives considered

**Build both now.** Rejected above, on audience, on the security guarantee and on
maintenance.

**Build a broad `[transform]`** with joins, aggregation and derived columns. This is a
new language, badly, and it blurs "reads spreadsheets faithfully" into "spreadsheet ETL
tool" — a crowded category where the project's advantage disappears.

**Build `[validate]` instead**, asserting that subtotal rows sum to their components.
Genuinely attractive: cheap, no new dependencies, no category expansion, and it guards
the failure this design is most exposed to — specifications pin ranges by coordinate,
so a republished workbook with two extra rows silently swallows footnotes into the
data. Held rather than refused, and the strongest candidate of the three when this
record is revisited.
