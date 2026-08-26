# ADR-0000: Template

**Status:** Template — copy this file, do not edit it in place.

**Date:** yyyy-mm-dd

## Context

What was true that forced a choice. The pressure, not the answer: what did not
work, what had to be reconciled, what would have gone wrong. Enough that a
reader who was not there can feel the problem before reading the decision.

## Decision

One paragraph, in the present tense, saying what we do. Not "we will" — an
accepted ADR describes how the system is.

## Consequences

What follows, including what got worse. A record with no costs in it is a
record nobody will trust. Say what this forecloses and what it makes someone
else's problem.

## Alternatives considered

What else was on the table and why it lost. This is the section that stops a
decision being reproposed in six months.

---

## How to use these records

An ADR is **immutable once accepted**. When a decision changes, write a new
record that supersedes it and mark the old one `Superseded by ADR-nnnn`; do not
rewrite history, because the reasoning that was true at the time is the thing of
value.

Statuses in use: `Accepted`, `Deferred` (a decision to not decide yet, with the
trigger to watch for), `Rejected`, `Superseded by ADR-nnnn`.

**What lives where.** These three overlap and would otherwise drift:

| | Answers | Changes when |
| --- | --- | --- |
| `README.md` | how do I use it? | the interface does |
| `claude.md` | where is the project up to? | the work does |
| `adrs/` | why is it like this? | never — superseded instead |

**Date** is when the record was written, not when the decision was taken.
ADR-0001 to ADR-0023 were all written on 2026-08-08, reconstructing decisions
made over the preceding days from the commits that carry them — so several cite
evidence from before their own date, and one or two from after. Where the timing
matters, the Context says so.
# ADR-0001: The workbook is never modified

**Status:** Accepted

**Date:** 2026-08-26

## Context

EDJAS exists because organisations run real work in spreadsheets and are not going
to stop. Every tool that has tried to improve that situation by asking people to
change their workbook — add a macro, insert a marker column, restructure a sheet so
it is machine-readable — has run into the same wall: the workbook belongs to someone
else, it is in use, and it is often the authoritative copy. A tool that needs the
file changed before it can help is a tool that needs permission, a change-control
conversation, and a person willing to take responsibility if the change breaks a
formula three sheets away.

There is a sharper version of the problem for published data. Government
statistical releases are the exact case EDJAS was built for, and they are immutable
by definition: the published file is the record. Anything that writes to it is not
merely inconvenient, it is wrong.

## Decision

EDJAS only ever reads a spreadsheet. It opens the file, takes the values it was
asked for, and closes it; there is no code path that writes to a workbook, and no
option to enable one. Everything EDJAS needs to know about an extraction comes from
somewhere other than the file being extracted. This is the property the whole design
is arranged around, not a feature of it.

## Consequences

Anything EDJAS needs to be told has to be told somewhere else, which is what forces
[ADR-0002](0002-extraction-is-described-beside-the-spreadsheet.md) and everything
downstream of it. A user cannot annotate a workbook to make it easier to extract;
they must describe the extraction separately, which is more work the first time and
less work every time after.

It also means EDJAS cannot repair what it reads. A workbook with values in the wrong
place stays that way, and the specification has to accommodate it. Several of the
uglier corners of the worked examples — padding cells, nil markers, footnotes inside
the data range — exist because fixing them at source is not an option available to us.

The promise is worth nothing if it is only intended, so it is tested: rendering a
report asserts the source file is byte-for-byte unchanged afterwards.

## Alternatives considered

**Write extraction instructions into the workbook.** This was the original design and
is recorded as rejected in [ADR-0002](0002-extraction-is-described-beside-the-spreadsheet.md).

**Offer an opt-in write mode** — normalise the sheet, add named ranges, tidy headings.
Rejected because the guarantee is only useful if it is absolute. "EDJAS never writes
to your file, unless you asked it to, in which case check which mode you are in" is
not a guarantee anyone can rely on, and the audience least able to check is the one
that most needs the assurance.
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
# ADR-0004: An expression is a reference followed by a pipeline

**Status:** Accepted

**Date:** 2026-08-26

## Context

Extracting a range rarely produces the shape a consumer wants. A table comes back as
rows of rows when what is needed is a list of objects; a column of numbers arrives as
text; a block is oriented the wrong way round. Something has to reshape it, and the
question was where that something lives and what it looks like.

The first attempt put the function first: `[records Sales]`, later written
`[f name]`. It read backwards. A reader parsing `[records Sales]` meets the verb
before knowing what it applies to, and chaining two transformations means nesting —
`[records [transpose Grid]]` — which inverts the reading order relative to the order
of operations and gets worse with every stage added.

## Decision

An extraction expression is a single range reference followed by zero or more
functions, separated by `|` and applied left to right after the value is read:
`[Grid | transpose | records]`. The brackets around the whole expression choose the
extraction shape — a bare reference yields a scalar, `[ref]` a list or list of rows,
`{ref}` a two-column range as an object — and the pipeline transforms whatever that
produced. A function may take arguments after its name: a number, a double-quoted
string, or a bare word read as another range reference.

## Consequences

Expressions read in the order they execute, and adding a stage appends rather than
nests. This is the property that makes `[Sales | records | round 3]` legible to
someone who has never seen the syntax before, and it is worth the change it cost.

Changing the notation on 2026-07-22 was breaking, and it invalidated every
specification written against the earlier form. There were no external users, so the
old syntax was removed rather than carried.

The pipe character is now significant inside expressions, which is why the tokeniser
has to be quote-aware: `[Tags | join ", "]` contains a pipe inside a string argument
that must not split the expression. That tokeniser later had to learn Excel's
single-quote convention for sheet names too, and the two quoting rules living in one
grammar is the least elegant part of the language.

Pipelines operate on one key at a time and cannot see the rest of the document, which
is a real limit — the reason a document-level stage keeps being proposed, and is
deferred in [ADR-0014](0014-transform-and-load-are-deferred.md).

## Alternatives considered

**Function-first, `[records Sales]`.** Implemented first and replaced. Reads
backwards; nests badly.

**Method syntax, `Sales.records().round(3)`.** Familiar to programmers, but it
suggests objects and arbitrary chaining to an audience that mostly does not write
code, and it invites the expectation that any Python method will work.

**No transformation at all**, leaving reshaping to the consumer. Rejected because the
most common need — turning a header row and data rows into records — is universal
enough that every user would write the same code, and because the shape of the JSON
is exactly what a specification ought to be choosing.
# ADR-0005: A specification cannot execute arbitrary code

**Status:** Accepted

**Date:** 2026-08-26

## Context

The pipeline in [ADR-0004](0004-an-expression-is-a-reference-and-a-pipeline.md) needs
to resolve a name like `records` to something that runs. The easy implementations all
lead somewhere bad: `eval` on the expression, importing a module named in the
specification, or a plugin hook that loads whatever is on the path.

What makes this more than a theoretical concern is the design's own encouragement to
share. One specification is meant to serve many workbooks and many people; the whole
argument for a separate file is that it can be committed, reviewed and passed around.
A format that is worth sharing is a format that will be received from someone else,
and it should not be possible to be harmed by reading a colleague's extraction.

## Decision

Functions are looked up in a fixed registry, never evaluated. A specification can
invoke only the built-in functions and any the calling program supplies through the
`functions=` argument to `read_spec`; there is no syntax that reaches Python, no
import mechanism, and no expression evaluation. An unknown function name is an error
naming the available ones.

## Consequences

A specification is inert. It can be read, mailed, committed and run without the
recipient auditing it first, and that is a property worth stating plainly in the
documentation — which it is, on the function reference page.

The cost is that the language cannot grow features by delegation. Anything a
specification needs to do must either be a built-in or be passed in by a program that
is already trusted, so genuinely novel transformations require writing Python rather
than a clever expression. That is the intended trade: extension happens at the
boundary, in the host program, not in the data file.

It also constrains what a future `[load]` section could look like, since a
destination that can write arbitrary paths or reach the network would weaken exactly
this guarantee — a principal reason it is deferred in
[ADR-0014](0014-transform-and-load-are-deferred.md).

## Alternatives considered

**`eval` on a restricted namespace.** Every published attempt at sandboxing Python
expressions has eventually been escaped, and the failure is silent and total. Not
worth the expressiveness.

**Entry-point plugins**, so third parties could register functions by installing a
package. Deferred rather than refused, but it moves the trust boundary from "the
program that calls EDJAS" to "whatever is installed in the environment", which is a
much harder thing to reason about and a much easier thing to get wrong.

**A small expression evaluator** for arithmetic and comparisons. Rejected on the same
grounds recorded in [ADR-0014](0014-transform-and-load-are-deferred.md): it is the
first step to reimplementing a programming language badly, and it dissolves the
one-sentence security claim into a list of caveats.
# ADR-0006: Formula cells yield the value the application cached

**Status:** Accepted

**Date:** 2026-08-26

## Context

A spreadsheet cell holding `=SUM(B2:B12)` has two readings: the formula text and the
number it last evaluated to. openpyxl exposes both, and EDJAS originally returned the
formula, which meant extracting a total produced the string `"=SUM(B2:B12)"` rather
than a number.

That is almost never what anyone wants. The audience for EDJAS is downstream of the
spreadsheet — dashboards, reports, other systems — and downstream consumers want the
figure. Nobody building a report on retail sales needs to know how the total was
arrived at; they need the total.

The complication is that the cached value is a fact about the file, not about the
formula. It is whatever the application wrote the last time it saved, and a workbook
produced by a tool that never evaluates formulas has no cache at all.

## Decision

Workbooks are opened with `data_only=True`, so a formula cell yields the value the
application cached when it last saved the file. EDJAS does not evaluate formulas and
has no formula engine. A cell whose formula has never been evaluated reads as `null`.

## Consequences

Totals, averages and derived columns extract as numbers, which is what makes EDJAS
useful against real workbooks at all; almost every government file in the worked
examples is full of computed cells.

The `null` case is a genuine trap and is documented as one. A workbook generated
entirely by another program — including, awkwardly, one generated by openpyxl — has
no cached values, so an extraction that works perfectly against a file saved by Excel
returns nulls against a file that looks identical but was written by a script. The
test suite has to inject cached values into the file's XML by hand to exercise the
working case, because openpyxl cannot write one.

EDJAS also has no way to detect a stale cache. If someone edits values without
recalculating, the extracted figures are the old ones and nothing indicates it.

## Alternatives considered

**Return the formula text.** The original behaviour. Useful to perhaps one user in a
hundred, and useless to the rest.

**Return both**, as an object carrying value and formula. Rejected: it complicates
every consumer to serve a rare need, and the pipeline functions would all have to
learn about the wrapper.

**Evaluate formulas ourselves.** Rejected decisively. Implementing Excel's function
library and evaluation semantics is a project several times the size of EDJAS, and
getting it subtly wrong would produce numbers that disagree with the workbook — far
worse than reporting a null.
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
# ADR-0011: The documentation site is built with the project's own templates

**Status:** Accepted

**Date:** 2026-08-26

## Context

Around four thousand words of documentation had accumulated across three Markdown
files, and the root README was doing seven jobs at once — pitch, language reference,
function catalogue, tutorial, API guide, showcase and architecture tour. Each was done
well; done all at once in one scrolling file, no reader was served. Meanwhile four
hand-drawn architecture diagrams were *linked* rather than embedded, so on GitHub they
rendered as bare hyperlinks, and seven rendered reports sat in the repository with no
route for a visitor to see them.

A documentation site was the answer. The question was what to build it with. A
generator would have brought search, navigation and a theme for nothing, at the cost
of a toolchain and a look shared with every other Python project.

The project already had a Jinja2 rendering stack with a stylesheet, because that is
what produces the example reports.

## Decision

The site is generated by `docs/build_site.py`, and its pages extend the very
`base.html` that the example reports extend, so a documentation page and a generated
report are the same document with different content. `base.html` gained one empty
`head_extra` block for the purpose; site-only styling lives in a separate stylesheet
and never in that file, which is inlined into every report and ships in the wheel.
Markdown is rendered with Python-Markdown, in a docs-only dependency group that never
reaches a user.

## Consequences

The site and the reports share a visual language by construction rather than by
discipline, and the claim that the documentation is built by the project's own
machinery is literally true.

Editing a template used by a shipped artefact is the obvious hazard, and it is
contained rather than trusted: a test asserts the seven committed reports are
byte-for-byte what the code produces, so any change to `base.html` that alters a
report fails the suite. That test also creates an obligation — a deliberate template
change now means regenerating the committed showcase before committing.

Inlining the diagrams turned out to be the sharp edge. Inline SVG is *document*-scoped:
all four diagrams defined `<marker id="arrow">`, and they shared class names with
differing values, so naive inlining silently gave three of them the wrong typography,
and one diagram's `.line` rule leaked onto the ledger rows `base.html` styles. Each
diagram's ids are namespaced and its selectors scoped to its own root. The fix has a
trap of its own, recorded in the code: a blanket search for rows would hoist a
sub-table's rows out of a cell, so the walk descends only into known wrappers.

What the site does not get is search, and a hand-written generator is a thing to
maintain. Navigation is a hand-ordered manifest, which is an editorial decision at the
scale of fifteen pages and would be a chore at fifty.

## Alternatives considered

**MkDocs with Material.** Search, navigation, mobile layout and dark mode for free,
and content stays portable Markdown. Rejected because the site would look like every
other Python project's, and because the dogfooding argument — a tool that turns data
into readable documents, documented by a site it built — is worth more here than
search is.

**Sphinx.** Heavier, and its output would have needed as much theming work as writing
the generator did.

**Minimal GitHub Pages with a Jekyll theme.** Least effort and least control,
particularly over the gallery, which is the part of the site doing the persuading.
# ADR-0012: The site composes canonical files, and reads code-derived content from the code

**Status:** Accepted

**Date:** 2026-08-26

## Context

Building the site meant deciding where its words come from, and the repository already
had a warning about that: the catalogue describing the seven examples existed in *four*
places — the root README, the companion package's README, `ATTRIBUTION.md` and a module
docstring — at four lengths and in four registers. They agreed on the facts that day.
They would not have kept agreeing.

Two of those files cannot simply be absorbed. `ATTRIBUTION.md` is a licensing document
that ships inside the wheel; a copy of it on a web page that drifts is an
Open Government Licence compliance problem, not a formatting one. The companion's
README is its PyPI description.

There was also a temptation in the other direction. Having decided the site should be
built by EDJAS itself, the appealing claim was that *everything* on it came out of a
spreadsheet — which would have meant putting the list of built-in functions in a
workbook, where it could disagree with the registry that actually implements them.

## Decision

Pages are composed from canonical files rather than copies: the colophon renders the
real `ATTRIBUTION.md` that ships in the package, the essay renders the companion's own
README, the tutorial embeds `examples/example.toml` verbatim, and each gallery page
takes its introduction from the comment header of the specification it describes.
Genuinely editorial tables — the gallery cards, the summary tables — are held in
`docs/site.xlsx` and extracted at build time by an ordinary EDJAS specification.
Anything derived from code is read from the code: the function reference comes from
the registry, API signatures from introspection, command-line help from the parsers.

## Consequences

Nothing is duplicated by hand, so nothing can quietly fall out of step, and the
four-way catalogue collapsed to one canonical home with the others pointing at it.
The licensing text a reader sees is the licensing text that ships.

The dogfooding claim is checkable rather than rhetorical, and — importantly — *bounded*.
A reference page that can disagree with the code is worse than no reference page, so
the boundary is stated on the colophon itself rather than left implicit.

Where a table exists in both the workbook and the code, tests assert the two agree, and
the workbook is generated by a script that derives those columns from the live
registries, so drift is impossible at generation time and caught at test time if
someone hand-edits the file.

The costs are real. Composition means a page's content is assembled from files in
several directories, so its source is less obvious than a single Markdown file would
be; the manifest has to support demoting headings and skipping a title so composed
parts nest correctly. And the site build now depends on the companion package's Python
API, so a refactor there breaks the documentation build — acceptable, since it is the
same repository and the same commit.

## Alternatives considered

**Copy the prose into the site and keep it in step by hand.** What the repository was
already doing in four places, with the drift not yet visible only because the files
were young.

**Put everything in the workbook**, maximising the dogfooding claim. Rejected: the
function list belongs to `DEFAULT_FUNCTIONS`, and a spreadsheet copy of it is a
promise to eventually document a function that does not exist.

**Generate everything from code, including the editorial framing.** Would have meant
inventing a place in the source for phrases like "a council revenue account", which is
prose, not data.
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
# Architectural decision records

Why EDJAS is the way it is. The [template](0000-template.md) explains the form and the
rules; the short version is that a record is **immutable once accepted**, and a decision
that changes gets a new record superseding the old one rather than an edit.

These fourteen were written on 2026-08-26, reconstructing decisions taken over the
preceding months from the commits that carry them. Several therefore cite evidence from
before their own date; where the timing matters, the Context says so.

| | Decision | Status |
| --- | --- | --- |
| [0001](0001-the-workbook-is-never-modified.md) | The workbook is never modified | Accepted |
| [0002](0002-extraction-is-described-beside-the-spreadsheet.md) | Extraction is described beside the spreadsheet, not inside it | Accepted |
| [0003](0003-specifications-are-toml.md) | Specifications are TOML | Accepted |
| [0004](0004-an-expression-is-a-reference-and-a-pipeline.md) | An expression is a reference followed by a pipeline | Accepted |
| [0005](0005-a-specification-cannot-execute-code.md) | A specification cannot execute arbitrary code | Accepted |
| [0006](0006-formula-cells-yield-the-cached-value.md) | Formula cells yield the value the application cached | Accepted |
| [0007](0007-unrepresentable-references-are-refused.md) | References with no single rectangle are refused, not guessed | Accepted |
| [0008](0008-table-names-are-references.md) | Excel Table names are references, with defined names taking precedence | Accepted |
| [0009](0009-rendering-is-optional-and-presentation-is-explicit.md) | Rendering is optional, and presentation choices are explicit | Accepted |
| [0010](0010-worked-examples-are-a-separate-distribution.md) | The worked examples ship as a separate distribution | Accepted |
| [0011](0011-the-documentation-site-uses-the-projects-own-templates.md) | The documentation site is built with the project's own templates | Accepted |
| [0012](0012-the-site-composes-canonical-files.md) | The site composes canonical files, and reads code-derived content from the code | Accepted |
| [0013](0013-ods-is-read-by-a-standard-library-reader.md) | OpenDocument files are read by a standard-library reader | Accepted |
| [0014](0014-transform-and-load-are-deferred.md) | A document-level `[transform]` and a `[load]` section are deferred | Deferred |

## The through-line

[0001](0001-the-workbook-is-never-modified.md) forces
[0002](0002-extraction-is-described-beside-the-spreadsheet.md), which needs a format
([0003](0003-specifications-are-toml.md)) and a language
([0004](0004-an-expression-is-a-reference-and-a-pipeline.md)) that must not become a
way to run code ([0005](0005-a-specification-cannot-execute-code.md)). Everything after
that is either fidelity to what the workbook actually contains
([0006](0006-formula-cells-yield-the-cached-value.md),
[0007](0007-unrepresentable-references-are-refused.md),
[0008](0008-table-names-are-references.md),
[0013](0013-ods-is-read-by-a-standard-library-reader.md)) or a decision about how far
the project should reach beyond reading
([0009](0009-rendering-is-optional-and-presentation-is-explicit.md),
[0010](0010-worked-examples-are-a-separate-distribution.md),
[0014](0014-transform-and-load-are-deferred.md)).

Refusing to guess recurs often enough to be a house rule: a union range, a two-sheet
range, a 26-hour duration and an unrecalculated formula are each reported rather than
approximated.
