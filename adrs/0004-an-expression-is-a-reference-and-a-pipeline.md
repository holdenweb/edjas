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
