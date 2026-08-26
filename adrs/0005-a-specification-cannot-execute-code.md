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
