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
