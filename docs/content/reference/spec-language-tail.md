Named ranges and tables are recommended over bare cell references: they survive layout
changes. This also fits the way well-produced government and statistical spreadsheets are
built, where the published data is laid out as named Excel Tables.

Resolution order is **defined name, then table name, then a raw A1 reference**. Excel
keeps defined names and table names in one namespace and forbids collisions, so the order
rarely matters in practice.

## Quoting sheet names

A sheet name containing spaces or punctuation must be wrapped in single quotes, exactly as
Excel writes it:

```toml
title  = "'Cover Sheet'!A1"
answer = "'Bob''s Data'!C3"     # an embedded apostrophe is doubled
```

The two kinds of quote mean different things inside an expression, and it is worth being
precise about them:

  - **Double quotes** delimit a *string argument* to a function. The quotes are stripped,
    and whitespace and `|` inside them are literal — which is how `[Tags | join ", "]`
    passes a comma followed by a space.
  - **Single quotes** are Excel's *sheet-name* quoting. The quotes are kept, so the sheet
    resolver can strip them itself, and a doubled `''` is an escaped apostrophe rather
    than a closing quote.

Sheet names and defined names are matched **case-insensitively**, as Excel does.

## Formula cells

A cell containing a formula yields its computed value, not the formula text, so extracting
a total or an average works as you would expect. EDJAS reads the value Excel cached the
last time it saved the workbook.

One consequence is worth knowing: a workbook that has never been recalculated in Excel —
one produced entirely by another tool, say — has no cached value, and such a cell reads as
`null`.

## What is not supported

Multi-area (union) references — `A1:A5,C1:C5`, or a named range defined as one — are
deliberately unsupported. A union has no single rectangular shape, so it has no
unambiguous mapping to JSON. EDJAS reports it clearly rather than guessing; give each
block its own key instead, which produces better-named output anyway.
