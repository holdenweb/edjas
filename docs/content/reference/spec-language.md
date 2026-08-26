A specification is a TOML file with an `[extract]` table mapping the output keys you want
to the values to pull from the spreadsheet:

```toml
[extract]
title  = "Summary!B2"                     # a single cell    -> scalar
hours  = "{Hours}"                        # a 2-column range -> object
prices = "{Prices | int}"                 # object, values coerced to int
sales  = "[Sales | records]"              # a table          -> list of objects
people = "[Grid | transpose | records]"   # a pipeline of transforms
```

Each value is an **extraction expression**: a reference to part of the workbook,
optionally followed by a [pipeline of functions](site:reference/functions.html).

The same expressions read Excel workbooks (`.xlsx`) and OpenDocument spreadsheets
(`.ods`) alike — EDJAS chooses by file extension and nothing in a specification depends
on which it is. That matters for public-sector data in particular, where some
publications are released only as ODS: the MHCLG local-authority revenue outturn tables,
for instance, have no Excel edition at all. An ODS file is a zip of XML, so reading one
needs no dependency beyond the standard library. Named ranges defined in an ODS file are
resolved exactly as Excel defined names are.
