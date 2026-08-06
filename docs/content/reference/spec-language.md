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
