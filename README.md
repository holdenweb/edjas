# EDJAS: Extract Data in JSON from Any Spreadsheet

- Sources at https://github.com/holdenweb/edjas
- Demonstration code at https://github.com/holdenweb/edjas-demo

This project is an attempt to help organisations that insist on managing
their businesses, or major aspects thereof, using spreadsheets.
Many articles have been written on the limitations of spreadsheet technology.
If you have any doubts then look at the "The Problem with Spreadsheets"
section of this [LinkedIn
article](https://www.linkedin.com/pulse/spreadsheets-inadequate-effective-management--gjsse/).
Some large organisations are now
[providing advice](https://www.gov.uk/guidance/creating-and-sharing-spreadsheets)
— although in many cases better advice might be:
_stop using spreadsheets for that_!

Rather than try to change the way people do business (imagine "If I Ruled the
World" playing softly in the background), EDJAS is intended to help people extract
that locked-up data more effectively, in simple and easy-to-understand ways
that don't affect existing workflows.

EDJAS leaves your spreadsheet **completely untouched**. Instead of adding anything
to the workbook, you write a small **specification file** — a TOML document —
describing what to extract. One spec can serve many workbooks, and one workbook can
have many specs, each tailored to a different audience.

## The specification file

A spec is a TOML file with an `[extract]` table mapping the output keys you want to
the values to pull from the spreadsheet:

```toml
[extract]
title  = "Summary!B2"                    # a single cell   -> scalar
hours  = "{Hours}"                        # a 2-column range -> object
prices = "{Prices | int}"                 # object, values coerced to int
sales  = "[Sales | records]"              # a table -> list of objects
people = "[Grid | transpose | records]"   # a pipeline of transforms
```

Each value is an **extraction expression**. There are three forms:

  - **`ref`** (a plain reference) — the value of a single cell.
  - **`[ref]`** — the range as a JSON list, or a list of row-lists if it is
    two-dimensional.
  - **`{ref}`** — a two-column range as a JSON object (left column names, right
    column values).

A `ref` is either a **named range** or an A1-style cell range (`D3:E9`), optionally
qualified with a sheet name (`Summary!D3:E9`). Named ranges are recommended: they
survive layout changes, whereas bare cell references do not.

## Transforming values with functions

Any expression may append a **pipeline** of functions, separated by `|`, applied left
to right after extraction — so `[Grid | transpose | records]` transposes the range,
then builds objects from it. Functions come from a fixed, built-in registry (no
arbitrary code runs). The built-ins:

| Function | Typical input | Result |
|----------|---------------|--------|
| `records` | `[table]` | first row is headings; remaining rows become a list of objects |
| `columns` | `[table]` | first row is headings; columns become a `{heading: [values]}` object |
| `transpose` | `[table]` | swaps rows and columns |
| `flatten` | `[table]` | flattens nested rows into a single list |
| `keys` / `values` / `items` | `{object}` | the object's keys, values, or `[key, value]` pairs |
| `invert` | `{object}` | swaps keys and values |
| `int` / `float` / `str` | any | coerces every value to that type |
| `round2` | any | rounds every floating-point value to two decimal places |
| `isodate` | any | formats date/time values as ISO-8601 strings |

### Function arguments

A function may take arguments after its name, separated by spaces. An argument is a
**number** (`2`), a **double-quoted string** (`", "`), or a **bare word**, which is
read as another range reference. The extracted value is always passed as the first
argument, so `[Price | round 2]` means `round(Price, 2)`. (Grouping parentheses are
reserved for a possible future extension and are not yet supported.)

## Usage

From the command line — pass the spreadsheet and the spec; JSON goes to standard
output:

```
edjas data.xlsx report.toml
```

As a library, `read_spec` returns the extracted data as a Python dict. Pass
`functions={...}` to add your own functions to (or override) the built-ins; each
receives the extracted value first, then any arguments:

```python
from edjas import read_spec
data = read_spec("data.xlsx", "report.toml",
                 functions={"join": lambda v, sep: sep.join(v)})
# ... lets the spec use:  tags = "[Tags | join \", \"]"
```

Date and time cells are serialised as ISO-8601 strings automatically.

A demonstration of the system can be found at [https://github.com/holdenweb/edjas-demo](https://github.com/holdenweb/edjas-demo).

This is particularly useful for audiences that have an interest in only a
limited number of features from a possibly quite large spreadsheet.
More generally, JSON is such a widely used format that spreadsheet data can
be re-used in a wide range of systems as appropriate.