# EDJAS: Extract Data in JSON from Any Spreadsheet

- Sources at https://github.com/holdenweb/edjas

This project helps organisations that manage their businesses, or major aspects thereof,
using spreadsheets. Many articles have been written on the limitations of spreadsheet
technology — see the "The Problem with Spreadsheets" section of this [LinkedIn
article](https://www.linkedin.com/pulse/spreadsheets-inadequate-effective-management--gjsse/)
— and some large organisations now
[provide advice](https://www.gov.uk/guidance/creating-and-sharing-spreadsheets), although
in many cases better advice might be: _stop using spreadsheets for that_!

Rather than try to change the way people do business, EDJAS helps people extract that
locked-up data effectively, in simple and easy-to-understand ways that don't affect
existing workflows.

EDJAS leaves your spreadsheet **completely untouched**. Instead of adding anything to the
workbook, you write a small **specification file** — a TOML document — describing what to
extract. One spec can serve many workbooks, and one workbook can have many specs, each
tailored to a different audience.

Excel workbooks (`.xlsx`) and OpenDocument spreadsheets (`.ods`) are both read, using the
same expressions — which matters for public-sector data, where whole publications are
released in ODS and nothing else. Neither format needs an extra dependency.

```
pip install edjas
```

## The specification file

A spec is a TOML file with an `[extract]` table mapping the output keys you want to the
values to pull from the spreadsheet:

```toml
[extract]
title  = "Summary!B2"                     # a single cell    -> scalar
hours  = "{Hours}"                        # a 2-column range -> object
prices = "{Prices | int}"                 # object, values coerced to int
sales  = "[Sales | records]"              # a table          -> list of objects
people = "[Grid | transpose | records]"   # a pipeline of transforms
```

Each value is an **extraction expression**. There are three forms:

  - **`ref`** — the value of a single cell.
  - **`[ref]`** — the range as a JSON list, or a list of row-lists if two-dimensional.
  - **`{ref}`** — a two-column range as an object (left column names, right column values).

A `ref` may be a **named range** (`Prices`), the name of an **Excel Table**
(`RevisionTriangles_Table1`), or an **A1-style cell range** (`D3:E9`), optionally
sheet-qualified (`Summary!D3:E9`). Named ranges and tables are recommended: they survive
layout changes. A sheet name containing spaces or punctuation is single-quoted exactly as
Excel writes it — `'Cover Sheet'!A1`, or `'Bob''s Data'!C3`.

A cell containing a **formula** yields its computed value, not the formula text. EDJAS
reads the value Excel cached when it last saved; a workbook never recalculated in Excel
has no cached value, and such a cell reads as `null`.

## Transforming values with functions

Any expression may append a **pipeline** of functions, separated by `|` and applied left
to right, so `[Grid | transpose | records]` transposes the range then builds objects from
it. Functions come from a fixed, built-in registry — a spec can only invoke registered
functions and **never executes arbitrary code**.

| Function | Typical input | Result |
|----------|---------------|--------|
| `records` | `[table]` | first row is headings; remaining rows become a list of objects |
| `columns` | `[table]` | first row is headings; columns become a `{heading: [values]}` object |
| `transpose` | `[table]` | swaps rows and columns |
| `flatten` | `[table]` | flattens nested rows into a single list |
| `keys` / `values` / `items` | `{object}` | the object's keys, values, or `[key, value]` pairs |
| `invert` | `{object}` | swaps keys and values |
| `int` / `float` / `str` | any | coerces every value to that type |
| `round` | any | rounds floating-point values; takes the number of places (default 2) |
| `isodate` | any | formats date/time values as ISO-8601 strings |

A function may take arguments after its name: a **number** (`2`), a **double-quoted
string** (`", "`), or a **bare word** read as another range reference. The extracted value
is passed first, so `[Price | round 2]` means `round(Price, 2)`.

## Usage

From the command line — pass the spreadsheet and the spec; JSON goes to standard output:

```
edjas data.xlsx report.toml
```

To find out what a workbook offers before writing a spec, ask it:

```
edjas data.xlsx --list-names
```

That lists the named ranges and, sheet by sheet, the Excel Tables, with the shape of
each and the extraction form that fits it. Names that cannot be used as references —
those whose cells were deleted, leaving Excel's `#REF!`, and those holding a constant
rather than a range — are counted with the reason; `--verbose` lists them individually.

As a library, `read_spec` returns the extracted data as a Python dict. Pass
`functions={...}` to add your own functions to (or override) the built-ins; each receives
the extracted value first, then any arguments:

```python
from edjas import read_spec
data = read_spec("data.xlsx", "report.toml",
                 functions={"join": lambda v, sep: sep.join(v)})
# ... lets the spec use:  tags = "[Tags | join \", \"]"
```

Date and time cells are serialised as ISO-8601 strings automatically.

A sample workbook and a spec exercising every construct ship in `examples/`:

```
edjas examples/example.xlsx examples/example.toml
```

## Rendering HTML reports

Extraction gives you a data structure; turning it into a readable document is a separate,
optional step. `render_report` feeds the extracted data into a
[Jinja2](https://jinja.palletsprojects.com/) template. The `demo` extra brings in Jinja2
together with seven worked examples — real, unmodified UK government workbooks published
under the Open Government Licence, each with its spec and the report it produces:

```
pip install edjas[demo]
edjas-examples                      # write all seven reports to the current directory
edjas-examples --example slgfs      # just one
```

```python
from edjas import render_report

html = render_report("data.xlsx", "report.toml",
                     template="report.html", templates_dir="templates")
```

Rendering only *reads* the workbook, exactly as plain extraction does.

## Documentation

The full documentation — a guided introduction, the complete specification-language and
API reference, a gallery of the seven worked examples with their live reports, and the
architecture diagrams — is built from this repository:

```
pip install edjas[demo] && python docs/build_site.py --serve
```

The [essay accompanying the examples](https://github.com/holdenweb/edjas/blob/main/edjas-examples/README.md)
discusses what the demonstrations show and where the approach applies; it is written for
people who work with spreadsheets rather than for programmers.

## Architecture

Four diagrams describe how EDJAS is put together, zooming in a level at a time. Each is a
self-contained SVG in the `images/` directory, and all four are shown together on the
documentation site's architecture page.

  - [**C4 system context**](https://github.com/holdenweb/edjas/blob/main/images/c4-context.svg)
    — the setting: the spreadsheet maintainer who carries on as before, the analyst who
    authors specs, and the systems the extracted JSON feeds.
  - [**C4 container diagram**](https://github.com/holdenweb/edjas/blob/main/images/c4-container.svg)
    — the command-line and library containers, the files they read, and where the JSON
    ends up.
  - [**C4 component diagram**](https://github.com/holdenweb/edjas/blob/main/images/c4-component.svg)
    — inside the library: the spec loader, expression evaluator, pipeline parser, workbook
    reader, pipeline executor and function registry.
  - [**Internal structure**](https://github.com/holdenweb/edjas/blob/main/images/internal-structure.svg)
    — the modules that make up the package, and how a single extraction run flows through
    them from the spec and workbook to standard output.
