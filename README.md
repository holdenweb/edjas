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

It lets you add data specifications to any existing spreadsheet by creating
named ranges in the spreadsheet. By default EDJAS will look for a range
name `Parameters` as its starting point, although this can be overridden on the command line.
This range should be precisely two columns wide, and EDJAS
treats the left-hand column as names and the right-hand column as values.
Normally, the values are used literally after extraction from the spreadsheet.
Two formats for the value are given special treatment.

  - **`[range-name]`**: the named range is exported as a JSON list or, if  it's two-dimensional a list of row lists.
  - **`{range_name}`**: The named range, which must be two columns wide, becomes a JSON object where the left-hand column specifies
    the names and the right-hand column specifies the values.

## Applying functions

Either markup form may name a function to transform the extracted value:
`[f name]` yields the result of applying `f` to the value `[name]` would produce,
and `{f name}` yields `f` applied to the value `{name}` would produce. The
function name and the range name are separated by a space (Excel range names
never contain spaces, so this is unambiguous).

For example, if `Sales` is a table whose first row holds column headings,
`[records Sales]` turns it into a list of JSON objects — one per data row.

Functions are resolved from a fixed, built-in registry — spreadsheets cannot run
arbitrary code. The functions shipped by default are:

| Function | Applied to | Result |
|----------|------------|--------|
| `records` | `[table]` | first row is headings; remaining rows become a list of objects |
| `columns` | `[table]` | first row is headings; columns become a `{heading: [values]}` object |
| `transpose` | `[table]` | swaps rows and columns |
| `flatten` | `[table]` | flattens nested rows into a single list |
| `keys` / `values` / `items` | `{object}` | the object's keys, values, or `[key, value]` pairs |
| `invert` | `{object}` | swaps keys and values |
| `int` / `float` / `str` | either | coerces every value to that type |
| `round2` | either | rounds every floating-point value to two decimal places |
| `isodate` | either | formats date/time values as ISO-8601 strings |

When using EDJAS as a library, `read_file(path, functions={...})` adds your own
functions to (and can override) the built-ins:

```python
from edjas import read_file
read_file("data.xlsx", functions={"upper": lambda v: [s.upper() for s in v]})
```

Date and time cells are serialised as ISO-8601 strings automatically.

The parameter details are used to extract data from the spreadsheet, which is then sent to standard output as JSON.

![Parameter specifications in EDJAS](https://raw.githubusercontent.com/holdenweb/edjas/main/images/parameters.png "Parameter specifications in EDJAS")

In the example shown, the `version` key has a dict value, and in that dict the `number` key has a value of "1.0.2".
The version number can therefore be referenced in the JSON output as `version.number`. The output from this example is shown below.


![Parameter data extracted from a spreadsheet](https://raw.githubusercontent.com/holdenweb/edjas/main/images/json.png "The parameter data")

A demonstration of the system can be found at [https://github.com/holdenweb/edjas-demo](https://github.com/holdenweb/edjas-demo).

This is particularly useful for audiences that have an interest in only a
limited number of features from a possibly quite large spreadsheet.
More generally, JSON is such a widely used format that spreadsheet data can
be re-used in a wide range of systems as appropriate.