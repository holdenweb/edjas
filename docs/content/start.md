This walks from nothing to a JSON extract and an HTML report, using a sample workbook that
ships with EDJAS. It takes about ten minutes.

## 1. Install

```
pip install edjas
```

That gives you the `edjas` command and the library. To render HTML reports as well, and to
get the seven worked examples, install the demo extra instead:

```
pip install edjas[demo]
```

## 2. The sample workbook

The repository ships `examples/example.xlsx`, a small café report spread over four sheets.
Its named ranges cover every shape you are likely to meet:

| Named range | Where | Shape |
|---|---|---|
| `Title`, `PeriodEnd`, `AvgSpend` | `Summary!B2:B4` | single cells |
| `Hours`, `Prices`, `Covers`, `Codes` | `Data` | two-column name/value ranges |
| `Tags` | `Data!M1:M3` | a single column |
| `Sales` | `Sales!A1:C4` | a table with a heading row |
| `Staff` | `Staff!A1:C2` | one field per row, so it needs transposing |

Here is the workbook itself, with each named range shaded in a colour:

![The example workbook, with each named range shaded](https://raw.githubusercontent.com/holdenweb/edjas/main/images/example_workbook.png)

## 3. Write a specification

Create `first.toml` next to the workbook:

```toml
[extract]
title = "Title"
tags  = "[Tags]"
hours = "{Hours}"
```

Three keys, one per [form](site:reference/spec-language.html): a scalar, a list and an
object.

## 4. Run it

```
edjas examples/example.xlsx first.toml
```

```json
{
  "title": "Riverside Cafe",
  "tags": ["Vegan", "Gluten-free", "Dairy-free"],
  "hours": {"Monday": "07:00-20:00", "...": "...", "Sunday": "Closed"}
}
```

The workbook has not been touched — it was opened read-only and closed again.

## 5. Reshape on the way out

Raw ranges are rarely the shape you want. Append a
[pipeline](site:reference/functions.html) to reshape as you extract:

```toml
[extract]
sales   = "[Sales | records]"              # rows become objects, keyed by the header row
staff   = "[Staff | transpose | records]"  # fields-down-the-page, so transpose first
covers  = "{Covers | int}"                 # this sheet holds its numbers as text
average = "AvgSpend | round 2"             # 8.75, from a cell holding 8.7451
period  = "PeriodEnd | isodate"            # "2026-03-31"
```

## 6. The whole thing

`examples/example.toml` is a worked specification exercising every construct, and it is
reproduced here in full — this is the actual file, not a copy of it:
