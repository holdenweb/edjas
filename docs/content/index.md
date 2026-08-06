## Why it works this way

Many organisations run significant parts of their business in spreadsheets. Rather than
try to change that — and plenty of people have tried — EDJAS helps you get the data out,
in a way that leaves the existing workflow completely alone.

Nothing is added to the workbook: no macros, no marker columns, no restructuring. The
description of what to extract lives in its own small TOML file, which means one
specification can serve many workbooks, and one workbook can have many specifications,
each tailored to a different audience.

That last point is the useful one in practice. A large spreadsheet usually contains far
more than any single reader wants; a specification names just the part that matters to
them. And because JSON is so widely understood, whatever you extract can feed dashboards,
databases, APIs, site builders or documents without further translation.
