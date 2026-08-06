Run it against the workbook to see all of it at once:

```
edjas examples/example.xlsx examples/example.toml
```

## 7. Render a report

Extraction gives you data; the optional reporting layer turns it into a document. It needs
Jinja2, which arrives with the `demo` extra.

```python
from edjas import render_report

html = render_report(
    "examples/example.xlsx", "examples/example.toml",
    template="report.html", templates_dir="templates",
)
```

The extracted dictionary reaches the template as `data`, and any extra keyword arguments
become further template variables. Rendering only reads the workbook, exactly as plain
extraction does.

## Where next

  - The [gallery](site:gallery/index.html) applies all of this to seven real government
    workbooks, with the reports they produce.
  - The [specification language](site:reference/spec-language.html) covers references,
    quoting and formula cells in full.
  - The [Python API](site:reference/api.html) covers `read_spec`, `render_report` and
    supplying your own functions.
