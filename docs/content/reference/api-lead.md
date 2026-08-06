`read_spec` is the main entry point: give it a workbook and a specification and it returns
a plain Python dictionary. `render_report` adds the optional HTML layer, and
`json_default` is the `json` encoder hook that lets date and time cells serialise.

Pass `functions={...}` to add your own functions to — or override — the built-in registry.
Each receives the extracted value first, then any arguments from the pipeline:

```python
from edjas import read_spec

data = read_spec("data.xlsx", "report.toml",
                 functions={"join": lambda v, sep: sep.join(v)})
# ... lets the specification use:  tags = '[Tags | join ", "]'
```

Rendering needs Jinja2, which arrives with the `demo` extra
(`pip install edjas[demo]`); the core install never imports it.

The signatures and descriptions below are read from the live objects by introspection.
