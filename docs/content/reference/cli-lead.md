EDJAS installs one command; the companion examples package installs a second.

`edjas` takes a spreadsheet and a specification and writes JSON to standard output, so it
composes with everything else on the command line:

```
edjas data.xlsx report.toml | jq '.sales'
edjas data.xlsx report.toml > extracted.json
```

`edjas-examples` builds the worked examples from their bundled workbooks, writing the
reports to the current directory unless told otherwise. It arrives with the `demo` extra.

The help below is captured from the commands themselves.
