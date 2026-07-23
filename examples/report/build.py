"""Build an HTML report from the bundled ONS Retail Sales Index workbook.

Needs the optional 'demo' extra (``pip install edjas[demo]``). Run from anywhere:

    python examples/report/build.py

It reads ``retail.xlsx`` through ``retail.toml`` — leaving the workbook untouched —
and writes ``report.html`` beside them.
"""

from pathlib import Path

from edjas import render_report

HERE = Path(__file__).parent


def main():
    html = render_report(
        spreadsheet=HERE / "retail.xlsx",
        spec=HERE / "retail.toml",
        template="report.html",
        templates_dir=HERE / "templates",
    )
    out = HERE / "report.html"
    out.write_text(html, encoding="utf-8")
    print(f"wrote {out} ({len(html):,} bytes)")


if __name__ == "__main__":
    main()
