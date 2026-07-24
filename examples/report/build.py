"""Build an HTML report from the bundled ONS Retail Sales Index workbook.

Needs the optional 'demo' extra (``pip install edjas[demo]``). Run from anywhere:

    python examples/report/build.py                     # faithful to the workbook
    python examples/report/build.py --shorten-headings  # tidy the long column titles

It reads ``retail.xlsx`` through ``retail.toml`` — leaving the workbook untouched —
and writes ``report.html`` beside them.
"""

import argparse
from pathlib import Path

from edjas import render_report

HERE = Path(__file__).parent

# Optional display labels for the two unwieldy revision-table column headings. This is
# applied only when --shorten-headings is passed; by default the report reproduces the
# ONS wording verbatim. The keys must match the workbook's headings exactly (a test
# guards against them going stale).
SHORTENED_HEADINGS = {
    "Revisions between first publication and estimates 12 months later "
    "(percentage points), Average over the last 5 years (mean revision)":
        "Mean revision (pp), 5-yr avg",
    "Revisions between first publication and estimates 12 months later "
    "(percentage points), Average over the last 5 years without regard to sign "
    "(average absolute revision)":
        "Average absolute revision (pp), 5-yr avg",
}


def main(shorten_headings=False, out=None):
    """Render the report and write it to ``out`` (default ``report.html`` here).

    Returns the rendered HTML. ``shorten_headings`` is off by default, so the report
    reproduces the workbook's column titles verbatim unless it is turned on.
    """
    html = render_report(
        spreadsheet=HERE / "retail.xlsx",
        spec=HERE / "retail.toml",
        template="report.html",
        templates_dir=HERE / "templates",
        headings=SHORTENED_HEADINGS if shorten_headings else None,
    )
    out = Path(out) if out is not None else HERE / "report.html"
    out.write_text(html, encoding="utf-8")
    print(f"wrote {out} ({len(html):,} bytes)")
    return html


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--shorten-headings",
        action="store_true",
        help="shorten the long ONS column headings for readability "
             "(default: keep them verbatim)",
    )
    args = parser.parse_args()
    main(shorten_headings=args.shorten_headings)
