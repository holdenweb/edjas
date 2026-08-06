"""Worked examples for EDJAS: real UK-government workbooks, specs and templates.

Installed via the main package's ``demo`` extra (``pip install edjas[demo]``), this
package bundles seven unmodified, openly-licensed workbooks (see ``ATTRIBUTION.md``
alongside this module), a TOML extraction spec for each, and the Jinja2 templates that
render the extracted data as HTML reports. Build them all from the command line:

    edjas-examples                       # write every report to the current directory
    edjas-examples --example slgfs       # just one
    edjas-examples --example retail --shorten-headings

or from Python:

    from edjas_examples import build
    html = build("pesa")                 # writes report-pesa.html, returns the HTML

The shapes shown:

  - retail : a flat statistics release (records tables) with a named Excel Table
  - slgfs  : a Scottish local-government revenue-account ledger (category -> subtotal)
  - pesa   : an HM Treasury expenditure ledger (section -> category -> subtotal)
  - dwp    : a departmental spending ledger on a quoted spaced sheet ('Table 1')
  - wales  : a GSS-format revenue account with mixed £/%/per-head columns
  - cra    : a COFOG-classified ledger with "-" nil markers
  - nbs    : a headline slice carved from a large multi-sheet reference workbook
"""

import argparse
from importlib.resources import as_file, files
from pathlib import Path

from edjas import render_report

__all__ = ["DATA", "TEMPLATES", "EXAMPLES", "SHORTENED_HEADINGS", "build", "main"]

_PKG = files("edjas_examples")
DATA = _PKG / "data"
TEMPLATES = _PKG / "templates"

# Optional display labels for retail's two unwieldy revision-table column headings.
# Applied only with --shorten-headings; by default the report keeps the ONS wording.
# The keys must match the workbook's headings exactly (a test guards against staleness).
SHORTENED_HEADINGS = {
    "Revisions between first publication and estimates 12 months later "
    "(percentage points), Average over the last 5 years (mean revision)":
        "Mean revision (pp), 5-yr avg",
    "Revisions between first publication and estimates 12 months later "
    "(percentage points), Average over the last 5 years without regard to sign "
    "(average absolute revision)":
        "Average absolute revision (pp), 5-yr avg",
}

EXAMPLES = {
    "retail": {
        "xlsx": "retail.xlsx", "spec": "retail.toml",
        "template": "report.html", "out": "report.html",
        "headings": SHORTENED_HEADINGS,
    },
    "slgfs": {
        "xlsx": "slgfs.xlsx", "spec": "slgfs.toml",
        "template": "ledger.html", "out": "report-slgfs.html",
        "context": {
            "source_name": "Scottish Local Government Finance Statistics 2024-25, "
                           "Scottish Government",
            "source_url": "https://www.gov.scot/collections/"
                          "scottish-local-government-finance-statistics/",
        },
    },
    "pesa": {
        "xlsx": "pesa.xlsx", "spec": "pesa.toml",
        "template": "ledger.html", "out": "report-pesa.html",
        "context": {
            "source_name": "Public Expenditure Statistical Analyses 2025, HM Treasury",
            "source_url": "https://www.gov.uk/government/collections/"
                          "public-expenditure-statistical-analyses-pesa",
        },
    },
    "dwp": {
        "xlsx": "dwp.xlsx", "spec": "dwp.toml",
        "template": "ledger.html", "out": "report-dwp.html",
        "context": {
            "source_name": "Public spending and administration budget 2021 to 2026, "
                           "Department for Work and Pensions",
            "source_url": "https://www.gov.uk/government/publications/"
                          "dwp-annual-report-and-accounts-2024-to-2025",
        },
    },
    "wales": {
        "xlsx": "wales.xlsx", "spec": "wales.toml",
        "template": "ledger.html", "out": "report-wales.html",
        "context": {
            "source_name": "Local authority revenue and capital outturn expenditure: "
                           "April 2024 to March 2025, Welsh Government",
            "source_url": "https://www.gov.wales/local-authority-revenue-and-capital-"
                          "outturn-expenditure-april-2024-march-2025",
        },
    },
    "cra": {
        "xlsx": "cra.xlsx", "spec": "cra.toml",
        "template": "ledger.html", "out": "report-cra.html",
        "context": {
            "source_name": "Country and Regional Analysis 2024, HM Treasury",
            "source_url": "https://www.gov.uk/government/statistics/"
                          "country-and-regional-analysis-2024",
        },
    },
    "nbs": {
        "xlsx": "nbs.xlsx", "spec": "nbs.toml",
        "template": "ledger.html", "out": "report-nbs.html",
        "context": {
            "source_name": "The UK national balance sheet estimates, "
                           "Office for National Statistics",
            "source_url": "https://www.ons.gov.uk/economy/nationalaccounts/uksectoraccounts/"
                          "bulletins/nationalbalancesheet/previousReleases",
        },
    },
}


def build(name, shorten_headings=False, out=None):
    """Render one example, write it to ``out``, and return the HTML.

    ``out`` defaults to the example's usual filename in the **current directory** —
    the package's own data is read-only once installed. ``shorten_headings`` is off by
    default, so reports reproduce each workbook's column titles verbatim.
    """
    cfg = EXAMPLES[name]
    headings = cfg.get("headings") if shorten_headings else None
    # as_file materialises each resource as a real filesystem path, so this also works
    # if the package is ever imported from a zip archive; for a normal directory
    # install it is a no-op passthrough.
    with (
        as_file(DATA / cfg["xlsx"]) as xlsx,
        as_file(DATA / cfg["spec"]) as spec,
        as_file(TEMPLATES) as templates,
    ):
        html = render_report(
            str(xlsx), str(spec),
            template=cfg["template"], templates_dir=str(templates),
            headings=headings, **cfg.get("context", {}),
        )
    out = Path(out) if out is not None else Path.cwd() / cfg["out"]
    out.write_text(html, encoding="utf-8")
    print(f"wrote {out} ({len(html):,} bytes)")
    return html


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="edjas-examples",
        description="Build the EDJAS example HTML reports from their bundled workbooks.",
    )
    parser.add_argument(
        "--example", choices=[*EXAMPLES, "all"], default="all",
        help="which example to build (default: all)",
    )
    parser.add_argument(
        "--shorten-headings", action="store_true",
        help="shorten retail's long ONS column headings (default: keep them verbatim)",
    )
    parser.add_argument(
        "--out-dir", default=".",
        help="directory to write the reports into (default: current directory)",
    )
    args = parser.parse_args(argv)
    names = list(EXAMPLES) if args.example == "all" else [args.example]
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    for name in names:
        build(name, shorten_headings=args.shorten_headings,
              out=out_dir / EXAMPLES[name]["out"])
