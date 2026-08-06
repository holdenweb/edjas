"""Tests for the bundled demonstrations in the ``edjas-examples`` companion package.

The whole module skips when ``edjas_examples`` is not installed (it arrives via the
``demo`` extra, and via the dev dependency group in this workspace), so a core-only
environment still collects and runs the rest of the suite. Importing ``edjas_examples``
itself needs no Jinja2, so the data-integrity check runs even without the renderer;
tests that actually render skip individually when Jinja2 is absent.
"""

from pathlib import Path

import pytest

edjas_examples = pytest.importorskip("edjas_examples")

from edjas.report import render_report  # noqa: E402

DATA = edjas_examples.DATA
TEMPLATES = edjas_examples.TEMPLATES
RENDERED = Path(__file__).parent.parent / "edjas-examples" / "rendered"

# The exact (verbatim) long column heading ONS ships in the revisions table.
LONG_HEADING = (
    "Revisions between first publication and estimates 12 months later "
    "(percentage points), Average over the last 5 years (mean revision)"
)


def test_render_report_from_bundled_example():
    """The bundled ONS workbook renders through the multi-level templates."""
    pytest.importorskip("jinja2")
    html = render_report(
        str(DATA / "retail.xlsx"),
        str(DATA / "retail.toml"),
        "report.html",
        str(TEMPLATES),
    )
    # header metadata, read via the single-quoted spaced sheet 'Cover Sheet'
    assert "Retail Sales Index" in html
    assert "published at 7.00am 20 June 2025" in html
    # contents records, read from the 'Contents' sheet
    assert "Standard errors" in html and "Response rates" in html
    # the named Excel Table, tidied by the pipeline's `round 3`
    assert "-0.137" in html and "0.427" in html
    # attribution present, and the template was actually rendered
    assert "Open Government Licence" in html
    assert "{{" not in html and "{%" not in html


def test_render_report_headings_default_keeps_original():
    """With no mapping (the default), headings are rendered exactly as extracted."""
    pytest.importorskip("jinja2")
    html = render_report(
        str(DATA / "retail.xlsx"), str(DATA / "retail.toml"), "report.html", str(TEMPLATES)
    )
    assert LONG_HEADING in html


def test_render_report_headings_rename_applies():
    """A supplied mapping renames the heading in the rendered report."""
    pytest.importorskip("jinja2")
    html = render_report(
        str(DATA / "retail.xlsx"), str(DATA / "retail.toml"), "report.html", str(TEMPLATES),
        headings={LONG_HEADING: "Mean revision (pp), 5-yr avg"},
    )
    assert "Mean revision (pp), 5-yr avg" in html
    assert LONG_HEADING not in html


def test_demo_heading_map_keys_exist_in_data():
    """The demo's shortening map must not go stale: every key exists in the data.

    This is a pure data-integrity check (read_spec over the package data), so it must
    run even without the optional Jinja2 renderer — importing edjas_examples needs no
    Jinja2.
    """
    from edjas import read_spec

    data = read_spec(str(DATA / "retail.xlsx"), str(DATA / "retail.toml"))
    present = {key for row in data["revisions"] for key in row}
    assert edjas_examples.SHORTENED_HEADINGS, "demo should define a non-empty shortening map"
    for long_key in edjas_examples.SHORTENED_HEADINGS:
        assert long_key in present, f"stale heading-map key: {long_key!r}"


def test_demo_build_retail_default_and_flag(tmp_path):
    """build('retail') defaults to verbatim headings; shorten_headings=True tidies them."""
    pytest.importorskip("jinja2")
    default_html = edjas_examples.build("retail", out=tmp_path / "default.html")  # default: off
    assert LONG_HEADING in default_html

    short_html = edjas_examples.build("retail", shorten_headings=True, out=tmp_path / "short.html")
    assert "Mean revision (pp), 5-yr avg" in short_html
    assert LONG_HEADING not in short_html


def test_demo_build_slgfs_ledger(tmp_path):
    """The SLGFS ledger renders its category/subtotal structure and formatted figures."""
    pytest.importorskip("jinja2")
    html = edjas_examples.build("slgfs", out=tmp_path / "slgfs.html")
    assert "Net Revenue Expenditure" in html            # title text
    assert "Total Education" in html                     # a subtotal row
    assert "All Services (GF + HRA)" in html             # the grand total
    assert "8,255,219" in html                           # a figure with thousands sep
    assert 'class="subtotal"' in html                    # row classification applied
    assert "Open Government Licence" in html             # attribution footer
    assert "{{" not in html and "{%" not in html


def test_demo_build_pesa_ledger(tmp_path):
    """The PESA ledger renders section headings, subtotals and formatted figures."""
    pytest.importorskip("jinja2")
    html = edjas_examples.build("pesa", out=tmp_path / "pesa.html")
    assert "Total Managed Expenditure" in html           # title text
    assert "CURRENT EXPENDITURE" in html                 # a section heading
    assert 'class="section"' in html                     # section row classification
    assert "Total resource DEL" in html                  # a subtotal
    assert "484,062" in html                             # a figure with thousands sep
    assert "Open Government Licence" in html
    assert "{{" not in html and "{%" not in html


def test_demo_build_dwp_ledger(tmp_path):
    """DWP renders via a quoted spaced sheet name; category headings shed "0" padding."""
    pytest.importorskip("jinja2")
    html = edjas_examples.build("dwp", out=tmp_path / "dwp.html")
    assert "Department for Work and Pensions" in html
    assert "Section A: Core Department" in html                   # a line item
    assert "Total departmental spending" in html                  # a grand total
    # 'Resource DEL2' pads its row with string "0"s: it must be a heading, not a data row
    assert '<th scope="colgroup" class="rowlabel" colspan="7">Resource DEL2</th>' in html
    assert "{{" not in html and "{%" not in html


def test_demo_build_wales_ledger(tmp_path):
    """Wales renders its contents table and formats mixed £/%/per-head columns."""
    pytest.importorskip("jinja2")
    html = edjas_examples.build("wales", out=tmp_path / "wales.html")
    assert "Local authority revenue and capital outturn" in html
    assert "Worksheets in this release" in html                   # contents section
    assert "All county and county borough council expenditure" in html   # a subtotal
    assert "Council tax requirement" in html
    # fractional values keep one decimal place rather than being rounded to whole
    assert "3,596,313.8" in html and "8.6" in html
    assert "{{" not in html and "{%" not in html


def test_demo_build_cra_ledger(tmp_path):
    """CRA keeps nil "-" line items as rows, not headings, and totals its functions."""
    pytest.importorskip("jinja2")
    html = edjas_examples.build("cra", out=tmp_path / "cra.html")
    assert "Total identifiable expenditure on services in England" in html
    assert "Total Expenditure on Services in England" in html     # the grand total
    # a real COFOG function heading is a section row
    assert '<th scope="colgroup" class="rowlabel" colspan="6">1. General public services</th>' in html
    # an all-nil sub-function is a LINE item (with its dashes), not a section heading
    assert "1.2 Foreign economic aid" in html
    assert 'colspan="6">1.2 Foreign economic aid' not in html
    assert "{{" not in html and "{%" not in html


def test_demo_build_nbs_slice(tmp_path):
    """NBS carves a readable headline slice from a large multi-sheet workbook."""
    pytest.importorskip("jinja2")
    html = edjas_examples.build("nbs", out=tmp_path / "nbs.html")
    assert "The UK national balance sheet" in html
    assert "net worth at start of 2024" in html                   # Table A subtitle
    assert "Reference Table" in html                              # contents rendered
    assert "Non-financial corporations" in html                   # a sector line item
    assert 'class="subtotal"' in html and "Total economy" in html # the roll-up row
    assert "2,451,431" in html                                    # formatted figure
    assert "Dwellings2" in html      # faithful: glued footnote marker is preserved
    assert "{{" not in html and "{%" not in html


def test_every_demo_example_builds(tmp_path):
    """Every registered example renders without error and produces real HTML."""
    pytest.importorskip("jinja2")
    for name in edjas_examples.EXAMPLES:
        html = edjas_examples.build(name, out=tmp_path / f"{name}.html")
        assert html.startswith("<!doctype html>")
        assert "{{" not in html and "{%" not in html


@pytest.mark.skipif(not RENDERED.is_dir(), reason="showcase output not in this checkout")
def test_committed_reports_match_freshly_rendered(tmp_path):
    """The showcase HTML in edjas-examples/rendered/ is exactly what the code produces.

    Rendering is deterministic -- nothing in the templates emits a timestamp -- so any
    template change that alters a report shows up here rather than as a stale committed
    artefact. The documentation site extends the same ``base.html``, so this is also the
    guard that site-only template changes never leak into users' generated reports.

    If a template change is deliberate, regenerate the showcase before committing:
        edjas-examples --out-dir edjas-examples/rendered
    """
    pytest.importorskip("jinja2")
    for name, cfg in edjas_examples.EXAMPLES.items():
        fresh = edjas_examples.build(name, out=tmp_path / cfg["out"])
        assert fresh == (RENDERED / cfg["out"]).read_text(encoding="utf-8"), (
            f"{name}: committed showcase is stale; "
            f"run `edjas-examples --out-dir edjas-examples/rendered`"
        )


def test_render_report_leaves_workbook_untouched():
    """Rendering only reads the workbook; the bundled file's bytes are unchanged."""
    pytest.importorskip("jinja2")
    before = (DATA / "retail.xlsx").read_bytes()
    render_report(
        str(DATA / "retail.xlsx"),
        str(DATA / "retail.toml"),
        "report.html",
        str(TEMPLATES),
    )
    assert (DATA / "retail.xlsx").read_bytes() == before
