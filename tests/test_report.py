"""Tests for the optional HTML reporting demo (``edjas.report.render_report``).

The renderer is gated behind the ``demo`` extra (Jinja2). The error-path test
simulates Jinja2 being absent; the rendering test skips when it is genuinely absent.
"""

import builtins
import importlib.util
from pathlib import Path

import pytest

from edjas.report import render_report, _rename_headings

EXAMPLE = Path(__file__).parent.parent / "examples" / "report"

# The exact (verbatim) long column heading ONS ships in the revisions table.
LONG_HEADING = (
    "Revisions between first publication and estimates 12 months later "
    "(percentage points), Average over the last 5 years (mean revision)"
)


def _load_build_module():
    """Import examples/report/build.py as a module (its __main__ guard stays inert)."""
    spec = importlib.util.spec_from_file_location("edjas_demo_build", EXAMPLE / "build.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_render_report_without_jinja2_gives_actionable_error(monkeypatch):
    """With Jinja2 missing, the error names the extra to install rather than leaking."""
    real_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name == "jinja2":
            raise ModuleNotFoundError("No module named 'jinja2'")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    with pytest.raises(ModuleNotFoundError, match=r"edjas\[demo\]"):
        render_report("x.xlsx", "y.toml", "report.html", ".")


def test_render_report_from_bundled_example():
    """The bundled ONS workbook renders through the multi-level templates."""
    pytest.importorskip("jinja2")
    html = render_report(
        EXAMPLE / "retail.xlsx",
        EXAMPLE / "retail.toml",
        "report.html",
        EXAMPLE / "templates",
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


def test_rename_headings_recursive_and_order_preserving():
    """Matching dict keys are renamed at any depth; others and order are untouched."""
    data = {
        "rows": [{"Long A": 1, "Keep": 2}, {"Long A": 3, "Keep": 4}],
        "nested": {"Long A": 5},
        "scalar": "x",
    }
    out = _rename_headings(data, {"Long A": "A"})
    assert out["rows"] == [{"A": 1, "Keep": 2}, {"A": 3, "Keep": 4}]
    assert out["nested"] == {"A": 5}
    assert out["scalar"] == "x"
    assert list(out["rows"][0]) == ["A", "Keep"]  # order preserved


def test_rename_headings_raises_on_collision():
    """A rename onto an existing/other-mapped key fails loudly rather than losing data."""
    with pytest.raises(ValueError, match="collision"):
        _rename_headings({"A": 1, "B": 2}, {"A": "B"})  # 'A' -> 'B' would drop 'B'
    with pytest.raises(ValueError, match="collision"):
        _rename_headings({"L1": 1, "L2": 2}, {"L1": "X", "L2": "X"})  # both -> 'X'


def test_render_report_headings_default_keeps_original():
    """With no mapping (the default), headings are rendered exactly as extracted."""
    pytest.importorskip("jinja2")
    html = render_report(
        EXAMPLE / "retail.xlsx", EXAMPLE / "retail.toml", "report.html", EXAMPLE / "templates"
    )
    assert LONG_HEADING in html


def test_render_report_headings_rename_applies():
    """A supplied mapping renames the heading in the rendered report."""
    pytest.importorskip("jinja2")
    html = render_report(
        EXAMPLE / "retail.xlsx", EXAMPLE / "retail.toml", "report.html", EXAMPLE / "templates",
        headings={LONG_HEADING: "Mean revision (pp), 5-yr avg"},
    )
    assert "Mean revision (pp), 5-yr avg" in html
    assert LONG_HEADING not in html


def test_demo_heading_map_keys_exist_in_data():
    """The demo's shortening map must not go stale: every key exists in the data.

    This is a pure data-integrity check (read_spec + a module load), so it must run
    even in a core-only environment without the optional Jinja2 extra.
    """
    from edjas import read_spec

    build = _load_build_module()
    data = read_spec(EXAMPLE / "retail.xlsx", EXAMPLE / "retail.toml")
    present = {key for row in data["revisions"] for key in row}
    assert build.SHORTENED_HEADINGS, "demo should define a non-empty shortening map"
    for long_key in build.SHORTENED_HEADINGS:
        assert long_key in present, f"stale heading-map key: {long_key!r}"


def test_demo_build_main_default_is_faithful_and_flag_shortens(tmp_path):
    """build.main() defaults to verbatim headings; shorten_headings=True tidies them."""
    pytest.importorskip("jinja2")
    build = _load_build_module()

    default_html = build.main(out=tmp_path / "default.html")  # default: off
    assert LONG_HEADING in default_html

    short_html = build.main(shorten_headings=True, out=tmp_path / "short.html")
    assert "Mean revision (pp), 5-yr avg" in short_html
    assert LONG_HEADING not in short_html


def test_render_report_leaves_workbook_untouched():
    """Rendering only reads the workbook; the bundled file's bytes are unchanged."""
    pytest.importorskip("jinja2")
    before = (EXAMPLE / "retail.xlsx").read_bytes()
    render_report(
        EXAMPLE / "retail.xlsx",
        EXAMPLE / "retail.toml",
        "report.html",
        EXAMPLE / "templates",
    )
    assert (EXAMPLE / "retail.xlsx").read_bytes() == before
