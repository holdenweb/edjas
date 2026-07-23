"""Tests for the optional HTML reporting demo (``edjas.report.render_report``).

The renderer is gated behind the ``demo`` extra (Jinja2). The error-path test
simulates Jinja2 being absent; the rendering test skips when it is genuinely absent.
"""

import builtins
from pathlib import Path

import pytest

from edjas.report import render_report

EXAMPLE = Path(__file__).parent.parent / "examples" / "report"


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
