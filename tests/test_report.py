"""Tests for the core HTML-report renderer (``edjas.report``) that need no example data.

``render_report`` is gated behind the ``demo`` extra (Jinja2): the error-path test
simulates Jinja2 being absent, and the ``_rename_headings`` unit tests exercise pure
core code. Tests that render the bundled example workbooks live in
``test_report_examples.py``, which skips when the ``edjas-examples`` companion package
is not installed.
"""

import builtins

import pytest

from edjas.report import render_report, _rename_headings


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
