"""Load a TOML extraction spec and apply it to an (unmodified) spreadsheet.

An EDJAS spec describes what to pull out of a workbook without touching the workbook
itself. It is a TOML file with an ``[extract]`` table mapping output keys to
extraction expressions::

    [extract]
    title  = "Summary!B2"                    # bare ref  -> scalar
    prices = "{Prices | int}"                # { }       -> object, then coerce
    sales  = "[Sales | records]"             # [ ]       -> list/table, then reshape

Each expression is evaluated by :func:`edjas.read_params.evaluate`, so the full
pipe-and-function language is available.
"""

import tomllib

import openpyxl

from . import functions as _functions
from . import ods as _ods
from .read_params import evaluate

__all__ = ["load_spec", "read_spec", "open_workbook"]


def open_workbook(spreadsheet):
    """Open ``spreadsheet`` for reading, whatever spreadsheet format it is in.

    An OpenDocument file is parsed by :mod:`edjas.ods` and presented as an openpyxl
    workbook, so the rest of EDJAS never needs to know which format it came from.
    Anything else is handed to openpyxl, with ``data_only=True`` so that a formula cell
    yields the value the application cached rather than the formula text. A workbook
    that has never been recalculated has no cache, so such a cell reads as None.
    """
    if _ods.is_ods(spreadsheet):
        return _ods.load_workbook(spreadsheet)
    return openpyxl.load_workbook(spreadsheet, data_only=True)


def load_spec(path):
    """Parse a TOML spec file and return its ``[extract]`` mapping."""
    with open(path, "rb") as f:
        data = tomllib.load(f)
    extract = data.get("extract")
    if not isinstance(extract, dict):
        raise ValueError(f"{path}: spec must contain an [extract] table")
    return extract


def read_spec(spreadsheet, spec, functions=None):
    """Extract data from ``spreadsheet`` as directed by the TOML ``spec`` file.

    ``spreadsheet`` may be an Excel workbook or an OpenDocument (``.ods``) one; the
    same expressions work against either. ``functions`` optionally adds to (or
    overrides) the built-in registry. The spreadsheet is opened read-only and never
    modified.
    """
    mapping = load_spec(spec)
    workbook = open_workbook(spreadsheet)
    registry = _functions.resolve(functions)
    result = {}
    for key, expr in mapping.items():
        try:
            result[key] = evaluate(workbook, expr, registry)
        except Exception as exc:  # add the spec key/expression to any failure
            raise ValueError(f"extracting {key!r} from {expr!r}: {exc}") from exc
    return result
