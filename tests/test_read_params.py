import json
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from edjas import json_default, read_file
from edjas.read_params import parse_pipeline

FIXTURE = Path(__file__).parent / "data" / "parameters.xlsx"


def make_workbook(tmp_path, cells, defined_names, extra_sheets=None):
    """Build a workbook from {coord: value} on Sheet1 and {name: ref} names.

    ``extra_sheets`` optionally adds further sheets as {sheet_name: {coord: value}},
    enabling cross-sheet references such as ``Data!D1:F2``. Defined names are
    anchored to Sheet1.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for coord, value in cells.items():
        ws[coord] = value
    for sheet_name, sheet_cells in (extra_sheets or {}).items():
        sheet = wb.create_sheet(sheet_name)
        for coord, value in sheet_cells.items():
            sheet[coord] = value
    for name, ref in defined_names.items():
        wb.defined_names.add(DefinedName(name, attr_text=f"Sheet1!{ref}"))
    path = tmp_path / "wb.xlsx"
    wb.save(path)
    return path


@pytest.fixture(scope="module")
def params():
    return read_file(FIXTURE)


def test_scalar(params):
    assert params["title"] == "Hubris Demo"


def test_named_range_dict_hours(params):
    assert params["hours"]["Monday"] == "7:00 am - 8:00 pm"
    assert params["hours"]["Sunday"] == "Closed"
    assert len(params["hours"]) == 7


def test_named_range_dict_prices(params):
    assert params["prices"] == {"Tea": 3.25, "Coffee": 4.0, "Bacon Sandwich": 8.25}


def test_matrix_rows(params):
    assert params["hours_list"][0] == ["Monday", "7:00 am - 8:00 pm"]
    assert params["hours_list"][5] == ["Saturday", "9:00 am - 5:00 pm"]
    assert len(params["hours_list"]) == 7


def test_horizontal_vector(params):
    assert params["h_vector"] == ["the", "quick", "brown", "fox"]


def test_vertical_vector(params):
    assert params["v_vector"] == ["jumps", "over", "the", "lazy", "dog"]


def test_expected_keys_present(params):
    assert set(params) >= {"title", "hours", "prices", "hours_list", "h_vector", "v_vector"}


def test_single_entry_dict(tmp_path):
    """A one-row {dict} range must read as a dict, not be flattened to a vector."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "version", "B1": "{version}",
            "A2": "number", "B2": "0.1.2",
        },
        {"Parameters": "$A$1:$B$1", "version": "$A$2:$B$2"},
    )
    assert read_file(path) == {"version": {"number": "0.1.2"}}


# --- pipeline parser -------------------------------------------------------

def test_parse_pipeline_bare_source():
    assert parse_pipeline("Sales") == ("Sales", [])


def test_parse_pipeline_chain_no_args():
    assert parse_pipeline("Grid | transpose | records") == (
        "Grid",
        [("transpose", []), ("records", [])],
    )


def test_parse_pipeline_quoted_arg_keeps_spaces_and_pipe():
    assert parse_pipeline('Words | join " | "') == (
        "Words",
        [("join", [(" | ", True)])],
    )


def test_parse_pipeline_mixed_args():
    assert parse_pipeline("Nums | clamp 0 10 Bounds") == (
        "Nums",
        [("clamp", [("0", False), ("10", False), ("Bounds", False)])],
    )


def test_parse_pipeline_rejects_multi_token_source():
    with pytest.raises(ValueError, match="single range reference"):
        parse_pipeline("a b | records")


def test_parse_pipeline_rejects_empty_stage():
    with pytest.raises(ValueError, match="Empty function stage"):
        parse_pipeline("Grid | transpose |")


def test_parse_pipeline_rejects_unterminated_string():
    with pytest.raises(ValueError, match="Unterminated string"):
        parse_pipeline('Words | join "oops')


# --- pipe-notation function markup -----------------------------------------

def test_markup_records(tmp_path):
    path = make_workbook(
        tmp_path,
        {
            "A1": "sales", "B1": "[Sales | records]",
            "D1": "Region", "E1": "Q1",
            "D2": "North", "E2": 100,
            "D3": "South", "E3": 200,
        },
        {"Parameters": "$A$1:$B$1", "Sales": "$D$1:$E$3"},
    )
    assert read_file(path) == {
        "sales": [
            {"Region": "North", "Q1": 100},
            {"Region": "South", "Q1": 200},
        ]
    }


def test_markup_function_on_object(tmp_path):
    """{Prices | int} coerces text-formatted values after building the object."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "prices", "B1": "{Prices | int}",
            "D1": "Tea", "E1": "3",
            "D2": "Coffee", "E2": "4",
        },
        {"Parameters": "$A$1:$B$1", "Prices": "$D$1:$E$2"},
    )
    assert read_file(path) == {"prices": {"Tea": 3, "Coffee": 4}}


def test_markup_transpose(tmp_path):
    path = make_workbook(
        tmp_path,
        {
            "A1": "grid", "B1": "[Grid | transpose]",
            "D1": 1, "E1": 2, "F1": 3,
            "D2": 4, "E2": 5, "F2": 6,
        },
        {"Parameters": "$A$1:$B$1", "Grid": "$D$1:$F$2"},
    )
    assert read_file(path) == {"grid": [[1, 4], [2, 5], [3, 6]]}


def test_markup_chained_pipeline(tmp_path):
    """Stages apply left to right. Grid is field-per-row (column-oriented);
    transpose turns fields into the header row, then records builds objects."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "people", "B1": "[Grid | transpose | records]",
            "D1": "name", "E1": "alice", "F1": "bob",
            "D2": "age", "E2": 30, "F2": 25,
        },
        {"Parameters": "$A$1:$B$1", "Grid": "$D$1:$F$2"},
    )
    assert read_file(path) == {
        "people": [
            {"name": "alice", "age": 30},
            {"name": "bob", "age": 25},
        ]
    }


def test_markup_no_function_unchanged(tmp_path):
    """A bare [name] with no pipeline still behaves as before."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "vec", "B1": "[Vec]",
            "D1": "a", "E1": "b", "F1": "c",
        },
        {"Parameters": "$A$1:$B$1", "Vec": "$D$1:$F$1"},
    )
    assert read_file(path) == {"vec": ["a", "b", "c"]}


def test_markup_unknown_function_raises(tmp_path):
    path = make_workbook(
        tmp_path,
        {
            "A1": "vec", "B1": "[Vec | bogus]",
            "D1": 1, "E1": 2,
        },
        {"Parameters": "$A$1:$B$1", "Vec": "$D$1:$E$1"},
    )
    with pytest.raises(ValueError, match="Unknown EDJAS function 'bogus'"):
        read_file(path)


def test_injected_custom_function(tmp_path):
    path = make_workbook(
        tmp_path,
        {
            "A1": "vec", "B1": "[Vec | double]",
            "D1": 1, "E1": 2, "F1": 3,
        },
        {"Parameters": "$A$1:$B$1", "Vec": "$D$1:$F$1"},
    )
    out = read_file(path, functions={"double": lambda v: [x * 2 for x in v]})
    assert out == {"vec": [2, 4, 6]}


def test_injected_function_with_numeric_arg(tmp_path):
    """A bare number argument is passed as an int/float literal."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "vec", "B1": "[Vec | scale 10]",
            "D1": 1, "E1": 2, "F1": 3,
        },
        {"Parameters": "$A$1:$B$1", "Vec": "$D$1:$F$1"},
    )
    out = read_file(path, functions={"scale": lambda v, n: [x * n for x in v]})
    assert out == {"vec": [10, 20, 30]}


def test_injected_function_with_range_reference_arg(tmp_path):
    """A bare-word argument is resolved as a named range and passed in."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "totals", "B1": "[Vec | add Other]",
            "D1": 1, "E1": 2, "F1": 3,
            "D2": 10, "E2": 20, "F2": 30,
        },
        {
            "Parameters": "$A$1:$B$1",
            "Vec": "$D$1:$F$1",
            "Other": "$D$2:$F$2",
        },
    )
    out = read_file(
        path,
        functions={"add": lambda v, other: [a + b for a, b in zip(v, other)]},
    )
    assert out == {"totals": [11, 22, 33]}


def test_injected_function_with_quoted_arg(tmp_path):
    """A double-quoted argument is passed as a string literal (spaces kept)."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "csv", "B1": '[Words | join ", "]',
            "D1": "the", "E1": "quick", "F1": "brown",
        },
        {"Parameters": "$A$1:$B$1", "Words": "$D$1:$F$1"},
    )
    out = read_file(path, functions={"join": lambda v, sep: sep.join(v)})
    assert out == {"csv": "the, quick, brown"}


# --- raw A1-style range specifications (not named ranges) ------------------

def test_raw_range_bare_source(tmp_path):
    """A bare [A1:B2] cell range works without a defined name."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "vec", "B1": "[D1:F1]",
            "D1": "a", "E1": "b", "F1": "c",
        },
        {"Parameters": "$A$1:$B$1"},
    )
    assert read_file(path) == {"vec": ["a", "b", "c"]}


def test_raw_range_list_source_with_function(tmp_path):
    """A raw range may be the source of a pipeline: [D1:F2 | transpose]."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "grid", "B1": "[D1:F2 | transpose]",
            "D1": 1, "E1": 2, "F1": 3,
            "D2": 4, "E2": 5, "F2": 6,
        },
        {"Parameters": "$A$1:$B$1"},
    )
    assert read_file(path) == {"grid": [[1, 4], [2, 5], [3, 6]]}


def test_raw_range_object_source_with_function(tmp_path):
    """A raw two-column range works as a {object} source: {D1:E2 | int}."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "prices", "B1": "{D1:E2 | int}",
            "D1": "Tea", "E1": "3",
            "D2": "Coffee", "E2": "4",
        },
        {"Parameters": "$A$1:$B$1"},
    )
    assert read_file(path) == {"prices": {"Tea": 3, "Coffee": 4}}


def test_raw_range_as_reference_argument(tmp_path):
    """A raw range may be a reference argument: [D1:F1 | add D2:F2]."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "totals", "B1": "[D1:F1 | add D2:F2]",
            "D1": 1, "E1": 2, "F1": 3,
            "D2": 10, "E2": 20, "F2": 30,
        },
        {"Parameters": "$A$1:$B$1"},
    )
    out = read_file(
        path,
        functions={"add": lambda v, other: [a + b for a, b in zip(v, other)]},
    )
    assert out == {"totals": [11, 22, 33]}


# --- cross-sheet raw range specifications (Sheet!A1:B4) ---------------------

def test_cross_sheet_raw_range_list_source(tmp_path):
    """A raw range qualified by another sheet name works as a source."""
    path = make_workbook(
        tmp_path,
        {"A1": "grid", "B1": "[Data!D1:F2 | transpose]"},
        {"Parameters": "$A$1:$B$1"},
        extra_sheets={"Data": {
            "D1": 1, "E1": 2, "F1": 3,
            "D2": 4, "E2": 5, "F2": 6,
        }},
    )
    assert read_file(path) == {"grid": [[1, 4], [2, 5], [3, 6]]}


def test_cross_sheet_raw_range_object_source(tmp_path):
    """A cross-sheet two-column raw range works as a {object} source."""
    path = make_workbook(
        tmp_path,
        {"A1": "prices", "B1": "{Data!D1:E2 | int}"},
        {"Parameters": "$A$1:$B$1"},
        extra_sheets={"Data": {
            "D1": "Tea", "E1": "3",
            "D2": "Coffee", "E2": "4",
        }},
    )
    assert read_file(path) == {"prices": {"Tea": 3, "Coffee": 4}}


def test_cross_sheet_raw_range_as_reference_argument(tmp_path):
    """A cross-sheet raw range works as a reference argument."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "totals", "B1": "[D1:F1 | add Data!D1:F1]",
            "D1": 1, "E1": 2, "F1": 3,
        },
        {"Parameters": "$A$1:$B$1"},
        extra_sheets={"Data": {"D1": 10, "E1": 20, "F1": 30}},
    )
    out = read_file(
        path,
        functions={"add": lambda v, other: [a + b for a, b in zip(v, other)]},
    )
    assert out == {"totals": [11, 22, 33]}


def test_date_cell_serialises(tmp_path):
    """A date cell reads as a datetime and serialises via json_default."""
    path = make_workbook(
        tmp_path,
        {"A1": "opened", "B1": datetime(2026, 6, 29, 9, 30)},
        {"Parameters": "$A$1:$B$1"},
    )
    data = read_file(path)
    assert data["opened"] == datetime(2026, 6, 29, 9, 30)
    assert json.loads(json.dumps(data, default=json_default)) == {
        "opened": "2026-06-29T09:30:00"
    }
