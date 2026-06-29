import json
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from edjas import json_default, read_file

FIXTURE = Path(__file__).parent / "data" / "parameters.xlsx"


def make_workbook(tmp_path, cells, defined_names):
    """Build a single-sheet workbook from {coord: value} and {name: ref}."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for coord, value in cells.items():
        ws[coord] = value
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


# --- function markup -------------------------------------------------------

def test_markup_records(tmp_path):
    path = make_workbook(
        tmp_path,
        {
            "A1": "sales", "B1": "[records Sales]",
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
    """{int Prices} coerces text-formatted values after building the object."""
    path = make_workbook(
        tmp_path,
        {
            "A1": "prices", "B1": "{int Prices}",
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
            "A1": "grid", "B1": "[transpose Grid]",
            "D1": 1, "E1": 2, "F1": 3,
            "D2": 4, "E2": 5, "F2": 6,
        },
        {"Parameters": "$A$1:$B$1", "Grid": "$D$1:$F$2"},
    )
    assert read_file(path) == {"grid": [[1, 4], [2, 5], [3, 6]]}


def test_markup_no_function_unchanged(tmp_path):
    """A bare [name] with no function token still behaves as before."""
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
            "A1": "vec", "B1": "[bogus Vec]",
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
            "A1": "vec", "B1": "[double Vec]",
            "D1": 1, "E1": 2, "F1": 3,
        },
        {"Parameters": "$A$1:$B$1", "Vec": "$D$1:$F$1"},
    )
    out = read_file(path, functions={"double": lambda v: [x * 2 for x in v]})
    assert out == {"vec": [2, 4, 6]}


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
