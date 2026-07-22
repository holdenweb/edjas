"""Guard the worked examples in README.md.

The README documents the JSON produced by ``examples/example.toml`` against
``examples/example.xlsx``. These tests assert those exact values, so an accidental
change to the engine, the sample workbook or the spec shows up as a failure rather
than as quietly wrong documentation.
"""

import json
from pathlib import Path

import pytest

from edjas import json_default, read_spec

EXAMPLES = Path(__file__).parent.parent / "examples"
XLSX = EXAMPLES / "example.xlsx"
SPEC = EXAMPLES / "example.toml"


@pytest.fixture(scope="module")
def out():
    return read_spec(XLSX, SPEC)


# --- the three extraction forms --------------------------------------------

def test_scalar_form(out):
    assert out["title"] == "Riverside Cafe"


def test_list_form(out):
    assert out["tags"] == ["Vegan", "Gluten-free", "Dairy-free"]


def test_object_form(out):
    assert out["hours"]["Monday"] == "07:00-20:00"
    assert out["hours"]["Sunday"] == "Closed"
    assert len(out["hours"]) == 7


# --- reference styles -------------------------------------------------------

def test_a1_reference_on_active_sheet(out):
    assert out["title_again"] == "Riverside Cafe"


def test_sheet_qualified_reference(out):
    assert out["prices"] == {"Tea": 3.25, "Coffee": 4.0, "Bacon roll": 8.25}


# --- reshaping --------------------------------------------------------------

def test_rows_as_they_sit(out):
    assert out["sales_rows"][0] == ["Region", "Q1", "Q2"]
    assert out["sales_rows"][1] == ["North", 1200, 1350]


def test_records(out):
    assert out["sales_records"] == [
        {"Region": "North", "Q1": 1200, "Q2": 1350},
        {"Region": "South", "Q1": 980, "Q2": 1010},
        {"Region": "East", "Q1": 1440, "Q2": 1390},
    ]


def test_columns(out):
    assert out["sales_columns"] == {
        "Region": ["North", "South", "East"],
        "Q1": [1200, 980, 1440],
        "Q2": [1350, 1010, 1390],
    }


def test_transpose(out):
    assert out["sales_transposed"][0] == ["Region", "North", "South", "East"]
    assert out["sales_transposed"][1] == ["Q1", 1200, 980, 1440]


def test_flatten(out):
    assert out["staff_flat"] == ["name", "Ada", "Grace", "role", "Barista", "Chef"]


def test_chained_pipeline(out):
    assert out["staff"] == [
        {"name": "Ada", "role": "Barista"},
        {"name": "Grace", "role": "Chef"},
    ]


# --- object functions -------------------------------------------------------

def test_keys_values_items(out):
    assert out["price_list"] == ["Tea", "Coffee", "Bacon roll"]
    assert out["price_values"] == [3.25, 4.0, 8.25]
    assert out["price_items"] == [["Tea", 3.25], ["Coffee", 4.0], ["Bacon roll", 8.25]]


def test_invert(out):
    assert out["code_names"] == {"Gluten-free": "GF", "Vegan": "VG"}


# --- coercion and formatting ------------------------------------------------

def test_int_coercion_of_text_numbers(out):
    assert out["covers"] == {"Mon": 128, "Tue": 143, "Wed": 97}


def test_float_coercion(out):
    assert out["covers_float"] == {"Mon": 128.0, "Tue": 143.0, "Wed": 97.0}


def test_str_coercion(out):
    assert out["prices_as_text"] == {"Tea": "3.25", "Coffee": "4", "Bacon roll": "8.25"}


def test_round_with_argument(out):
    """AvgSpend holds 8.7451; round is a built-in that takes a digits argument."""
    assert out["average_spend"] == 8.75   # round 2
    assert out["spend_rounded"] == 8.75   # round, digits defaulting to 2
    assert out["spend_whole"] == 9.0      # round 0


def test_isodate(out):
    assert out["period_ending"] == "2026-03-31"


def test_dates_serialise_without_isodate(out):
    """A raw date cell reaches JSON as ISO-8601 too, and agrees with isodate."""
    encoded = json.loads(json.dumps(out["period_raw"], default=json_default))
    assert encoded == "2026-03-31"
    assert encoded == out["period_ending"]


# --- arguments (no built-in takes them; injected functions do) --------------

def test_quoted_string_argument(tmp_path):
    spec = tmp_path / "args.toml"
    spec.write_text('[extract]\ntag_line = \'[Tags | join ", "]\'\n')
    out = read_spec(XLSX, spec, functions={"join": lambda v, sep: sep.join(v)})
    assert out == {"tag_line": "Vegan, Gluten-free, Dairy-free"}


def test_numeric_and_reference_arguments(tmp_path):
    spec = tmp_path / "args.toml"
    spec.write_text("[extract]\nrow = '[Sales | pick 1]'\n")
    out = read_spec(XLSX, spec, functions={"pick": lambda table, n: table[n]})
    assert out == {"row": ["North", 1200, 1350]}
