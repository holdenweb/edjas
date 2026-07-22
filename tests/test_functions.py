import json
from datetime import date, datetime, time

import pytest

from edjas import json_default
from edjas.functions import DEFAULT_FUNCTIONS, lookup, resolve


def fn(name):
    return DEFAULT_FUNCTIONS[name]


# --- reshape ---------------------------------------------------------------

def test_records_header_row_to_objects():
    table = [["Region", "Q1"], ["North", 100], ["South", 200]]
    assert fn("records")(table) == [
        {"Region": "North", "Q1": 100},
        {"Region": "South", "Q1": 200},
    ]


def test_records_requires_2d():
    with pytest.raises(ValueError):
        fn("records")(["not", "a", "table"])


def test_columns_groups_by_header():
    table = [["Region", "Q1"], ["North", 100], ["South", 200]]
    assert fn("columns")(table) == {"Region": ["North", "South"], "Q1": [100, 200]}


def test_transpose():
    assert fn("transpose")([[1, 2, 3], [4, 5, 6]]) == [[1, 4], [2, 5], [3, 6]]


def test_flatten_nested():
    assert fn("flatten")([[1, 2], [3, [4, 5]]]) == [1, 2, 3, 4, 5]


# --- object functions ------------------------------------------------------

def test_keys_values_items():
    d = {"a": 1, "b": 2}
    assert fn("keys")(d) == ["a", "b"]
    assert fn("values")(d) == [1, 2]
    assert fn("items")(d) == [["a", 1], ["b", 2]]


def test_invert():
    assert fn("invert")({"x": 1}) == {1: "x"}


def test_object_function_rejects_non_dict():
    with pytest.raises(ValueError):
        fn("keys")([1, 2, 3])


# --- coercion / formatting -------------------------------------------------

def test_int_coerces_vector():
    assert fn("int")(["1", "2", "3"]) == [1, 2, 3]


def test_float_coerces_table_cells():
    assert fn("float")([["1", "2"], ["3", "4"]]) == [[1.0, 2.0], [3.0, 4.0]]


def test_str_coerces_dict_values_not_keys():
    assert fn("str")({"a": 1, "b": 2}) == {"a": "1", "b": "2"}


def test_coercion_passes_through_none():
    assert fn("int")([1, None, 3]) == [1, None, 3]


def test_round_defaults_to_two_places():
    assert fn("round")([1.23456, 2.0, "x"]) == [1.23, 2.0, "x"]


def test_round_takes_a_digits_argument():
    assert fn("round")([1.23456], 3) == [1.235]
    assert fn("round")([1.23456], 0) == [1.0]


def test_round_leaves_non_floats_alone():
    assert fn("round")({"a": 7, "b": "text", "c": None}) == {"a": 7, "b": "text", "c": None}


def test_isodate_scalars_and_passthrough():
    assert fn("isodate")(date(2026, 6, 29)) == "2026-06-29"
    assert fn("isodate")([datetime(2026, 6, 29, 9, 30), "open"]) == [
        "2026-06-29T09:30:00",
        "open",
    ]


def test_isodate_drops_midnight_time():
    """A date-only cell arrives as midnight; render it as a plain date."""
    assert fn("isodate")(datetime(2026, 3, 31, 0, 0)) == "2026-03-31"
    # a real midnight-adjacent time is still shown in full
    assert fn("isodate")(datetime(2026, 3, 31, 0, 1)) == "2026-03-31T00:01:00"


# --- registry plumbing -----------------------------------------------------

def test_resolve_overlays_custom_functions():
    registry = resolve({"double": lambda v: [x * 2 for x in v]})
    assert registry["double"]([1, 2]) == [2, 4]
    assert "records" in registry  # built-ins still present


def test_resolve_none_returns_defaults_copy():
    registry = resolve()
    registry["extra"] = object()
    assert "extra" not in DEFAULT_FUNCTIONS  # did not mutate the shared dict


def test_lookup_unknown_raises_with_available_names():
    with pytest.raises(ValueError, match="Unknown EDJAS function 'nope'"):
        lookup(DEFAULT_FUNCTIONS, "nope")


# --- json_default ----------------------------------------------------------

def test_json_default_serialises_dates():
    payload = {"d": date(2026, 6, 29), "t": time(9, 30), "dt": datetime(2026, 6, 29, 9, 30)}
    out = json.loads(json.dumps(payload, default=json_default))
    assert out == {"d": "2026-06-29", "t": "09:30:00", "dt": "2026-06-29T09:30:00"}


def test_json_default_drops_midnight_time_like_isodate():
    """Automatic serialisation and isodate must agree on date-only cells."""
    midnight = datetime(2026, 3, 31, 0, 0)
    assert json.loads(json.dumps(midnight, default=json_default)) == "2026-03-31"
    assert fn("isodate")(midnight) == "2026-03-31"


def test_json_default_still_raises_for_unknown_types():
    with pytest.raises(TypeError):
        json.dumps({"x": object()}, default=json_default)
