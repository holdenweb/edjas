"""Functions applied by the EDJAS pipeline notation.

An extraction expression may append a pipeline of functions, separated by ``|``
and applied left to right after the value is extracted -- for example
``[Sales | records]``, ``{Prices | int}``, or, on a scalar, ``Total | round 2``.
Each function receives the value produced so far as its first positional argument,
followed by any pipeline arguments (a number, a double-quoted string, or a bare
word naming another range); so ``[Price | round 2]`` calls ``round(Price, 2)``.
Functions are looked up in a controlled registry rather than evaluated, so a spec
can only invoke registered functions and never executes arbitrary code.

Each function returns a JSON-serialisable result. :func:`resolve` merges
caller-supplied functions over the built-in :data:`DEFAULT_FUNCTIONS` and
:func:`lookup` fetches one by name; :func:`json_default` is the ``json`` encoder
hook that lets date/time cells serialise instead of raising.
"""

from datetime import date, datetime, time

__all__ = ["DEFAULT_FUNCTIONS", "resolve", "lookup", "json_default"]


# --- helpers ---------------------------------------------------------------

def _is_table(value):
    """True if ``value`` is a non-empty list whose every element is a list."""
    return (
        isinstance(value, list)
        and len(value) > 0
        and all(isinstance(row, list) for row in value)
    )


def _require_table(value, func_name):
    if not _is_table(value):
        raise ValueError(
            f"{func_name} expects a 2-D range (a list of rows); "
            f"got {type(value).__name__}"
        )


def _require_dict(value, func_name):
    if not isinstance(value, dict):
        raise ValueError(
            f"{func_name} expects an object (use {{name}} markup); "
            f"got {type(value).__name__}"
        )


def _map_scalars(fn, value):
    """Apply ``fn`` to every scalar in a vector, table, or object's values."""
    if isinstance(value, list):
        return [_map_scalars(fn, item) for item in value]
    if isinstance(value, dict):
        return {key: _map_scalars(fn, val) for key, val in value.items()}
    return fn(value)


def _coercer(cast):
    """Build a function that coerces every (non-None) scalar via ``cast``."""

    def scalar(x):
        return x if x is None else cast(x)

    return lambda value: _map_scalars(scalar, value)


# --- reshape functions -----------------------------------------------------

def records(table):
    """Treat row 0 as headers; return rows 1..n as a list of objects."""
    _require_table(table, "records")
    header, *rows = table
    return [dict(zip(header, row)) for row in rows]


def columns(table):
    """Treat row 0 as headers; return ``{header: [column values], ...}``."""
    _require_table(table, "columns")
    header, *rows = table
    cols = list(zip(*rows)) if rows else [()] * len(header)
    return {head: list(col) for head, col in zip(header, cols)}


def transpose(table):
    """Swap rows and columns of a 2-D range."""
    _require_table(table, "transpose")
    return [list(row) for row in zip(*table)]


def flatten(value):
    """Flatten an arbitrarily nested vector/table into a single vector."""
    out = []

    def walk(node):
        if isinstance(node, list):
            for item in node:
                walk(item)
        else:
            out.append(node)

    walk(value)
    return out


# --- object functions ------------------------------------------------------

def keys(obj):
    _require_dict(obj, "keys")
    return list(obj)


def values(obj):
    _require_dict(obj, "values")
    return list(obj.values())


def items(obj):
    _require_dict(obj, "items")
    return [[key, val] for key, val in obj.items()]


def invert(obj):
    _require_dict(obj, "invert")
    return {val: key for key, val in obj.items()}


# --- coercion / formatting -------------------------------------------------

def _iso(value):
    """ISO-8601 text for a date/time, dropping a time component of exact midnight.

    Spreadsheets store a date-only cell as midnight on that day, so rendering it as
    ``2026-03-31`` rather than ``2026-03-31T00:00:00`` reflects what was meant.
    """
    if isinstance(value, datetime) and value.time() == time(0, 0):
        return value.date().isoformat()
    return value.isoformat()


def _isodate_scalar(x):
    return _iso(x) if isinstance(x, (date, time)) else x


def round_values(value, digits=2):
    """Round every floating-point value to ``digits`` decimal places."""

    def scalar(x):
        return round(x, digits) if isinstance(x, float) else x

    return _map_scalars(scalar, value)


isodate = lambda value: _map_scalars(_isodate_scalar, value)  # noqa: E731


DEFAULT_FUNCTIONS = {
    # reshape
    "records": records,
    "columns": columns,
    "transpose": transpose,
    "flatten": flatten,
    # object
    "keys": keys,
    "values": values,
    "items": items,
    "invert": invert,
    # coercion / formatting
    "int": _coercer(int),
    "float": _coercer(float),
    "str": _coercer(str),
    "round": round_values,
    "isodate": isodate,
}


def resolve(functions=None):
    """Return the default registry, overlaid with caller-supplied functions."""
    if not functions:
        return dict(DEFAULT_FUNCTIONS)
    return {**DEFAULT_FUNCTIONS, **functions}


def lookup(registry, name):
    """Return ``registry[name]`` or raise a helpful ``ValueError``."""
    try:
        return registry[name]
    except KeyError:
        available = ", ".join(sorted(registry))
        raise ValueError(
            f"Unknown EDJAS function {name!r}; available: {available}"
        ) from None


def json_default(obj):
    """``json`` encoder hook: serialise date/datetime/time as ISO-8601.

    A datetime at exactly midnight becomes a plain date, matching ``isodate``.
    """
    if isinstance(obj, (date, time)):
        return _iso(obj)
    raise TypeError(
        f"Object of type {type(obj).__name__} is not JSON serializable"
    )
