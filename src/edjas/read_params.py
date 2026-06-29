import sys

import openpyxl

from . import functions as _functions


def parse_ref(inner):
    """Split markup contents into an optional function token and a range ref.

    ``"records Sales"`` -> ``("records", "Sales")``; ``"Sales"`` -> ``(None,
    "Sales")``. Splitting on whitespace is unambiguous because Excel named ranges
    and cell references (e.g. ``D3:E9``) never contain spaces.
    """
    parts = inner.split(None, 1)
    if len(parts) == 2:
        return parts[0], parts[1].strip()
    return None, parts[0].strip()


def extract_values(sheet, range_spec, flatten=True):
    result = []
    for row in sheet[range_spec]:
        result.append([c.value for c in row])
    if not flatten:
        return result
    if len(result) == 1:
        return result[0]
    elif len(result[0]) == 1:
        return [r[0] for r in result]
    else:
        return result

def range_values(wb, range_spec, flatten=True):
    if range_spec in wb.defined_names:
        range_spec = wb.defined_names[range_spec].attr_text
    if "!" in range_spec:
        sheet_name, cell_refs = range_spec.split("!")
        sheet = wb[sheet_name]
    else:
        sheet = wb.active
        cell_refs = range_spec
    return extract_values(sheet, cell_refs, flatten=flatten)

def _apply(func, value, registry):
    """Apply the named function from ``registry`` to ``value`` (or pass through)."""
    if func is None:
        return value
    return _functions.lookup(registry, func)(value)


def range_to_dict(workbook, range_spec, registry=None):
    if registry is None:
        registry = _functions.resolve()
    # Get the rows in the given range
    rows = range_values(workbook, range_spec, flatten=False)
    if len(rows[0]) != 2:
        raise ValueError(f"Range spec {range_spec} should have two columns")
    # Initialize the result dictionary
    result = {}
    i = 0
    for key, value in rows:
        # Skip empty rows, but complain about floating values
        if key is None:
            if value is None:
                continue
            else:
                raise ValueError("Empty key not expected on value {value!r} - programming error?")
        # Check if the value is a range name enclosed in braces: dictionary
        if type(value) is str:
            if value.startswith("{") and value.endswith("}"):
                # "{[f] name}": extract an object, optionally apply function f.
                func, ref_range_spec = parse_ref(value[1:-1])
                extracted = range_to_dict(workbook, ref_range_spec, registry)
                result[key] = _apply(func, extracted, registry)
            # Otherwise "[[f] name]" references a list or matrix.
            elif value.startswith("[") and value.endswith("]"):
                func, ref_range_spec = parse_ref(value[1:-1])
                extracted = range_values(workbook, ref_range_spec)
                result[key] = _apply(func, extracted, registry)
            else:
                result[key] = value
        else:
            # Single value
            result[key] = value
    return result

def read_file(file_name, range_name="Parameters", functions=None):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(file_name, data_only=False)
    registry = _functions.resolve(functions)
    return range_to_dict(workbook, range_name, registry)

if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit("Requires spreadsheet arguments")
    if len(sys.argv) > 3:
        sys.exit("Sorry, only handling one or two arguments for now")
    from pprint import pprint
    data = read_file(*sys.argv[1:])
    pprint(data)
