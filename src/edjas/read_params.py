import sys

import openpyxl

from . import functions as _functions


def _scan(inner):
    """Tokenise pipeline text into pipe separators and tagged word tokens.

    Returns a list of ``("pipe", None, False)`` and ``("tok", text, is_string)``
    items. A double-quoted substring becomes one token with ``is_string=True`` and
    its quotes stripped; whitespace and ``|`` inside quotes are literal.
    """
    tokens = []
    buf = []
    is_string = False
    have_token = False
    in_quote = False

    def flush():
        nonlocal buf, is_string, have_token
        if have_token:
            tokens.append(("tok", "".join(buf), is_string))
        buf = []
        is_string = False
        have_token = False

    for ch in inner:
        if in_quote:
            if ch == '"':
                in_quote = False
            else:
                buf.append(ch)
            continue
        if ch == '"':
            in_quote = True
            is_string = True
            have_token = True  # an empty "" is still a (string) token
        elif ch == "|":
            flush()
            tokens.append(("pipe", None, False))
        elif ch.isspace():
            flush()
        else:
            buf.append(ch)
            have_token = True
    flush()
    if in_quote:
        raise ValueError(f"Unterminated string in markup: {inner!r}")
    return tokens


def parse_pipeline(inner):
    """Parse ``source | func args | func args`` into a ref and a list of stages.

    Returns ``(source_ref, [(func_name, [(arg_text, is_string), ...]), ...])``.
    The leading segment must be a single range reference; each subsequent
    ``|``-delimited stage is a function name followed by space-separated args.
    """
    segments = [[]]
    for kind, text, is_string in _scan(inner):
        if kind == "pipe":
            segments.append([])
        else:
            segments[-1].append((text, is_string))

    source = segments[0]
    if len(source) != 1 or source[0][1]:  # exactly one, unquoted, token
        raise ValueError(f"Markup must start with a single range reference: {inner!r}")
    source_ref = source[0][0]

    stages = []
    for seg in segments[1:]:
        if not seg:
            raise ValueError(f"Empty function stage in markup: {inner!r}")
        func_name = seg[0][0]
        stages.append((func_name, seg[1:]))
    return source_ref, stages


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

def _resolve_arg(workbook, text, is_string):
    """Resolve a stage argument: quoted -> string, numeric -> number, else a range."""
    if is_string:
        return text
    try:
        return int(text)
    except ValueError:
        pass
    try:
        return float(text)
    except ValueError:
        pass
    # A bare word is a range reference, read exactly as ``[word]`` would be.
    return range_values(workbook, text)


def apply_pipeline(workbook, registry, value, stages):
    """Apply each ``(func_name, args)`` stage to ``value`` left to right."""
    for func_name, args in stages:
        func = _functions.lookup(registry, func_name)
        resolved = [_resolve_arg(workbook, text, is_string) for text, is_string in args]
        value = func(value, *resolved)
    return value


def range_to_dict(workbook, range_spec, registry=None):
    if registry is None:
        registry = _functions.resolve()
    # Get the rows in the given range
    rows = range_values(workbook, range_spec, flatten=False)
    if len(rows[0]) != 2:
        raise ValueError(f"Range spec {range_spec} should have two columns")
    # Initialize the result dictionary
    result = {}
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
                # "{name | f ... }": extract an object, then apply the pipeline.
                source_ref, stages = parse_pipeline(value[1:-1])
                extracted = range_to_dict(workbook, source_ref, registry)
                result[key] = apply_pipeline(workbook, registry, extracted, stages)
            # "[name | f ... ]" extracts a list/matrix, then applies the pipeline.
            elif value.startswith("[") and value.endswith("]"):
                source_ref, stages = parse_pipeline(value[1:-1])
                extracted = range_values(workbook, source_ref)
                result[key] = apply_pipeline(workbook, registry, extracted, stages)
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
