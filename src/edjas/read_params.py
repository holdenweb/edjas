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

def _resolve_sheet_and_refs(wb, range_spec):
    """Resolve a named range or A1-style spec to ``(sheet, cell_refs)``.

    A defined name is expanded to its ``Sheet!refs`` text; an explicit ``!`` selects
    that sheet; otherwise the active sheet is used.
    """
    if range_spec in wb.defined_names:
        range_spec = wb.defined_names[range_spec].attr_text
    if "!" in range_spec:
        sheet_name, cell_refs = range_spec.split("!")
        sheet = wb[sheet_name]
    else:
        sheet = wb.active
        cell_refs = range_spec
    return sheet, cell_refs


def range_values(wb, range_spec, flatten=True):
    sheet, cell_refs = _resolve_sheet_and_refs(wb, range_spec)
    return extract_values(sheet, cell_refs, flatten=flatten)


def read_scalar(wb, range_spec):
    """Return the value of a single referenced cell (top-left if a range is given)."""
    sheet, cell_refs = _resolve_sheet_and_refs(wb, range_spec)
    target = sheet[cell_refs]
    if isinstance(target, tuple):  # a multi-cell range -> take the top-left cell
        return target[0][0].value
    return target.value  # a lone cell reference is a single Cell

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


def range_to_dict(workbook, range_spec):
    """Read a two-column range into an object: column 0 keys, column 1 values.

    Blank rows are skipped; a value with no key is an error.
    """
    rows = range_values(workbook, range_spec, flatten=False)
    if len(rows[0]) != 2:
        raise ValueError(f"Range spec {range_spec} should have two columns")
    result = {}
    for key, value in rows:
        if key is None:
            if value is None:
                continue
            raise ValueError(f"Value {value!r} in {range_spec} has no key")
        result[key] = value
    return result


def evaluate(workbook, expr, registry):
    """Evaluate one spec expression against the workbook.

    ``[ref | ...]`` extracts a list/table, ``{ref | ...}`` an object, and a bare
    ``ref`` a scalar. Any trailing ``| func`` stages are then applied.
    """
    expr = expr.strip()
    if expr.startswith("[") and expr.endswith("]"):
        source_ref, stages = parse_pipeline(expr[1:-1])
        value = range_values(workbook, source_ref)
    elif expr.startswith("{") and expr.endswith("}"):
        source_ref, stages = parse_pipeline(expr[1:-1])
        value = range_to_dict(workbook, source_ref)
    else:
        source_ref, stages = parse_pipeline(expr)
        value = read_scalar(workbook, source_ref)
    return apply_pipeline(workbook, registry, value, stages)
