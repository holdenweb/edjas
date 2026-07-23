from . import functions as _functions


def _scan(inner):
    """Tokenise pipeline text into pipe separators and tagged word tokens.

    Returns a list of ``("pipe", None, False)`` and ``("tok", text, is_string)``
    items. Two kinds of quoting group characters into a token so that whitespace and
    ``|`` inside them are literal:

    - **Double quotes** delimit a string *argument*: the quotes are stripped and the
      token is tagged ``is_string=True`` (an empty ``""`` is still a token).
    - **Single quotes** are Excel's *sheet-name* quoting (``'Cover Sheet'!A1``): the
      quotes are kept (so the sheet resolver can strip them) and the token stays a
      plain reference (``is_string=False``). A doubled ``''`` is an escaped apostrophe,
      not a closing quote, matching how Excel writes ``'Bob''s Data'``.
    """
    tokens = []
    buf = []
    is_string = False
    have_token = False

    def flush():
        nonlocal buf, is_string, have_token
        if have_token:
            tokens.append(("tok", "".join(buf), is_string))
        buf = []
        is_string = False
        have_token = False

    i, n = 0, len(inner)
    while i < n:
        ch = inner[i]
        if ch == '"':  # a double-quoted string argument; quotes stripped
            is_string = True
            have_token = True
            i += 1
            while i < n and inner[i] != '"':
                buf.append(inner[i])
                i += 1
            if i >= n:
                raise ValueError(f"Unterminated string in markup: {inner!r}")
            i += 1  # skip the closing quote
        elif ch == "'":  # Excel sheet-name quoting; quotes kept, '' escapes a quote
            have_token = True
            buf.append("'")
            i += 1
            while i < n:
                if inner[i] == "'":
                    if i + 1 < n and inner[i + 1] == "'":
                        buf.append("''")  # doubled apostrophe: literal, stay in quote
                        i += 2
                        continue
                    buf.append("'")
                    i += 1
                    break
                buf.append(inner[i])
                i += 1
            else:
                raise ValueError(f"Unterminated sheet name in markup: {inner!r}")
        elif ch == "|":
            flush()
            tokens.append(("pipe", None, False))
            i += 1
        elif ch.isspace():
            flush()
            i += 1
        else:
            buf.append(ch)
            have_token = True
            i += 1
    flush()
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

def _unquote_sheet(sheet_name):
    """Strip Excel's quoting from a sheet name (`'Bob''s Data'` -> `Bob's Data`)."""
    if sheet_name.startswith("'") and sheet_name.endswith("'"):
        return sheet_name[1:-1].replace("''", "'")
    return sheet_name


def _lookup_defined_name(wb, name):
    """Return the defined name matching ``name``, case-insensitively like Excel.

    An exact-case match is preferred; otherwise the first case-insensitive match is
    used (Excel forbids defined names that differ only in case). Returns None if the
    name is not defined.
    """
    if name in wb.defined_names:
        return wb.defined_names[name]
    lowered = name.lower()
    for key, defined_name in wb.defined_names.items():
        if key.lower() == lowered:
            return defined_name
    return None


def _get_sheet(wb, sheet_name):
    """Fetch a worksheet by name, falling back to a case-insensitive match."""
    try:
        return wb[sheet_name]
    except KeyError:
        lowered = sheet_name.lower()
        for title in wb.sheetnames:
            if title.lower() == lowered:
                return wb[title]
        raise


def _lookup_table(wb, name):
    """Return ``(sheet, ref)`` for an Excel Table named ``name``, else None.

    Excel Tables (``ws.tables``) are per-sheet, not workbook-level like defined names,
    so every sheet is scanned. ``ws.tables.items()`` yields ``(name, "A5:D7")`` — the
    value is the A1 range string, which already spans the table's header and data rows.
    Matching is case-insensitive, as Excel treats table names. This is how GSS
    good-practice government spreadsheets label their data (named Tables, not names).
    """
    lowered = name.lower()
    for sheet in wb.worksheets:
        for table_name, ref in sheet.tables.items():
            if table_name.lower() == lowered:
                return sheet, ref
    return None


def _resolve_sheet_and_refs(wb, range_spec):
    """Resolve a named range or A1-style spec to ``(sheet, cell_refs)``.

    A defined name is expanded to its ``Sheet!refs`` text; an Excel Table name resolves
    directly to ``(sheet, ref)``; an explicit ``!`` selects that sheet; otherwise the
    active sheet is used. Defined names, table names and sheet names are matched
    case-insensitively, as Excel does. Excel wraps a sheet name that contains spaces or
    other special characters in single quotes (doubling any embedded quote), so the
    quotes are stripped before the sheet is looked up.

    Precedence is defined name, then table name, then a raw A1/sheet reference. Excel
    keeps defined names and table names in one namespace and forbids collisions, so the
    order rarely matters; defined-name-first leaves existing behaviour untouched.
    """
    defined = _lookup_defined_name(wb, range_spec)
    if defined is not None:
        range_spec = defined.attr_text
    elif "!" not in range_spec and ":" not in range_spec:
        # A bare identifier that is not a defined name may be an Excel Table name.
        table = _lookup_table(wb, range_spec)
        if table is not None:
            return table
    if "," in range_spec:  # a union of areas: EDJAS reads a single rectangle only
        raise ValueError(f"multi-area (union) ranges are not supported: {range_spec}")
    if "!" in range_spec:
        # A sheet name never contains '!', so split on the first one only.
        sheet_name, cell_refs = range_spec.split("!", 1)
        sheet = _get_sheet(wb, _unquote_sheet(sheet_name))
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
