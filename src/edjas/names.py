"""List what a workbook offers as references, for someone writing a specification.

Writing a specification begins with the same question every time: what can I actually
refer to in this file? A named range and an Excel Table are both usable as a reference
in an extraction expression, but they are stored in different places -- names are
workbook-wide, tables belong to a sheet -- so neither Excel nor openpyxl shows both in
one list.

Both are listed here, with the range each resolves to and its shape, because the shape
is what decides which of the three extraction forms to write.

Real workbooks carry a great deal of dead name, and it is worth knowing that this is
normal rather than damage. A statistical release is usually last year's file with the
figures changed; sheets get deleted over the years and the names pointing into them do
not, so Excel keeps the name and rewrites its target to ``#REF!``. A defined name may
also hold a constant or an array rather than a reference at all. Neither means the
sheets are broken, and both are reported as what they are rather than being quietly
omitted.
"""

import collections
import re
import sys

from openpyxl.utils import range_boundaries
from openpyxl.utils.exceptions import CellCoordinatesException

from .read_params import _unquote_sheet

__all__ = ["describe", "report"]


def _shape(refs):
    """``(rows, columns)`` covered by an A1 range, or None if it will not parse."""
    try:
        left, top, right, bottom = range_boundaries(str(refs).replace("$", ""))
    except (ValueError, TypeError, CellCoordinatesException):
        return None
    if None in (left, top, right, bottom):
        return None
    return bottom - top + 1, right - left + 1


def _suggest(shape):
    """The extraction form that fits a range of this shape."""
    if shape is None:
        return "ref"
    rows, columns = shape
    if rows == 1 and columns == 1:
        return "ref"
    if columns == 2:
        return "{ref}"
    if rows > 1 and columns > 1:
        return "[ref | records]"
    return "[ref]"


def _classify(target, workbook):
    """Why a defined name cannot be used as a reference, or None if it can."""
    if not target:
        return "empty definition"
    if "#REF" in target:
        return "points at cells that were deleted (#REF!)"
    if target.startswith(("{", '"')) or target.lstrip("-").replace(".", "", 1).isdigit():
        # A stored constant or array, not a reference. These often contain a '!' inside
        # a quoted string, so they must be recognised before any sheet is looked for.
        return "holds a constant, not a range"
    if "," in target:
        return "multi-area (union) range"
    if "!" not in target:
        return "no sheet in the reference"
    sheet = _unquote_sheet(target.split("!", 1)[0])
    if sheet not in workbook.sheetnames:
        return f"names a sheet that is not in this workbook ({sheet!r})"
    return None


def _defined_names(workbook):
    found = []
    for name, entry in workbook.defined_names.items():
        target = entry.attr_text or ""
        problem = _classify(target, workbook)
        sheet, refs = None, target
        if problem is None:
            sheet_name, refs = target.split("!", 1)
            sheet = _unquote_sheet(sheet_name)
        found.append({
            "name": name, "sheet": sheet, "refs": refs, "target": target,
            "shape": None if problem else _shape(refs), "problem": problem,
        })
    return found


def _tables(workbook):
    return [
        {
            "name": name, "sheet": sheet.title, "refs": refs,
            "target": f"{sheet.title}!{refs}", "shape": _shape(refs), "problem": None,
        }
        for sheet in workbook.worksheets
        for name, refs in sheet.tables.items()
    ]


def describe(workbook):
    """Everything referenceable in an open workbook, as plain data."""
    return {
        "sheets": list(workbook.sheetnames),
        "names": _defined_names(workbook),
        "tables": _tables(workbook),
    }


def _section(entries, heading, verbose, out):
    usable = [e for e in entries if not e["problem"]]
    dead = [e for e in entries if e["problem"]]
    print(f"  {heading}", file=out)
    if not entries:
        print("    none", file=out)
        return

    if usable:
        width = max(len(e["name"]) for e in usable)
        for entry in sorted(usable, key=lambda e: e["name"].lower()):
            print(f"    {entry['name']:<{width}}  {entry['sheet']}!{entry['refs']}", file=out)
            if entry["shape"]:
                rows, columns = entry["shape"]
                form = _suggest(entry["shape"]).replace("ref", entry["name"])
                print(f"    {'':<{width}}  {rows} x {columns}   e.g.  {form}", file=out)
    elif dead:
        print("    none usable", file=out)

    if not dead:
        return
    if verbose:
        width = max(len(e["name"]) for e in dead)
        for entry in sorted(dead, key=lambda e: e["name"].lower()):
            print(f"    {entry['name']:<{width}}  {entry['problem']}", file=out)
        return
    # A workbook copied forward for years accumulates dead names in the hundreds, and
    # listing them buries whatever is usable, so they are counted by reason instead.
    reasons = collections.Counter(
        re.sub(r"\(.*\)", "", e["problem"]).strip() for e in dead
    )
    summary = ", ".join(f"{count} {reason}" for reason, count in reasons.most_common())
    print(f"    ({len(dead)} unusable: {summary} -- use --verbose to list them)", file=out)


def report(workbook, path, verbose=False, out=None):
    """Print what ``workbook`` offers as references."""
    out = out or sys.stdout
    found = describe(workbook)
    print(path, file=out)
    print(f"  {len(found['sheets'])} sheet(s): {', '.join(found['sheets'])}", file=out)
    _section(found["names"], "Named ranges", verbose, out)
    _section(found["tables"], "Excel Tables", verbose, out)
