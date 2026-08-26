"""Tests for reading OpenDocument spreadsheets (``edjas.ods``).

Fixtures are built here rather than committed, because an ODS file is only a zip of
XML and writing it out by hand keeps the awkward cases -- repeated cells, covered
cells, padding runs -- visible in the test that depends on them. That matters most for
``table:number-columns-repeated``: it is how ODS compresses runs of identical cells,
and mishandling it shifts a whole sheet sideways without any error.
"""

import datetime as dt
import zipfile

import pytest

from edjas import read_spec
from edjas.ods import is_ods, load_workbook

CONTENT = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body><office:spreadsheet>
{sheets}
{names}
 </office:spreadsheet></office:body>
</office:document-content>
"""


def cell(value=None, kind="string", repeat=None, covered=False):
    """One ``table:table-cell`` (or a covered cell) as XML."""
    tag = "table:covered-table-cell" if covered else "table:table-cell"
    times = f' table:number-columns-repeated="{repeat}"' if repeat else ""
    if value is None:
        return f"<{tag}{times}/>"
    if kind == "string":
        return f"<{tag}{times} office:value-type=\"string\"><text:p>{value}</text:p></{tag}>"
    attribute = {
        "float": "office:value", "percentage": "office:value", "currency": "office:value",
        "date": "office:date-value", "time": "office:time-value",
        "boolean": "office:boolean-value",
    }[kind]
    return f'<{tag}{times} office:value-type="{kind}" {attribute}="{value}"/>'


def row(cells, repeat=None):
    times = f' table:number-rows-repeated="{repeat}"' if repeat else ""
    return f"<table:table-row{times}>{''.join(cells)}</table:table-row>"


def group(rows, kind="table-row-group"):
    """Wrap rows the way print titles and outline grouping do."""
    return f"<table:{kind}>{''.join(rows)}</table:{kind}>"


def make_ods(tmp_path, sheets, named_ranges=None, name="book.ods"):
    """Write a minimal .ods. ``sheets`` maps a sheet name to a list of row XML."""
    body = "".join(
        f'<table:table table:name="{title}">{"".join(rows)}</table:table>'
        for title, rows in sheets.items()
    )
    names = ""
    if named_ranges:
        entries = "".join(
            f'<table:named-range table:name="{key}" table:cell-range-address="{ref}"/>'
            for key, ref in named_ranges.items()
        )
        names = f"<table:named-expressions>{entries}</table:named-expressions>"
    path = tmp_path / name
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
        archive.writestr("content.xml", CONTENT.format(sheets=body, names=names))
    return path


# --- recognising the format --------------------------------------------------

def test_is_ods_by_extension():
    assert is_ods("a.ods") and is_ods("A.ODS") and is_ods("/tmp/x.Ods")
    assert not is_ods("a.xlsx") and not is_ods("a.ods.xlsx")


# --- values and types --------------------------------------------------------

def test_reads_each_value_type(tmp_path):
    path = make_ods(tmp_path, {"S": [row([
        cell("hello"),
        cell("42", "float"),
        cell("3.5", "float"),
        cell("true", "boolean"),
        cell("2026-03-31", "date"),
        cell("0.25", "percentage"),
    ])]})
    sheet = load_workbook(path)["S"]
    assert [c.value for c in sheet[1]] == [
        "hello", 42, 3.5, True, dt.datetime(2026, 3, 31), 0.25,
    ]


def test_whole_floats_become_ints(tmp_path):
    """ODS stores every number as a float; a whole one should not read as 42.0."""
    path = make_ods(tmp_path, {"S": [row([cell("42", "float")])]})
    value = load_workbook(path)["S"]["A1"].value
    assert value == 42 and isinstance(value, int)


def test_reads_time_values(tmp_path):
    path = make_ods(tmp_path, {"S": [row([cell("PT08H30M00S", "time")])]})
    assert load_workbook(path)["S"]["A1"].value == dt.time(8, 30)


def test_empty_cells_are_none(tmp_path):
    path = make_ods(tmp_path, {"S": [row([cell(), cell("x")])]})
    sheet = load_workbook(path)["S"]
    assert sheet["A1"].value is None and sheet["B1"].value == "x"


# --- the repeat compression --------------------------------------------------

def test_repeated_columns_shift_later_cells(tmp_path):
    """A run of repeated cells occupies every column it claims.

    Ignoring the repeat count would place "end" in column C instead of column F, which
    is how a whole sheet ends up read one column at a time out of alignment.
    """
    path = make_ods(tmp_path, {"S": [row([
        cell("start"), cell("mid", repeat=4), cell("end"),
    ])]})
    sheet = load_workbook(path)["S"]
    assert sheet["A1"].value == "start"
    assert [sheet.cell(row=1, column=c).value for c in range(2, 6)] == ["mid"] * 4
    assert sheet["F1"].value == "end"


def test_empty_repeated_cells_still_advance_the_column(tmp_path):
    path = make_ods(tmp_path, {"S": [row([cell("a"), cell(repeat=3), cell("b")])]})
    sheet = load_workbook(path)["S"]
    assert sheet["A1"].value == "a" and sheet["E1"].value == "b"


def test_repeated_rows_are_materialised(tmp_path):
    path = make_ods(tmp_path, {"S": [row([cell("x")], repeat=3), row([cell("last")])]})
    sheet = load_workbook(path)["S"]
    assert [sheet.cell(row=r, column=1).value for r in range(1, 5)] == [
        "x", "x", "x", "last",
    ]


def test_padding_to_the_full_grid_costs_nothing(tmp_path):
    """ODS pads sheets out to the whole grid; those runs must not be materialised.

    Real files routinely claim 1,048,576 rows and 16,384 columns this way. Creating
    those cells would exhaust memory for a sheet holding two values.
    """
    path = make_ods(tmp_path, {"S": [
        row([cell("only"), cell(repeat=16000)]),
        row([cell()], repeat=1048000),
    ]})
    sheet = load_workbook(path)["S"]
    assert sheet["A1"].value == "only"
    assert sheet.max_row == 1 and sheet.max_column == 1


def test_covered_cells_hold_no_value_but_take_their_place(tmp_path):
    """A cell hidden under a merge still occupies its column."""
    path = make_ods(tmp_path, {"S": [row([
        cell("merged"), cell(covered=True, repeat=2), cell("after"),
    ])]})
    sheet = load_workbook(path)["S"]
    assert sheet["A1"].value == "merged"
    assert sheet["B1"].value is None and sheet["C1"].value is None
    assert sheet["D1"].value == "after"


# --- rows nested in ODF's wrappers -------------------------------------------

@pytest.mark.parametrize(
    "kind", ["table-header-rows", "table-rows", "table-row-group"]
)
def test_rows_inside_a_wrapper_keep_their_place(tmp_path, kind):
    """Print titles and outline groups nest rows; they are still rows of the sheet.

    Skipping them would not merely lose those rows: every row below would move up,
    silently, so a spec reading A2 would get what is really in A3.
    """
    path = make_ods(tmp_path, {"S": [
        group([row([cell("wrapped")])], kind),
        row([cell("after")]),
    ]})
    sheet = load_workbook(path)["S"]
    assert sheet["A1"].value == "wrapped"
    assert sheet["A2"].value == "after"


def test_nested_row_groups_are_flattened_in_order(tmp_path):
    path = make_ods(tmp_path, {"S": [
        row([cell("top")]),
        group([row([cell("outer")]), group([row([cell("inner")])])]),
        row([cell("bottom")]),
    ]})
    sheet = load_workbook(path)["S"]
    assert [sheet.cell(row=r, column=1).value for r in range(1, 5)] == [
        "top", "outer", "inner", "bottom",
    ]


def test_a_table_nested_in_a_cell_is_not_hoisted_into_the_sheet(tmp_path):
    """A table may legally sit inside a cell; its rows are not rows of this sheet.

    This is why the row walk descends only into the three row wrappers. A blanket
    search for table-row would pull the sub-table's rows up into the parent, which is
    the same misalignment in the opposite direction.
    """
    inner = (
        '<table:table table:name="Nested">'
        f"{row([cell('sub')])}"
        "</table:table>"
    )
    outer = (
        '<table:table-row><table:table-cell office:value-type="string">'
        f"<text:p>host</text:p>{inner}"
        "</table:table-cell></table:table-row>"
    )
    path = make_ods(tmp_path, {"S": [outer, row([cell("next")])]})
    workbook = load_workbook(path)
    assert workbook.sheetnames == ["S"]          # the sub-table is not a sheet
    assert workbook["S"]["A1"].value == "host"   # nor is its text part of the value
    assert workbook["S"]["A2"].value == "next"


# --- text that is not plain character data -----------------------------------

def test_explicit_spaces_and_tabs_are_preserved(tmp_path):
    """Indentation survives: ODS encodes runs of spaces as <text:s>."""
    body = (
        '<table:table-cell office:value-type="string">'
        '<text:p><text:s text:c="4"/>Education<text:tab/>total</text:p>'
        "</table:table-cell>"
    )
    path = make_ods(tmp_path, {"S": [f"<table:table-row>{body}</table:table-row>"]})
    assert load_workbook(path)["S"]["A1"].value == "    Education\ttotal"


def test_a_cell_comment_is_not_part_of_the_value(tmp_path):
    """An office:annotation holds an author, a date and a note -- none of them data."""
    body = (
        '<table:table-cell office:value-type="string">'
        "<office:annotation><text:p>check this figure</text:p></office:annotation>"
        "<text:p>Adur</text:p></table:table-cell>"
    )
    path = make_ods(tmp_path, {"S": [f"<table:table-row>{body}</table:table-row>"]})
    assert load_workbook(path)["S"]["A1"].value == "Adur"


# --- sheets and named ranges -------------------------------------------------

def test_multiple_sheets_keep_their_names(tmp_path):
    path = make_ods(tmp_path, {
        "Front Page": [row([cell("cover")])],
        "Data": [row([cell("1", "float")])],
    })
    workbook = load_workbook(path)
    assert workbook.sheetnames == ["Front Page", "Data"]
    assert workbook["Front Page"]["A1"].value == "cover"


def test_named_ranges_translate_to_excel_form(tmp_path):
    """ODS writes Sheet.$A$1:Sheet.$B$1; EDJAS resolves 'Sheet'!$A$1:$B$1."""
    path = make_ods(
        tmp_path,
        {"Data": [row([cell("a"), cell("b")])]},
        named_ranges={"Pair": "Data.$A$1:Data.$B$1"},
    )
    workbook = load_workbook(path)
    assert "Pair" in workbook.defined_names
    assert workbook.defined_names["Pair"].attr_text == "'Data'!$A$1:$B$1"


def test_absolute_sheet_marker_is_stripped_from_a_named_range(tmp_path):
    """LibreOffice writes $Sheet.$A$1; the $ is not part of the sheet's name."""
    path = make_ods(
        tmp_path,
        {"Data": [row([cell("a")])]},
        named_ranges={"Cell": "$Data.$A$1"},
    )
    workbook = load_workbook(path)
    assert workbook.defined_names["Cell"].attr_text == "'Data'!$A$1"


def test_a_range_spanning_two_sheets_is_declined(tmp_path):
    """A 3-D range has no single-rectangle equivalent, so it is not registered.

    Reading it as the same shape on the first sheet alone would return plausible,
    wrong data -- the failure mode multi-area references are refused to avoid.
    """
    path = make_ods(
        tmp_path,
        {"One": [row([cell("a")])], "Two": [row([cell("b")])]},
        named_ranges={"Across": "One.$A$1:Two.$A$1"},
    )
    assert "Across" not in load_workbook(path).defined_names


def test_sheet_scoped_names_do_not_reach_the_workbook(tmp_path):
    """A name defined inside a sheet is local to it, and must not collide globally."""
    sheets = {
        "First": [row([cell("a")]),
                  '<table:named-expressions><table:named-range table:name="Local"'
                  ' table:cell-range-address="First.$A$1"/></table:named-expressions>'],
        "Second": [row([cell("b")])],
    }
    path = make_ods(tmp_path, sheets, named_ranges={"Global": "Second.$A$1"})
    workbook = load_workbook(path)
    assert "Global" in workbook.defined_names
    assert "Local" not in workbook.defined_names


def test_named_range_on_a_sheet_whose_name_has_a_space(tmp_path):
    path = make_ods(
        tmp_path,
        {"Cover Sheet": [row([cell("title")])]},
        named_ranges={"Title": "'Cover Sheet'.$A$1"},
    )
    workbook = load_workbook(path)
    assert workbook.defined_names["Title"].attr_text == "'Cover Sheet'!$A$1"


# --- end to end through the ordinary API -------------------------------------

def test_read_spec_reads_an_ods_workbook(tmp_path):
    """The whole expression language works against ODS, unchanged."""
    path = make_ods(
        tmp_path,
        {
            "Cover Sheet": [row([cell("Quarterly report")])],
            "Sales": [
                row([cell("Region"), cell("Q1")]),
                row([cell("North"), cell("1200", "float")]),
                row([cell("South"), cell("980", "float")]),
            ],
        },
        named_ranges={"Sales": "Sales.$A$1:Sales.$B$3"},
    )
    spec = tmp_path / "spec.toml"
    spec.write_text(
        "[extract]\n"
        "title = \"'Cover Sheet'!A1\"\n"
        "rows  = '[Sales | records]'\n"
        "total = '[Sales!B2:B3 | flatten]'\n"
    )
    assert read_spec(path, spec) == {
        "title": "Quarterly report",
        "rows": [{"Region": "North", "Q1": 1200}, {"Region": "South", "Q1": 980}],
        "total": [1200, 980],
    }


def test_ods_workbook_is_not_modified(tmp_path):
    """Reading an ODS leaves it byte-for-byte as it was, exactly as for xlsx."""
    path = make_ods(tmp_path, {"S": [row([cell("x")])]})
    before = path.read_bytes()
    load_workbook(path)
    assert path.read_bytes() == before


# --- failure modes -----------------------------------------------------------

def test_a_file_that_is_not_a_zip_is_reported_clearly(tmp_path):
    path = tmp_path / "fake.ods"
    path.write_text("this is not a spreadsheet")
    with pytest.raises(ValueError, match="not a zip archive"):
        load_workbook(path)


def test_a_zip_without_content_is_reported_clearly(tmp_path):
    path = tmp_path / "empty.ods"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
    with pytest.raises(ValueError, match="no content.xml"):
        load_workbook(path)


def test_a_spreadsheet_with_no_sheets_is_reported_clearly(tmp_path):
    path = make_ods(tmp_path, {})
    with pytest.raises(ValueError, match="no sheets"):
        load_workbook(path)


def test_unparseable_content_is_reported_clearly(tmp_path):
    """A damaged content.xml must not escape as a raw ParseError.

    read_spec and the command line both handle ValueError; anything else reaches the
    user as a traceback.
    """
    path = tmp_path / "broken.ods"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("content.xml", "<office:document-content><unclosed>")
    with pytest.raises(ValueError, match="not valid XML"):
        load_workbook(path)


def test_a_file_demanding_absurdly_many_cells_is_refused(tmp_path):
    """A repeated row that *has* content is materialised, so it needs a ceiling.

    Empty runs are free, but a few hundred bytes claiming a million copies of a
    populated row would otherwise ask for more memory than the machine has.
    """
    path = make_ods(tmp_path, {"S": [
        row([cell("x"), cell("y"), cell("z")], repeat=2_000_000),   # 6M cells
    ]})
    with pytest.raises(ValueError, match="probably malformed"):
        load_workbook(path)


def test_a_large_but_plausible_sheet_is_still_read(tmp_path):
    """The ceiling must not reject a genuinely big spreadsheet."""
    path = make_ods(tmp_path, {"S": [row([cell("v")] * 20, repeat=1000)]})
    sheet = load_workbook(path)["S"]
    assert sheet.max_row == 1000 and sheet.max_column == 20


def test_a_malformed_repeat_count_does_not_crash(tmp_path):
    path = make_ods(tmp_path, {"S": [
        f'<table:table-row><table:table-cell table:number-columns-repeated="wat"'
        f' office:value-type="string"><text:p>a</text:p></table:table-cell>'
        f"{cell('b')}</table:table-row>"
    ]})
    sheet = load_workbook(path)["S"]
    assert sheet["A1"].value == "a" and sheet["B1"].value == "b"


def test_a_malformed_number_keeps_the_cell(tmp_path):
    path = make_ods(tmp_path, {"S": [row([cell("not-a-number", "float"), cell("ok")])]})
    sheet = load_workbook(path)["S"]
    assert sheet["A1"].value == "not-a-number" and sheet["B1"].value == "ok"


def test_durations_of_a_day_or_more_are_not_wrapped(tmp_path):
    """26h30m is not 02:30. A duration that is not a time of day is left as written."""
    path = make_ods(tmp_path, {"S": [row([cell("PT26H30M00S", "time")])]})
    assert load_workbook(path)["S"]["A1"].value == "PT26H30M00S"
