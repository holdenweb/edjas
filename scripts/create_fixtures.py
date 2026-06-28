"""Regenerate the test fixture used by the EDJAS test suite.

Running this script rewrites ``tests/data/parameters.xlsx`` with the exact
layout that ``tests/test_read_params.py`` asserts against:

  * a two-column ``Parameters`` range whose values exercise scalars, ``{dict}``
    references, ``[range]`` references, and direct cell-range references;
  * a ``hours`` named range (a 7-row day/opening-hours table);
  * a ``prices`` named range living on a second ``SubParameters`` sheet.

Run from the project root:  ``python scripts/create_fixtures.py``
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

FIXTURE = Path(__file__).resolve().parent.parent / "tests" / "data" / "parameters.xlsx"

HOURS = [
    ("Monday", "7:00 am - 8:00 pm"),
    ("Tuesday", "7:00 am - 8:00 pm"),
    ("Wednesday", "7:00 am - 8:00 pm"),
    ("Thursday", "7:00 am - 8:00 pm"),
    ("Friday", "7:00 am - 8:00 pm"),
    ("Saturday", "9:00 am - 5:00 pm"),
    ("Sunday", "Closed"),
]
H_VECTOR = ["the", "quick", "brown", "fox"]
V_VECTOR = ["jumps", "over", "the", "lazy", "dog"]
PRICES = [("Tea", 3.25), ("Coffee", 4.0), ("Bacon Sandwich", 8.25)]


def build():
    wb = Workbook()

    params = wb.active
    params.title = "Parameters"

    # Column A/B: the parameter specifications EDJAS reads.
    params["A3"], params["B3"] = "hours", "{hours}"
    params["A4"], params["B4"] = "prices", "{prices}"
    params["A5"], params["B5"] = "title", "Hubris Demo"
    params["A6"], params["B6"] = "hours_list", "[D3:E9]"
    params["A7"], params["B7"] = "h_vector", "[H5:K5]"
    params["A8"], params["B8"] = "v_vector", "[H3:H7]"

    # Columns D/E: the hours table, referenced by {hours} and [D3:E9].
    for offset, (day, opening) in enumerate(HOURS):
        params.cell(row=3 + offset, column=4, value=day)
        params.cell(row=3 + offset, column=5, value=opening)

    # Column H rows 3-7: the vertical vector, referenced by [H3:H7].
    for offset, word in enumerate(V_VECTOR):
        params.cell(row=3 + offset, column=8, value=word)

    # Row 5 columns H-K: the horizontal vector, referenced by [H5:K5].
    for offset, word in enumerate(H_VECTOR):
        params.cell(row=5, column=8 + offset, value=word)

    sub = wb.create_sheet("SubParameters")
    for offset, (item, price) in enumerate(PRICES):
        sub.cell(row=1 + offset, column=1, value=item)
        sub.cell(row=1 + offset, column=2, value=price)

    wb.defined_names.add(DefinedName("Parameters", attr_text="Parameters!$A$1:$B$15"))
    wb.defined_names.add(DefinedName("hours", attr_text="Parameters!$D$3:$E$9"))
    wb.defined_names.add(DefinedName("prices", attr_text="SubParameters!$A$1:$B$3"))

    FIXTURE.parent.mkdir(parents=True, exist_ok=True)
    wb.save(FIXTURE)
    print(f"Wrote fixture: {FIXTURE}")


if __name__ == "__main__":
    build()
