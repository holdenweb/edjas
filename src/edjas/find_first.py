from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet


def column(sheet: Worksheet) -> Cell:
    """
    Return the first cell of the first non-blank column in the sheet.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet to search.

    Returns:
        int: The index of the first non-blank column (1-based indexing), or None if all columns are blank.
    """
    for col in sheet.iter_cols():
        for cell in col:
            if cell is not None and cell.value:  # Check if the cell is not empty
                return cell

def row(sheet: Worksheet) -> Cell:
    """
    Return the first cell of the first non-blank row in the sheet.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet to search.

    """
    for row in sheet.iter_rows():
        for cell  in row:
            if cell is not None and cell.value:  # Check if the cell is not empty
                return cell

def top_left(sheet: Worksheet):
    """
    Returns the cell
    """
    data_range = sheet.calculate_dimension()
    return sheet[data_range][0][0]
