from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet


def column(sheet: Worksheet) -> Cell:
    """
    Return the first cell of the first non-blank column in the sheet.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet to search.

    Returns:
        openpyxl.cell.Cell: The first non-blank cell scanning column by column,
        or None if every cell is blank.
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

    Returns:
        openpyxl.cell.Cell: The first non-blank cell scanning row by row,
        or None if every cell is blank.
    """
    for row in sheet.iter_rows():
        for cell  in row:
            if cell is not None and cell.value:  # Check if the cell is not empty
                return cell

def top_left(sheet: Worksheet) -> Cell:
    """
    Return the top-left cell of the sheet's used data range.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet to search.

    Returns:
        openpyxl.cell.Cell: The cell at the top-left corner of the used range.
    """
    data_range = sheet.calculate_dimension()
    return sheet[data_range][0][0]
