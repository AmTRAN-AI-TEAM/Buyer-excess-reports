from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from excess_builder import ROW1_HEADERS, ROW2_HEADERS


def write_excess_workbook(rows: list[list[Any]], output_file: Path) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Excess"

    worksheet.append(ROW1_HEADERS)
    worksheet.append(ROW2_HEADERS)
    for row in rows:
        worksheet.append(row)

    _apply_layout(worksheet, len(rows) + 2)
    workbook.save(output_file)


def _apply_layout(worksheet: Any, last_row: int) -> None:
    worksheet.merge_cells("AM1:AQ1")
    worksheet.auto_filter.ref = f"A2:AZ{max(last_row, 2)}"
    worksheet.freeze_panes = "A3"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    group_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(color="FFFFFF", bold=True)
    group_font = Font(bold=True)
    thin_gray = Side(style="thin", color="D9D9D9")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    for row in worksheet.iter_rows(min_row=1, max_row=max(last_row, 2), max_col=52):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    for cell in worksheet[1]:
        if cell.value is not None:
            cell.fill = group_fill
            cell.font = group_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for cell in worksheet[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    widths = {
        "A": 16,
        "B": 16,
        "C": 16,
        "D": 16,
        "E": 16,
        "F": 16,
        "G": 16,
        "H": 42,
        "I": 14,
        "J": 14,
        "K": 10,
        "L": 10,
        "M": 12,
        "AR": 18,
        "AS": 42,
        "AT": 18,
        "AU": 22,
        "AV": 28,
        "AW": 16,
        "AX": 14,
        "AY": 18,
        "AZ": 12,
    }

    for column in range(1, 53):
        letter = get_column_letter(column)
        worksheet.column_dimensions[letter].width = widths.get(letter, 12)

    worksheet.row_dimensions[1].height = 18
    worksheet.row_dimensions[2].height = 34

    for column in range(14, 44):
        for row in range(3, last_row + 1):
            worksheet.cell(row, column).number_format = "#,##0.00"

    for column in range(1, 8):
        for row in range(3, last_row + 1):
            worksheet.cell(row, column).number_format = "@"
