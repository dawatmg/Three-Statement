"""
Helper utilities for professional Excel formatting.
"""

from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side,
    numbers,
)
from openpyxl.utils import get_column_letter

from config import (
    HEADER_FILL_COLOR,
    HEADER_FONT_COLOR,
    NUMBER_FORMAT,
    SECTION_FILL_COLOR,
    SUBTOTAL_FILL_COLOR,
    ALTERNATING_ROW_COLOR,
    LABEL_COLUMN_WIDTH,
    DATA_COLUMN_WIDTH,
)


# ---------------------------------------------------------------------------
# Style factories
# ---------------------------------------------------------------------------

def _fill(hex_color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=hex_color)


def _border_top() -> Border:
    thin = Side(style="thin")
    return Border(top=thin)


def apply_header_style(cell) -> None:
    """Dark-blue background, bold white text, centred."""
    cell.fill = _fill(HEADER_FILL_COLOR)
    cell.font = Font(bold=True, color=HEADER_FONT_COLOR, size=11)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def apply_section_style(cell) -> None:
    """Light-blue background, bold dark text."""
    cell.fill = _fill(SECTION_FILL_COLOR)
    cell.font = Font(bold=True, size=10)
    cell.alignment = Alignment(horizontal="left", vertical="center")


def apply_subtotal_style(cell, is_label: bool = False) -> None:
    """Very light-blue background, bold text, top border."""
    cell.fill = _fill(SUBTOTAL_FILL_COLOR)
    cell.font = Font(bold=True, size=10)
    cell.border = _border_top()
    if is_label:
        cell.alignment = Alignment(horizontal="left", vertical="center")
    else:
        cell.alignment = Alignment(horizontal="right", vertical="center")


def apply_label_style(cell, row_index: int) -> None:
    """Regular row – alternating background for readability."""
    if row_index % 2 == 0:
        cell.fill = _fill(ALTERNATING_ROW_COLOR)
    cell.font = Font(size=10)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def apply_number_style(cell, row_index: int) -> None:
    """Right-aligned number with thousands separator."""
    if row_index % 2 == 0:
        cell.fill = _fill(ALTERNATING_ROW_COLOR)
    cell.font = Font(size=10)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.number_format = NUMBER_FORMAT


# ---------------------------------------------------------------------------
# Sheet helpers
# ---------------------------------------------------------------------------

def write_header_row(ws, columns: list[str], row: int = 1) -> None:
    """Write and style the column header row."""
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=row, column=col_idx, value=col_name)
        apply_header_style(cell)
    ws.row_dimensions[row].height = 30


def set_column_widths(ws, num_data_cols: int) -> None:
    """Set label column width and uniform data column widths."""
    ws.column_dimensions["A"].width = LABEL_COLUMN_WIDTH
    for col_idx in range(2, num_data_cols + 2):
        ws.column_dimensions[get_column_letter(col_idx)].width = DATA_COLUMN_WIDTH


def freeze_header(ws, row: int = 2, col: int = 2) -> None:
    """Freeze the header row and label column."""
    ws.freeze_panes = ws.cell(row=row, column=col)


def write_data_row(
    ws,
    row_idx: int,
    label: str,
    values: list,
    style: str = "normal",
) -> None:
    """
    Write a single data row.

    Parameters
    ----------
    style : str
        One of ``"normal"``, ``"section"``, or ``"subtotal"``.
    """
    label_cell = ws.cell(row=row_idx, column=1, value=label)

    if style == "section":
        apply_section_style(label_cell)
    elif style == "subtotal":
        apply_subtotal_style(label_cell, is_label=True)
    else:
        apply_label_style(label_cell, row_idx)

    for col_idx, val in enumerate(values, start=2):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        if style == "section":
            apply_section_style(cell)
            cell.value = None  # section rows have no numeric data
        elif style == "subtotal":
            apply_subtotal_style(cell)
            apply_number_style(cell, row_idx)
        else:
            apply_number_style(cell, row_idx)
