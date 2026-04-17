"""
excel_formatter.py — Apply professional formatting to Data_* sheets and generate Index sheet.

Public API:
    format_workbook(wb, tables) -> None

Called by excel_writer.write_statements() before wb.save(). Modifies cell values
(÷1M unit conversion) and applies openpyxl styles. Does not change sheet structure.
"""

from __future__ import annotations
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from fetcher_gaap import StatementTable
from datetime import date

# ── Colours (ARGB) ────────────────────────────────────────────────────────
NAVY_DARK = "FF1F3864"
NAVY_MID  = "FF2D4A82"
BLUE_MID  = "FF2E75B6"
GREY_SEP  = "FFEEEEEE"
ROW_ALT   = "FFF5F8FF"
ROW_WHITE = "FFFFFFFF"
BLUE_HDR  = "FFDDE8F5"

# ── Row classification ────────────────────────────────────────────────────
SECTION_HEADERS = {"Income Statement", "Balance Sheet", "Cash Flow"}

SUBTOTAL_CONCEPTS = {
    "Gross Profit", "Total Operating Expense", "Operating Income",
    "Pre-tax Income", "Net Income",
    "Total Current Assets", "Total Assets",
    "Total Current Liabilities", "Total Liabilities", "Total Equity",
    "Operating Cash Flow", "Free Cash Flow",
}

SHEET_DESCRIPTIONS = {
    "Data_Financials(Q)": "季報三表合一（IS + BS + CF，from 10-Q）",
    "Data_Financials(Y)": "年報三表合一（IS + BS + CF，from 10-K）",
    "Data_EPS_Recon":     "Non-GAAP EPS 調節表（from 8-K）",
    "Data_NonGAAP":       "Non-GAAP 指標（AI 提取）",
    "Data_Meta":          "申報資訊（Ticker、公司名、抓取日期）",
}


def _is_eps_concept(c: str) -> bool:
    return any(k in (c or "") for k in ("EPS", "Per Share", "per share"))


def _sheet_description(name: str) -> str:
    if name in SHEET_DESCRIPTIONS:
        return SHEET_DESCRIPTIONS[name]
    if name.startswith("Data_Seg_"):
        return f"Segment 細項：{name[9:]}"
    return name


def _fill(hex_argb: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_argb)


# ── Column widths ─────────────────────────────────────────────────────────

def _apply_column_widths(ws) -> None:
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 24
    for col in range(3, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].width = 13


# ── Freeze panes ──────────────────────────────────────────────────────────

def _set_freeze_panes(ws) -> None:
    ws.freeze_panes = "C3"


# ── Row styles ────────────────────────────────────────────────────────────

def _apply_row_styles(ws) -> None:
    """Apply fill and font styles to all rows."""
    white_font  = Font(color="FFFFFFFF", bold=True, size=11)
    small_font  = Font(color="FFAABBCC", size=9)

    # Row 1: ticker / quarter labels — dark navy
    for cell in ws[1]:
        cell.fill = _fill(NAVY_DARK)
        cell.font = white_font

    # Row 2: filing dates — medium navy
    for cell in ws[2]:
        cell.fill = _fill(NAVY_MID)
        cell.font = small_font

    # Row 3+: classify by col A value
    for row_idx in range(3, ws.max_row + 1):
        concept = ws.cell(row=row_idx, column=1).value or ""
        concept = str(concept).strip()

        if concept in SECTION_HEADERS:
            row_fill  = _fill(BLUE_MID)
            row_font  = Font(color="FFFFFFFF", bold=True, size=10)
            row_height = 16
        elif concept == "":
            row_fill  = _fill(GREY_SEP)
            row_font  = Font(size=9)
            row_height = 6
        else:
            row_fill  = _fill(ROW_WHITE) if row_idx % 2 == 0 else _fill(ROW_ALT)
            bold      = concept in SUBTOTAL_CONCEPTS
            row_font  = Font(bold=bold) if bold else Font()
            row_height = None

        for cell in ws[row_idx]:
            cell.fill = row_fill
            cell.font = row_font
        if row_height is not None:
            ws.row_dimensions[row_idx].height = row_height


# ── Public API ────────────────────────────────────────────────────────────

def format_workbook(wb: Workbook, tables: list[StatementTable]) -> None:
    """Apply formatting to all Data_* sheets."""
    for ws in wb.worksheets:
        if not ws.title.startswith("Data_"):
            continue
        _apply_column_widths(ws)
        _set_freeze_panes(ws)
        _apply_row_styles(ws)
