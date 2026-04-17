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


# ── Number formatting and unit conversion ────────────────────────────────

FMT_FINANCIAL = "#,##0.0_ ;[Red](#,##0.0)"
FMT_EPS       = "#,##0.00_ ;[Red](#,##0.00)"
FMT_SHARES    = "#,##0"


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


def _apply_number_formats(ws) -> None:
    """Convert values to millions and apply number formats. Skips section/blank rows."""
    for row_idx in range(3, ws.max_row + 1):
        concept = str(ws.cell(row=row_idx, column=1).value or "").strip()

        # Section headers and blank separators have no numeric data
        if concept in SECTION_HEADERS or concept == "":
            continue

        is_eps    = _is_eps_concept(concept)
        is_shares = "Shares" in concept

        if is_eps:
            fmt = FMT_EPS
            divisor = 1
        elif is_shares:
            fmt = FMT_SHARES
            divisor = 1_000_000
        else:
            fmt = FMT_FINANCIAL
            divisor = 1_000_000

        for col_idx in range(3, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell.value, (int, float)):
                cell.value = cell.value / divisor
                cell.number_format = fmt


# ── Index sheet ───────────────────────────────────────────────────────────

def _build_index_sheet(wb: Workbook, tables: list) -> None:
    """Insert or replace the Index sheet at position 0."""
    if "Index" in wb.sheetnames:
        del wb["Index"]

    ticker       = tables[0].ticker if tables else ""
    company_name = ""
    meta = next((t for t in tables if t.sheet_name == "Data_Meta"), None)
    if meta and len(meta.concepts) > 1 and len(meta.values) > 1 and meta.values[1]:
        company_name = meta.values[1][0] or ""

    header_text = f"{ticker} — {company_name}" if company_name else ticker

    ws = wb.create_sheet("Index", 0)

    # Row 1: company header
    ws["A1"] = header_text
    ws["A1"].fill = _fill(NAVY_DARK)
    ws["A1"].font = Font(color="FFFFFFFF", bold=True, size=14)
    ws.merge_cells("A1:D1")

    # Row 2: metadata
    ws["A2"] = f"抓取日期：{date.today()}　　資料來源：SEC EDGAR"
    ws["A2"].fill = _fill(NAVY_MID)
    ws["A2"].font = Font(color="FFAABBCC", size=9)
    ws.merge_cells("A2:D2")

    # Row 3: blank
    ws.row_dimensions[3].height = 6

    # Row 4: column headers
    hdr_font = Font(bold=True, size=10)
    hdr_fill = _fill(BLUE_HDR)
    for col, label in enumerate(["Sheet", "說明", "最早期間", "最新期間"], start=1):
        cell = ws.cell(row=4, column=col, value=label)
        cell.font = hdr_font
        cell.fill = hdr_fill

    # Row 5+: one row per Data_* sheet
    data_sheets = [t for t in tables if t.sheet_name.startswith("Data_")]
    for i, tbl in enumerate(data_sheets):
        row = 5 + i
        earliest = tbl.quarter_labels[0]  if tbl.quarter_labels else "—"
        latest   = tbl.quarter_labels[-1] if tbl.quarter_labels else "—"

        is_primary = tbl.sheet_name in ("Data_Financials(Q)", "Data_Financials(Y)")
        name_font  = Font(color="FF1F3864" if is_primary else "FF666666",
                          bold=is_primary, size=10)
        row_fill   = _fill(ROW_WHITE) if row % 2 == 0 else _fill(ROW_ALT)

        for col, val in enumerate([tbl.sheet_name,
                                    _sheet_description(tbl.sheet_name),
                                    earliest, latest], start=1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = row_fill
            if col == 1:
                cell.font = name_font

    # Column widths
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12


# ── Public API ────────────────────────────────────────────────────────────

def format_workbook(wb: Workbook, tables: list[StatementTable]) -> None:
    """Apply formatting to all Data_* sheets."""
    _build_index_sheet(wb, tables)
    for ws in wb.worksheets:
        if not ws.title.startswith("Data_"):
            continue
        _apply_column_widths(ws)
        _set_freeze_panes(ws)
        _apply_row_styles(ws)
        if ws.title != "Data_Meta":
            _apply_number_formats(ws)
