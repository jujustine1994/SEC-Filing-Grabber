"""Tests for excel_formatter.py."""
import pytest
from openpyxl import Workbook
from excel_formatter import format_workbook


def _make_wb(sheet_name="Data_Financials(Q)"):
    """Minimal workbook with one Data_* sheet and two data columns."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws["A1"] = "AAPL"
    ws["C1"] = "FY2023Q1"
    ws["D1"] = "FY2023Q2"
    ws["C2"] = "2023-02-03"
    ws["D2"] = "2023-05-05"
    ws["A3"] = "Income Statement"   # section header
    ws["A4"] = "Revenue"
    ws["B4"] = "Revenues"
    ws["C4"] = 117154000000.0
    ws["D4"] = 94836000000.0
    ws["A5"] = ""                   # blank separator
    ws["A6"] = "Basic EPS"
    ws["C6"] = 1.52
    ws["D6"] = 1.20
    ws["A7"] = "Basic Shares"
    ws["C7"] = 15787000000.0
    ws["D7"] = 15813000000.0
    return wb


# ── column widths ──────────────────────────────────────────────────────────

def test_col_a_width():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"].column_dimensions["A"].width == 22

def test_col_b_width():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"].column_dimensions["B"].width == 24

def test_data_col_width():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"].column_dimensions["C"].width == 13

def test_data_col_d_width():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"].column_dimensions["D"].width == 13


# ── freeze panes ───────────────────────────────────────────────────────────

def test_freeze_panes():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"].freeze_panes == "C3"

def test_freeze_panes_seg_sheet():
    wb = _make_wb(sheet_name="Data_Seg_Revenue")
    format_workbook(wb, [])
    assert wb["Data_Seg_Revenue"].freeze_panes == "C3"


# ── row styles ─────────────────────────────────────────────────────────────

def _rgb(ws, cell_ref: str) -> str:
    """Return fgColor ARGB string of a cell's fill."""
    return ws[cell_ref].fill.fgColor.rgb


def test_row1_fill_navy_dark():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert _rgb(ws, "A1") == "FF1F3864"

def test_row1_font_bold_white():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert ws["A1"].font.bold is True
    assert ws["A1"].font.color.rgb == "FFFFFFFF"

def test_row2_fill_navy_mid():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert _rgb(ws, "A2") == "FF2D4A82"

def test_section_header_fill_blue_mid():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # A3 = "Income Statement"
    assert _rgb(ws, "A3") == "FF2E75B6"

def test_section_header_font_bold_white():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert ws["A3"].font.bold is True
    assert ws["A3"].font.color.rgb == "FFFFFFFF"

def test_blank_separator_fill_grey():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # A5 = ""
    assert _rgb(ws, "A5") == "FFEEEEEE"

def test_section_header_row_height():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert ws.row_dimensions[3].height == 16

def test_blank_separator_row_height():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert ws.row_dimensions[5].height == 6

def test_data_row_alternating_white():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # row_idx=4, even → ROW_WHITE
    assert _rgb(ws, "A4") == "FFFFFFFF"

def test_data_row_alternating_blue():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # row_idx=7, odd → ROW_ALT
    assert _rgb(ws, "A7") == "FFF5F8FF"

def test_subtotal_row_bold():
    wb = _make_wb()
    # Add a Gross Profit row
    wb["Data_Financials(Q)"]["A8"] = "Gross Profit"
    wb["Data_Financials(Q)"]["C8"] = 5000000000.0
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert ws["A8"].font.bold is True

def test_normal_row_not_bold():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # A4 = "Revenue" (not a subtotal)
    assert ws["A4"].font.bold is not True
