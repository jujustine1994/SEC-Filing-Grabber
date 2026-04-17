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
