"""Tests for excel_formatter.py."""
import pytest
from openpyxl import Workbook
from excel_formatter import format_workbook, FMT_FINANCIAL, FMT_EPS, FMT_SHARES


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


# ── number formatting + unit conversion ────────────────────────────────────

def test_revenue_divided_by_million():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # C4 = "Revenue" row, raw = 117154000000
    assert ws["C4"].value == pytest.approx(117154.0)

def test_revenue_second_col_divided():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert ws["D4"].value == pytest.approx(94836.0)

def test_revenue_number_format():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"]["C4"].number_format == FMT_FINANCIAL

def test_eps_not_divided():
    wb = _make_wb()
    format_workbook(wb, [])
    # C6 = "Basic EPS" = 1.52 → must stay 1.52
    assert wb["Data_Financials(Q)"]["C6"].value == pytest.approx(1.52)

def test_eps_number_format():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"]["C6"].number_format == FMT_EPS

def test_shares_divided_by_million():
    wb = _make_wb()
    format_workbook(wb, [])
    # C7 = "Basic Shares" = 15787000000 → 15787.0
    assert wb["Data_Financials(Q)"]["C7"].value == pytest.approx(15787.0)

def test_shares_number_format():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"]["C7"].number_format == FMT_SHARES

def test_section_header_values_unchanged():
    wb = _make_wb()
    format_workbook(wb, [])
    # C3 = "Income Statement" row — all None
    assert wb["Data_Financials(Q)"]["C3"].value is None

def test_data_meta_values_not_converted():
    """Data_Meta contains strings — must not be divided."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Data_Meta"
    ws["A1"] = "AAPL"
    ws["C1"] = "FY2023Q1"
    ws["A3"] = "Ticker"
    ws["C3"] = "AAPL"
    format_workbook(wb, [])
    assert wb["Data_Meta"]["C3"].value == "AAPL"

def test_seg_sheet_financial_converted():
    wb = _make_wb(sheet_name="Data_Seg_Revenue")
    wb["Data_Seg_Revenue"]["A4"] = "Americas"
    wb["Data_Seg_Revenue"]["C4"] = 50000000000.0
    format_workbook(wb, [])
    assert wb["Data_Seg_Revenue"]["C4"].value == pytest.approx(50000.0)


# ── Index sheet ────────────────────────────────────────────────────────────

from fetcher_gaap import StatementTable

def _make_tables(sheet_name="Data_Financials(Q)", ticker="AAPL",
                 qs=None, dates=None):
    qs    = qs    or ["FY2020Q1", "FY2024Q4"]
    dates = dates or ["2020-04-30", "2025-01-30"]
    return [StatementTable(
        sheet_name=sheet_name, ticker=ticker,
        quarter_labels=qs, filing_dates=dates,
        concepts=["Revenue"], values=[[100.0, 200.0]], labels=["Revenues"],
    )]


def test_index_sheet_created():
    wb = _make_wb()
    format_workbook(wb, _make_tables())
    assert "Index" in wb.sheetnames

def test_index_sheet_is_first():
    wb = _make_wb()
    format_workbook(wb, _make_tables())
    assert wb.sheetnames[0] == "Index"

def test_index_sheet_ticker_in_a1():
    wb = _make_wb()
    format_workbook(wb, _make_tables(ticker="AAPL"))
    ws = wb["Index"]
    assert "AAPL" in str(ws["A1"].value)

def test_index_lists_data_sheet():
    wb = _make_wb()
    format_workbook(wb, _make_tables())
    ws = wb["Index"]
    col_a_values = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    assert "Data_Financials(Q)" in col_a_values

def test_index_shows_earliest_period():
    wb = _make_wb()
    format_workbook(wb, _make_tables(qs=["FY2010Q1", "FY2024Q4"]))
    ws = wb["Index"]
    all_values = [ws.cell(row=r, column=c).value
                  for r in range(1, ws.max_row + 1)
                  for c in range(1, 5)]
    assert "FY2010Q1" in all_values

def test_index_shows_latest_period():
    wb = _make_wb()
    format_workbook(wb, _make_tables(qs=["FY2010Q1", "FY2024Q4"]))
    ws = wb["Index"]
    all_values = [ws.cell(row=r, column=c).value
                  for r in range(1, ws.max_row + 1)
                  for c in range(1, 5)]
    assert "FY2024Q4" in all_values

def test_index_not_deleted_on_reformat():
    """Index must not be deleted by excel_writer (doesn't start with Data_)."""
    wb = _make_wb()
    format_workbook(wb, _make_tables())
    # Simulate a second write: excel_writer deletes Data_* sheets
    for name in list(wb.sheetnames):
        if name.startswith("Data_"):
            del wb[name]
    assert "Index" in wb.sheetnames
