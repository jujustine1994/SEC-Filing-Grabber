"""Tests for comparison_writer.py — 跨公司比較 Excel 輸出。"""
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

from comparison import ComparisonResult
from comparison_writer import (
    write_compare_data_sheet,
    write_snapshot_sheets,
    write_chart_sheets,
    write_comparison_workbook,
)


def _sample_result():
    return ComparisonResult(
        metrics={
            "Revenue": {
                "NVDA": {"FY2024Q1": 100.0, "FY2024Q2": 110.0},
                "AMD": {"FY2024Q1": 80.0, "FY2024Q2": 90.0},
            },
        },
        period_ends={
            "NVDA": {"FY2024Q1": "2024-03-31", "FY2024Q2": "2024-06-30"},
            "AMD": {"FY2024Q1": "2024-03-31", "FY2024Q2": "2024-06-30"},
        },
        failures=[],
    )


# ── Compare_Data ─────────────────────────────────────────────────────────

def test_compare_data_sheet_has_metric_header_and_period_columns():
    wb = Workbook()
    result = _sample_result()
    write_compare_data_sheet(wb, result, ["Revenue"])
    ws = wb["Compare_Data"]

    assert ws["A1"].value == "Revenue"
    header_row = [c.value for c in ws[2]]
    assert "FY2024Q1" in header_row
    assert "FY2024Q2" in header_row


def test_compare_data_sheet_has_static_period_end_row():
    wb = Workbook()
    result = _sample_result()
    write_compare_data_sheet(wb, result, ["Revenue"])
    ws = wb["Compare_Data"]

    # 原始 period_ends 是 "2024-03-31" 這種帶連字號格式，寫進表裡要轉成
    # 不帶分隔符的 "YYYYMMDD"，跟 Snapshot 輸入格要求的格式一致
    row3 = [c.value for c in ws[3]]
    assert "20240331" in row3
    for cell in ws[3]:
        if cell.value:
            assert not str(cell.value).startswith("=")


def test_compare_data_sheet_lists_company_rows_with_values():
    wb = Workbook()
    result = _sample_result()
    write_compare_data_sheet(wb, result, ["Revenue"])
    ws = wb["Compare_Data"]

    company_col_a = [c.value for c in ws["A"]]
    assert "NVDA" in company_col_a
    assert "AMD" in company_col_a


def test_compare_data_sheet_returns_block_ranges():
    wb = Workbook()
    result = _sample_result()
    ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    assert "Revenue" in ranges
    start, end = ranges["Revenue"]
    assert end > start


def test_compare_data_sheet_stacks_multiple_metric_blocks():
    wb = Workbook()
    result = _sample_result()
    result.metrics["Gross Margin (%)"] = {"NVDA": {"FY2024Q1": 50.0}}
    ranges = write_compare_data_sheet(wb, result, ["Revenue", "Gross Margin (%)"])
    rev_start, rev_end = ranges["Revenue"]
    gm_start, gm_end = ranges["Gross Margin (%)"]
    assert gm_start > rev_end


# ── Snapshot / Snapshot_Manual ──────────────────────────────────────────

def test_snapshot_sheet_has_yellow_input_cell_and_formulas():
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_snapshot_sheets(wb, result, ["Revenue"], block_ranges, default_date="20240331")

    ws = wb["Snapshot"]
    assert ws["B1"].value == "20240331"
    assert ws["B1"].fill.fgColor.rgb in ("00FFFF00", "FFFFFF00")

    body = [[c.value for c in row] for row in ws.iter_rows(min_row=3)]
    formula_cells = [v for row in body for v in row if isinstance(v, str) and v.startswith("=")]
    assert formula_cells, "Snapshot 應該用公式，不是寫死的值"
    assert any("INDEX" in f and "MATCH" in f for f in formula_cells)


def test_snapshot_input_format_matches_period_end_row_format():
    """Snapshot 黃底輸入格要求打 YYYYMMDD，跟 Compare_Data 的期末結算日列
    （已從 "2024-03-31" 轉成 "20240331"）格式一致，MATCH 才對得起來。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_snapshot_sheets(wb, result, ["Revenue"], block_ranges, default_date="20240331")

    data_ws = wb["Compare_Data"]
    period_end_row = [c.value for c in data_ws[3]]
    assert wb["Snapshot"]["B1"].value in period_end_row


def test_snapshot_manual_sheet_is_blank_with_same_headers():
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_snapshot_sheets(wb, result, ["Revenue"], block_ranges, default_date="20240331")

    ws = wb["Snapshot_Manual"]
    header_row = [c.value for c in ws[1]]
    assert "Revenue" in header_row
    company_col = [c.value for c in ws["A"]]
    assert "NVDA" in company_col
    for row in ws.iter_rows(min_row=2, min_col=2):
        for cell in row:
            assert cell.value is None


# ── Chart_<指標> ─────────────────────────────────────────────────────────

def test_chart_sheet_created_per_metric_with_line_chart():
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    assert "Chart_Revenue" in wb.sheetnames
    ws = wb["Chart_Revenue"]
    assert len(ws._charts) == 1


def test_chart_sheet_name_truncates_long_metric_names():
    wb = Workbook()
    result = _sample_result()
    long_name = "A Very Long Metric Name That Exceeds Excel Sheet Name Limit (%)"
    result.metrics[long_name] = {"NVDA": {"FY2024Q1": 1.0}}
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue", long_name])
    write_chart_sheets(wb, ["Revenue", long_name], block_ranges)

    assert all(len(name) <= 31 for name in wb.sheetnames)


# ── write_comparison_workbook ────────────────────────────────────────────

def test_write_comparison_workbook_produces_all_expected_sheets():
    result = _sample_result()
    with tempfile.TemporaryDirectory() as tmp:
        out_path = Path(tmp) / "compare_test.xlsx"
        write_comparison_workbook(result, ["Revenue"], out_path, snapshot_date="20240331")

        assert out_path.exists()
        wb = load_workbook(out_path)
        assert wb.sheetnames == ["Compare_Data", "Snapshot", "Snapshot_Manual", "Chart_Revenue"]
