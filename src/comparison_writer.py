"""comparison_writer.py — 把 comparison.py 的資料結構寫成跨公司比較 Excel。

Sheet 結構（見 docs/superpowers/specs/2026-08-20-cross-company-comparison-design.md）：
  Compare_Data    — 唯一一張原始資料表，每個指標一個區塊往下疊
  Snapshot        — 活的，公式驅動的單一時間點快照
  Snapshot_Manual — 空白，供人工貼值凍結存檔
  Chart_<指標>     — 每個指標各一張，只放圖表
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from comparison import ComparisonResult
from excel_formatter import unit_format_for
from i18n import t

_HEADER_FONT = Font(bold=True)
_BLOCK_GAP = 1  # 區塊之間空幾列
_YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


def write_compare_data_sheet(
    wb: Workbook, result: ComparisonResult, metric_names: list[str]
) -> dict[str, tuple[int, int]]:
    """寫 Compare_Data。回傳 {指標名: (資料列起, 資料列迄)}（不含標題/期末結算日列），
    給 Snapshot 的 MATCH 公式與 Chart 的資料來源 range 用。"""
    ws = wb.active
    ws.title = "Compare_Data"

    all_companies = sorted({
        company
        for metric_data in result.metrics.values()
        for company in metric_data
    })

    block_ranges: dict[str, tuple[int, int]] = {}
    row = 1
    for metric_name in metric_names:
        metric_data = result.metrics.get(metric_name, {})
        fmt, divisor = unit_format_for(metric_name)

        # 收集這個指標出現過的所有期間標籤，依標籤字串排序（FYyyyyQq 天然可字串排序）
        periods: list[str] = sorted({
            label for company_data in metric_data.values() for label in company_data
        })

        # 標題列
        title_cell = ws.cell(row=row, column=1, value=metric_name)
        title_cell.font = _HEADER_FONT
        header_row = row + 1
        ws.cell(row=header_row, column=1, value=t("compare.xls.company"))
        for col, period in enumerate(periods, start=2):
            ws.cell(row=header_row, column=col, value=period)

        # 期末結算日列（靜態文字，供 Snapshot 用）。fetcher_gaap 給的原始格式是
        # "YYYY-MM-DD"，這裡去掉分隔符轉成 "YYYYMMDD"——跟 Snapshot 黃底輸入格
        # 要求使用者打的格式一致，MATCH 才對得起來，不用在公式裡另外做轉換。
        end_date_row = header_row + 1
        ws.cell(row=end_date_row, column=1, value=t("compare.xls.period_end"))
        for col, period in enumerate(periods, start=2):
            end_date = ""
            for company in all_companies:
                end_date = result.period_ends.get(company, {}).get(period, "")
                if end_date:
                    break
            ws.cell(row=end_date_row, column=col, value=end_date.replace("-", ""))

        # 公司資料列
        data_start = end_date_row + 1
        for offset, company in enumerate(all_companies):
            r = data_start + offset
            ws.cell(row=r, column=1, value=company)
            company_data = metric_data.get(company, {})
            for col, period in enumerate(periods, start=2):
                value = company_data.get(period)
                cell = ws.cell(row=r, column=col, value=value)
                if isinstance(value, (int, float)):
                    cell.value = value / divisor
                    cell.number_format = fmt
        data_end = data_start + len(all_companies) - 1

        block_ranges[metric_name] = (data_start, data_end)
        row = data_end + 1 + _BLOCK_GAP

    return block_ranges


def write_snapshot_sheets(
    wb: Workbook,
    result: ComparisonResult,
    metric_names: list[str],
    block_ranges: dict[str, tuple[int, int]],
    default_date: str = "",
) -> None:
    """寫 Snapshot（活公式）與 Snapshot_Manual（空白，供人工貼值）。

    Snapshot 用 INDEX/MATCH 對 Compare_Data 每個指標區塊的「期末結算日」列
    （每個區塊的資料起始列往上一列，見 write_compare_data_sheet 的排版）取值，
    改 B1 的日期 Excel 會自動重算——這是刻意選擇用真公式，不是寫死算好的值，
    因為這份只給人在 Excel 裡看，沒有下游腳本要讀它（讀 Snapshot_Manual）。
    """
    all_companies = sorted({
        company
        for metric_data in result.metrics.values()
        for company in metric_data
    })
    data_ws = wb["Compare_Data"]

    snap = wb.create_sheet("Snapshot")
    snap["A1"] = t("compare.xls.timepoint")
    snap["B1"] = default_date
    snap["B1"].fill = _YELLOW_FILL

    header_row = 2
    snap.cell(row=header_row, column=1, value=t("compare.xls.company"))
    for col, metric_name in enumerate(metric_names, start=2):
        snap.cell(row=header_row, column=col, value=metric_name)

    for r_offset, company in enumerate(all_companies):
        r = header_row + 1 + r_offset
        snap.cell(row=r, column=1, value=company)
        for col, metric_name in enumerate(metric_names, start=2):
            if metric_name not in block_ranges:
                continue
            data_start, data_end = block_ranges[metric_name]
            end_date_row = data_start - 1   # 期末結算日列緊接在資料列上方

            company_row = None
            for rr in range(data_start, data_end + 1):
                if data_ws.cell(row=rr, column=1).value == company:
                    company_row = rr
                    break
            if company_row is None:
                continue

            last_col = data_ws.max_column
            end_date_range = (
                f"Compare_Data!$B${end_date_row}:${get_column_letter(last_col)}${end_date_row}"
            )
            data_range = (
                f"Compare_Data!$B${company_row}:${get_column_letter(last_col)}${company_row}"
            )
            formula = f'=INDEX({data_range},MATCH($B$1,{end_date_range},0))'
            snap.cell(row=r, column=col, value=formula)

    # Snapshot_Manual：同樣的表頭，資料格留空供人工貼值
    manual = wb.create_sheet("Snapshot_Manual")
    manual.cell(row=1, column=1, value=t("compare.xls.company"))
    for col, metric_name in enumerate(metric_names, start=2):
        manual.cell(row=1, column=col, value=metric_name)
    for r_offset, company in enumerate(all_companies):
        manual.cell(row=2 + r_offset, column=1, value=company)


def _chart_sheet_name(metric_name: str) -> str:
    """Chart_<指標> 但要塞進 Excel 的 31 字元 sheet 名稱上限。"""
    prefix = "Chart_"
    max_metric_len = 31 - len(prefix)
    safe_name = "".join(ch for ch in metric_name if ch not in '[]:*?/\\')
    return prefix + safe_name[:max_metric_len]


def write_chart_sheets(
    wb: Workbook, metric_names: list[str], block_ranges: dict[str, tuple[int, int]]
) -> None:
    """每個指標各一張 sheet，只放一張折線圖（歷史趨勢，一條線一家公司）。
    使用者要看長條圖版本，在 Excel 裡對圖表右鍵「變更圖表類型」自己切，
    這裡不用同一指標產兩份圖表物件。"""
    data_ws = wb["Compare_Data"]

    for metric_name in metric_names:
        if metric_name not in block_ranges:
            continue
        data_start, data_end = block_ranges[metric_name]
        header_row = data_start - 2
        last_col = data_ws.max_column

        chart = LineChart()
        chart.title = metric_name
        chart.style = 2
        chart.y_axis.title = metric_name
        chart.x_axis.title = t("gui.compare.period")

        data_ref = Reference(
            data_ws, min_col=1, max_col=last_col, min_row=data_start, max_row=data_end
        )
        chart.add_data(data_ref, titles_from_data=True, from_rows=True)

        categories_ref = Reference(
            data_ws, min_col=2, max_col=last_col, min_row=header_row, max_row=header_row
        )
        chart.set_categories(categories_ref)

        sheet_name = _chart_sheet_name(metric_name)
        chart_ws = wb.create_sheet(sheet_name)
        chart_ws.add_chart(chart, "B2")


def write_comparison_workbook(
    result: ComparisonResult,
    metric_names: list[str],
    output_path: Path,
    snapshot_date: str = "",
) -> None:
    """組出完整跨公司比較 Excel 並存檔。"""
    wb = Workbook()
    block_ranges = write_compare_data_sheet(wb, result, metric_names)
    write_snapshot_sheets(wb, result, metric_names, block_ranges, default_date=snapshot_date)
    write_chart_sheets(wb, metric_names, block_ranges)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
