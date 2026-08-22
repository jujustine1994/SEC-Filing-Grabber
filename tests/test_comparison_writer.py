"""Tests for comparison_writer.py — 跨公司比較 Excel 輸出。"""
import tempfile
from datetime import date, datetime
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
    assert ws["B1"].value == date(2024, 3, 31)
    assert ws["B1"].fill.fgColor.rgb in ("00FFFF00", "FFFFFF00")

    body = [[c.value for c in row] for row in ws.iter_rows(min_row=3)]
    formula_cells = [v for row in body for v in row if isinstance(v, str) and v.startswith("=")]
    assert formula_cells, "Snapshot 應該用公式，不是寫死的值"
    assert any("INDEX" in f and "SUMPRODUCT" in f for f in formula_cells)


def test_snapshot_input_cell_is_a_real_date_not_text():
    """2026-08-21 CTH 回報：B1 原本是純文字要求剛好打中 YYYYMMDD，使用者打
    數字會被 Excel 自動轉成數值型別，跟 Compare_Data 期末結算日列（文字）
    型別對不上，MATCH 抓不到值。B1 改成真正的日期型別，使用者可以打一般
    日期格式，不用湊 8 碼數字，也不會再有型別不匹配的問題。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_snapshot_sheets(wb, result, ["Revenue"], block_ranges, default_date="20240331")

    b1 = wb["Snapshot"]["B1"]
    assert isinstance(b1.value, (date, datetime))
    assert b1.number_format != "General"


def test_snapshot_formula_finds_nearest_period_not_later_than_input_date():
    """2026-08-21 CTH 要求：輸入日期不用剛好對到期末結算日，抓「不晚於這天
    的最近一期」——分析師查某個時間點看得到的數字，不能用未來才公布的財報
    回填過去的時間點。公式要用 <= 比較（不是精確比對），且要排除空白欄位
    （該公司那期沒資料）。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_snapshot_sheets(wb, result, ["Revenue"], block_ranges, default_date="")

    ws = wb["Snapshot"]
    formula = next(
        c.value for row in ws.iter_rows(min_row=3) for c in row
        if isinstance(c.value, str) and c.value.startswith("=")
    )
    assert "<=" in formula
    assert "MATCH" not in formula  # 不是精確比對
    assert "IFERROR" in formula    # 找不到任何一期時要顯示空白，不是報錯


def test_snapshot_sheet_hints_that_any_date_is_accepted():
    """CTH 回報 Snapshot 的黃格不知道要填什麼——提示文字要講清楚可以打一般
    日期，不用剛好對到某一期。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_snapshot_sheets(wb, result, ["Revenue"], block_ranges, default_date="")

    ws = wb["Snapshot"]
    all_text = " ".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
    assert "2025/12/31" in all_text or "2025" in all_text


def test_snapshot_sheet_lists_available_dates_for_reference():
    """光講格式還不夠，使用者不知道「有哪些日期可以填」——列出 Compare_Data
    裡實際存在的期末結算日，照現有資料範圍給選項，不用自己去翻 Compare_Data。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_snapshot_sheets(wb, result, ["Revenue"], block_ranges, default_date="")

    ws = wb["Snapshot"]
    all_text = " ".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
    assert "20240331" in all_text
    assert "20240630" in all_text


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


def test_chart_sheet_uses_period_end_date_row_as_categories():
    """X 軸類別要用 Compare_Data 的期末結算日列（絕對日期，如 20240331），
    不是財季標籤列（如 FY2024Q1）——不同公司財年結束月不同，財季標籤字串
    排序無法反映真實時間順序，會讓 Excel 折線圖亂連線。期末結算日列緊接在
    資料列上方（data_start - 1），財季標籤列則在再上一列（data_start - 2）。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    data_start, _ = block_ranges["Revenue"]
    ws = wb["Chart_Revenue"]
    chart = ws._charts[0]
    cat_ref = chart.series[0].cat.numRef.f if chart.series[0].cat.numRef else chart.series[0].cat.strRef.f
    assert f"${data_start - 1}:" in cat_ref or f"${data_start - 1}$" in cat_ref
    assert f"${data_start - 2}" not in cat_ref


def test_chart_sheet_categories_use_str_ref_not_num_ref():
    """2026-08-22 CTH 截圖回報「中間斷線＋圖例被吃」：根因是 openpyxl 的
    `chart.set_categories()` 不管儲存格實際內容永遠寫成 numRef（數值參照），
    但期末結算日是文字（"20240331"），Excel 拿數值參照指向文字儲存格解析
    不出來，類別軸整個讀不到值，連帶把圖例/座標軸擠壓變形。每個 series 的
    `cat` 必須是 strRef，不能是 numRef。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    chart = wb["Chart_Revenue"]._charts[0]
    for series in chart.series:
        assert series.cat.numRef is None, "類別軸不可以是 numRef——期末結算日是文字"
        assert series.cat.strRef is not None
        assert series.cat.strRef.f  # 有指到範圍


def test_chart_sheet_axis_positions_and_tick_labels_are_explicit():
    """2026-08-22 CTH 截圖回報：X 軸完全沒有日期標籤、圖例被擠壓成一行。
    根因是 openpyxl 兩個軸預設都是 axPos="l"（見 openpyxl.chart.axis.
    _BaseAxis），對 Y 軸剛好對、對 X 軸（該在底部）是錯的，Excel 拿到矛盾
    設定會整排 X 軸標籤不畫。兩軸的位置與刻度標籤顯示都要明講，不能依賴
    openpyxl 的預設值。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    chart = wb["Chart_Revenue"]._charts[0]
    assert chart.x_axis.axPos == "b"
    assert chart.y_axis.axPos == "l"
    assert chart.x_axis.tickLblPos == "nextTo"
    assert chart.y_axis.tickLblPos == "nextTo"


def test_chart_sheet_axes_explicitly_not_deleted():
    """2026-08-22 CTH 截圖回報「中間斷線＋圖例被吃＋沒有單位」：實測用 Excel
    COM 比對原生 Excel 建立的圖表 XML 才抓到真正根因——openpyxl 完全不寫
    <c:delete> 元素，Excel 拿到「沒寫」的軸會保守地不畫刻度標籤（跟明講
    delete=False 待遇不同，這是實測結果不是規格文件寫的）。兩軸都要明講
    delete=False，不能依賴 openpyxl 的預設（不寫）。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    chart = wb["Chart_Revenue"]._charts[0]
    assert chart.x_axis.delete is False
    assert chart.y_axis.delete is False


def test_chart_sheet_sets_display_blanks_as_gap():
    """缺值期間不能被 Excel 直接連到下一個有值的點，否則折線會誤導使用者
    以為中間有連續資料。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    chart = wb["Chart_Revenue"]._charts[0]
    assert chart.display_blanks == "gap"


# ── F3 圖表版面調校（2026-08-21，CTH 截圖回報 5 項，CTH 已確認最終方案）──

def test_chart_sheet_is_double_default_size():
    """CTH 要求長寬各拉長一倍，openpyxl 預設 15cm×7.5cm 偏小，改成 30cm×15cm。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    chart = wb["Chart_Revenue"]._charts[0]
    assert chart.width == 30
    assert chart.height == 15


def test_chart_sheet_legend_does_not_overlay():
    """2026-08-22 CTH 截圖回報圖例文字疊在 X 軸日期上看不清楚：實測用 Excel
    COM 比對原生 Excel 建立的圖表 XML 抓到根因——openpyxl 的 `Legend` 沒有
    `overlay` 屬性時完全不寫該欄位，Excel 拿到「沒寫」會讓圖例跟 X 軸標題/
    刻度標籤擠在同一條窄帶、直接疊在一起。原生 Excel 輸出一定帶
    `overlay="0"`（不要疊加，另外保留專屬空間），這裡要明講，不能依賴
    openpyxl 的預設（不寫）。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    chart = wb["Chart_Revenue"]._charts[0]
    assert chart.legend.overlay is False


def test_chart_sheet_legend_at_bottom():
    """圖例放下方橫排，不放右邊——CTH 確認：公司數量不固定，右側直排公司
    一多會佔掉整個右半邊圖表，下方橫排可以隨公司數自動換行撐住。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    chart = wb["Chart_Revenue"]._charts[0]
    assert chart.legend.position == "b"


def test_chart_sheet_y_axis_uses_same_number_format_as_cells():
    """Y 軸數字格式跟 Compare_Data 儲存格同一套規則（excel_formatter.unit_format_for），
    不要另外定義一套會漂移的格式。金額類指標軸標題帶 ($mm) 單位；百分比類
    指標維持指標名稱本身（格式已經看得出是 %，不用重複講）。"""
    wb = Workbook()
    result = _sample_result()
    result.metrics["Gross Margin (%)"] = {"NVDA": {"FY2024Q1": 50.0}}
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue", "Gross Margin (%)"])
    write_chart_sheets(wb, ["Revenue", "Gross Margin (%)"], block_ranges)

    from excel_formatter import FMT_FINANCIAL, FMT_PERCENT

    rev_chart = wb["Chart_Revenue"]._charts[0]
    assert rev_chart.y_axis.numFmt.formatCode == FMT_FINANCIAL
    assert rev_chart.y_axis.title.text.rich.p[0].r[0].t == "Revenue ($mm)"

    margin_chart = wb["Chart_Gross Margin (%)"]._charts[0]
    assert margin_chart.y_axis.numFmt.formatCode == FMT_PERCENT
    assert margin_chart.y_axis.title.text.rich.p[0].r[0].t == "Gross Margin (%)"


def test_chart_sheet_x_axis_skips_labels_when_many_periods():
    """新發現的問題（不在原本 F3 5 項清單裡）：接上 D0-1 Q4 合成後跨公司比較
    的時間跨度可以拉到 60-70 欄，全部日期標籤硬擠在 X 軸上會疊字看不清楚。
    CTH 確認：目標大約 15 個可視標籤，用 tickLblSkip 跳著顯示。"""
    wb = Workbook()
    periods = [f"FY{2009 + i // 4}Q{i % 4 + 1}" for i in range(60)]
    result = _sample_result()
    result.metrics["Revenue"]["NVDA"] = {p: float(i) for i, p in enumerate(periods)}
    result.period_ends["NVDA"] = {
        p: f"{2009 + i // 4}-{i % 4 + 1:02d}-01" for i, p in enumerate(periods)
    }
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    chart = wb["Chart_Revenue"]._charts[0]
    n_periods = len(periods)
    expected_skip = max(1, n_periods // 15)
    assert chart.x_axis.tickLblSkip == expected_skip
    assert chart.x_axis.tickMarkSkip == expected_skip


def test_chart_sheet_x_axis_no_skip_when_few_periods():
    """期間數少的話（原本的 2 期樣本資料）不用跳著顯示，skip 應該是 1。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    chart = wb["Chart_Revenue"]._charts[0]
    assert chart.x_axis.tickLblSkip == 1


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


# ── 同一日曆季、各公司期末日不同（2026-08-22）───────────────────────────────

def test_period_end_row_takes_the_latest_date_in_the_calendar_quarter():
    """跨公司改用日曆季對齊後，同一欄各公司的期末日不再相同。

    NVDA 那一季結束在 2025-07-27，AMD 是 2025-06-28。期末結算日列只有一格，
    要取**最晚**的那一個：Snapshot 用它做「不晚於 B1」的判斷，取早的那個會
    讓使用者把 B1 設在 7/1 就看到 NVDA 還沒結算完的那一季數字。
    """
    result = ComparisonResult(
        metrics={"Revenue": {"NVDA": {"2025Q2": 46743.0}, "AMD": {"2025Q2": 7685.0}}},
        period_ends={"NVDA": {"2025Q2": "2025-07-27"}, "AMD": {"2025Q2": "2025-06-28"}},
        failures=[],
    )
    wb = Workbook()
    write_compare_data_sheet(wb, result, ["Revenue"])

    assert wb["Compare_Data"]["B3"].value == "20250727"


def test_compare_data_sheet_drops_periods_with_no_data_at_all():
    """所有公司、所有指標都沒值的期間不要出現在表上。

    合成 Q4 時年報沒有期末日的那幾欄（label 退回 `FY2009Q4` 這種財季標籤）
    值全是空的，改用日曆季當欄位鍵之後它們會被排到最右邊——圖表 X 軸多出
    兩格 2026Q2 之後的空白，看起來像資料抓錯。
    """
    result = ComparisonResult(
        metrics={"Revenue": {
            "NVDA": {"2025Q1": 44062.0, "2025Q2": 46743.0, "FY2009Q4": None},
            "AMD": {"2025Q1": 7438.0, "2025Q2": 7685.0, "FY2009Q4": None},
        }},
        period_ends={
            "NVDA": {"2025Q1": "2025-04-27", "2025Q2": "2025-07-27", "FY2009Q4": ""},
            "AMD": {"2025Q1": "2025-03-29", "2025Q2": "2025-06-28", "FY2009Q4": ""},
        },
        failures=[],
    )
    wb = Workbook()
    write_compare_data_sheet(wb, result, ["Revenue"])

    header_row = [c.value for c in wb["Compare_Data"][2]]
    assert "FY2009Q4" not in header_row
    assert header_row[1:] == ["2025Q1", "2025Q2"]


def test_compare_data_sheet_keeps_period_when_only_one_company_has_data():
    """只有一家有值的期間要留著——那是真資料，不是空欄。"""
    result = ComparisonResult(
        metrics={"Revenue": {
            "NVDA": {"2025Q1": 44062.0, "2025Q2": 46743.0},
            "AMD": {"2025Q1": 7438.0, "2025Q2": None},
        }},
        period_ends={
            "NVDA": {"2025Q1": "2025-04-27", "2025Q2": "2025-07-27"},
            "AMD": {"2025Q1": "2025-03-29", "2025Q2": ""},
        },
        failures=[],
    )
    wb = Workbook()
    write_compare_data_sheet(wb, result, ["Revenue"])

    header_row = [c.value for c in wb["Compare_Data"][2]]
    assert header_row[1:] == ["2025Q1", "2025Q2"]


def test_chart_titles_do_not_overlay_and_use_auto_layout():
    """2026-08-22 CTH 截圖回報 Y 軸標題壓在「50,000.0」刻度上、X 軸標題「期間」
    掉進日期標籤那一排裡面。跟先前圖例那個 bug 同一類：openpyxl 的
    `Title` 沒有 `overlay` / `layout` 屬性時完全不寫這兩個元素，Excel 拿到
    「沒寫」的標題會直接畫在既有內容上面，而不是另外撥一條專屬空間。
    原生 Excel 輸出的每個標題（圖表、兩個軸）一定帶 `<c:layout/>`（明講
    用自動版面）加 `<c:overlay val="0"/>`。三個標題都要明講。"""
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    chart = wb["Chart_Revenue"]._charts[0]
    for title in (chart.title, chart.x_axis.title, chart.y_axis.title):
        assert title.overlay is False
        assert title.layout is not None


def test_chart_title_overlay_is_written_into_the_saved_file():
    """光設屬性不夠——Excel 讀的是存檔後的 XML，要在真的 .xlsx 裡確認。"""
    import zipfile

    result = _sample_result()
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "chart_xml.xlsx"
        write_comparison_workbook(result, ["Revenue"], out)
        with zipfile.ZipFile(out) as z:
            name = next(n for n in z.namelist() if "charts/chart" in n)
            xml = z.read(name).decode("utf-8")

    # 圖表標題 + X 軸標題 + Y 軸標題 + 圖例 = 4 個 overlay；
    # layout 只有三個標題有（圖例的版面由 legendPos 決定）
    assert xml.count('<overlay val="0"/>') == 4
    assert xml.count("<layout/>") == 3
