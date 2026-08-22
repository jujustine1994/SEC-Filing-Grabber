"""comparison_writer.py — 把 comparison.py 的資料結構寫成跨公司比較 Excel。

Sheet 結構（見 docs/superpowers/specs/2026-08-20-cross-company-comparison-design.md）：
  Compare_Data    — 唯一一張原始資料表，每個指標一個區塊往下疊
  Snapshot        — 活的，公式驅動的單一時間點快照
  Snapshot_Manual — 空白，供人工貼值凍結存檔
  Chart_<指標>     — 每個指標各一張，只放圖表
"""

from __future__ import annotations

import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.chart.layout import Layout
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from comparison import ComparisonResult
from excel_formatter import FMT_FINANCIAL, unit_format_for
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

    # 所有公司、所有指標都沒值的期間整欄拿掉。合成 Q4 時年報沒有期末日的那幾
    # 欄（`comparison._aligned_labels()` 算不出日曆季，退回 `FY2009Q4` 這種財季
    # 標籤）值全是空的，排序時又會被排到日曆季後面，圖表 X 軸就多出兩格最新
    # 一期之後的空白，看起來像資料抓錯。**判斷要跨所有指標一起做**：每個區塊
    # 各自篩會讓區塊之間欄數不同，而 write_chart_sheets() 用的是全表
    # `max_column`，窄的區塊會被讀到別人的欄位。
    empty_periods = {
        label
        for metric_data in result.metrics.values()
        for company_data in metric_data.values()
        for label in company_data
    } - {
        label
        for metric_data in result.metrics.values()
        for company_data in metric_data.values()
        for label, value in company_data.items()
        if value is not None
    }

    block_ranges: dict[str, tuple[int, int]] = {}
    row = 1
    for metric_name in metric_names:
        metric_data = result.metrics.get(metric_name, {})
        fmt, divisor = unit_format_for(metric_name)

        # 收集這個指標出現過的所有期間標籤，依標籤字串排序（日曆季 `2025Q2`
        # 與財季 `FY2025Q2` 都天然可字串排序）
        periods: list[str] = sorted({
            label for company_data in metric_data.values() for label in company_data
        } - empty_periods)

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
        # 期間鍵是日曆季（跨公司對齊，見 comparison._aligned_labels()），同一欄
        # 各公司的實際期末日不會一樣——NVDA 那一季結束在 7/27，AMD 是 6/28。
        # 這一格只放得下一個日期，取**最晚**的：Snapshot 拿它做「不晚於 B1」
        # 的判斷，取早的那個會讓 B1 設在 7/1 就顯示 NVDA 還沒結算完的數字。
        for col, period in enumerate(periods, start=2):
            end_date = max(
                (result.period_ends.get(company, {}).get(period, "")
                 for company in all_companies),
                default="",
            )
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

    Snapshot 對 Compare_Data 每個指標區塊的「期末結算日」列（每個區塊的資料
    起始列往上一列，見 write_compare_data_sheet 的排版）取值，改 B1 的日期
    Excel 會自動重算——這是刻意選擇用真公式，不是寫死算好的值，因為這份只給
    人在 Excel 裡看，沒有下游腳本要讀它（讀 Snapshot_Manual）。

    B1 是**真正的 Excel 日期**（不是文字），使用者可以打任何日期（如
    2024/7/15），不需要剛好對到某一期的期末結算日——公式抓的是「不晚於這天
    的最近一期」，符合分析師「這個時間點看得到的最新數字是什麼」的直覺
    （不能用未來才公布的數字回填過去的時間點）。2026-08-21 CTH 回報：原本
    B1 是純文字要求剛好打中 `YYYYMMDD`，使用者打數字會被 Excel 自動轉成
    數值型別，跟 Compare_Data 期末結算日列的文字型別對不上，MATCH 直接
    抓不到值；順便換掉這個 exact-match 限制。
    """
    all_companies = sorted({
        company
        for metric_data in result.metrics.values()
        for company in metric_data
    })
    data_ws = wb["Compare_Data"]

    snap = wb.create_sheet("Snapshot")
    snap["A1"] = t("compare.xls.timepoint")
    if default_date:
        try:
            snap["B1"] = datetime.datetime.strptime(default_date, "%Y%m%d").date()
        except ValueError:
            snap["B1"] = default_date  # 給不了合法日期就照原樣寫，不讓整個產檔失敗
    snap["B1"].number_format = "yyyy/mm/dd"
    snap["B1"].fill = _YELLOW_FILL

    # CTH 回報過黃格不知道要填什麼——A1 標籤只寫「時間點」看不出格式，旁邊
    # 補一句話講清楚格式，再把 Compare_Data 裡實際存在的期末結算日列出來，
    # 使用者不用自己去翻 Compare_Data 猜有哪些日期可以填。
    available_dates = sorted({
        end.replace("-", "")
        for company_map in result.period_ends.values()
        for end in company_map.values()
        if end
    })
    snap.cell(row=1, column=3,
              value=t("compare.xls.snapshot_format_hint"))
    snap.cell(row=1, column=4,
              value=t("compare.xls.snapshot_available_dates", dates="、".join(available_dates)))

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
            # 「不晚於 B1 的最近一期」：期末結算日是固定 8 碼、左補零的
            # YYYYMMDD 文字，字串比較順序跟數值順序一致，可以直接用 <=。
            # SUMPRODUCT(MAX(...)) 取代 MATCH(...,0) 的精確比對——對每一格算
            # 「這家公司這期有資料 且 不晚於目標日」則給它的欄位序號、否則
            # 給 0，取最大值就是最近（最晚但不超過目標日）那一格的位置。
            #
            # 注意：「有沒有資料」要看 data_range（這家公司自己那一列），
            # 不能看 end_date_range（期末結算日列）——後者是**所有公司期間
            # 標籤的聯集**，A 公司有 Q2、B 公司沒有時，聯集的日期列在 Q2
            # 那格仍然有值，若拿它判斷空白會誤判「B 這期有資料」，實際
            # INDEX 到 B 真正空白的儲存格，Excel 把空格當 0 處理，算出錯誤
            # 的 0（這裡是實測 Excel COM 抓出來的真的會發生，不是理論推測）。
            #
            # 目標日期比任何一期都早時 MAX 會是 0——不能只靠 IFERROR 接住，
            # 因為 INDEX(range,0) 在 Excel 不會報錯，而是回傳整個範圍的隱含
            # 第一格（同樣是實測發現，會誤顯示成 Q1 的值），要用 IF 明確擋
            # 掉 0 的情況才會顯示空白。
            offset_expr = (
                f'SUMPRODUCT(MAX(({data_range}<>"")'
                f'*({end_date_range}<=TEXT($B$1,"yyyymmdd"))'
                f'*(COLUMN({end_date_range})-COLUMN(INDEX({end_date_range},1,1))+1)))'
            )
            formula = f'=IFERROR(IF({offset_expr}=0,"",INDEX({data_range},{offset_expr})),"")'
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


def _pin_title_layout(title) -> None:
    """明講標題「用自動版面、不要疊在別的東西上面」。

    2026-08-22 CTH 截圖回報 Y 軸標題壓在「50,000.0」刻度數字上、X 軸標題
    「期間」掉進日期標籤那一排裡面。跟前一輪圖例那個 bug 同一類：openpyxl 的
    `Title` 沒有 `overlay` / `layout` 屬性時**完全不寫這兩個元素**，Excel 拿到
    「沒寫」的標題會直接畫在既有內容上面，而不是另外撥一條專屬空間給它。
    原生 Excel 輸出的每個標題一定帶 `<c:layout/>`（空元素＝明講用自動版面）
    加 `<c:overlay val="0"/>`。圖表標題與兩個軸標題都要補。
    """
    if title is None:
        return
    title.overlay = False
    title.layout = Layout()


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
        end_date_row = data_start - 1  # 期末結算日列，緊接在資料列上方（絕對日期，真實時間可排序）
        last_col = data_ws.max_column

        chart = LineChart()
        chart.title = metric_name
        chart.style = 2
        chart.x_axis.title = t("gui.compare.period")
        # 2026-08-22 CTH 截圖回報「中間斷線＋圖例被吃＋沒有單位」，實測抓出
        # 真正根因：openpyxl 產生的 <c:catAx>／<c:valAx> 完全沒有寫
        # <c:delete> 元素。拿 Excel COM 原生建立的圖表（Shapes/ChartObjects
        # 直接建，不經 openpyxl）比對 XML，發現原生輸出**一定**帶
        # `<c:delete val="0"/>`。OOXML 規格上 delete 沒寫預設就是 false，
        # 但 Excel 實際渲染時對「沒寫」跟「明講 false」待遇不同——沒寫會
        # 保守地不畫刻度標籤，連帶把圖例／座標軸空間計算搞壞、擠壓變形。
        # 這是用 PowerShell 呼叫 Excel COM 實測、比對原生輸出 XML 才抓到的，
        # 純看 openpyxl 文件或程式碼猜不到。兩軸都要明講 delete=False。
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        # openpyxl 兩個軸預設都是 axPos="l"（見 openpyxl.chart.axis._BaseAxis
        # 的預設值）——對 Y 軸（valAx）剛好是對的，對 X 軸（catAx）是錯的，
        # 該在底部卻被標成左側。原生 Excel 輸出兩軸位置都明講，這裡跟進。
        chart.x_axis.axPos = "b"
        chart.y_axis.axPos = "l"
        # 同理明講兩軸的刻度標籤要顯示、跟軸線交叉方式，都是原生 Excel
        # 輸出必有的欄位，openpyxl 預設不會寫，一起補齊不留模糊地帶。
        chart.x_axis.tickLblPos = "nextTo"
        chart.y_axis.tickLblPos = "nextTo"
        chart.x_axis.crosses = "autoZero"
        chart.y_axis.crosses = "autoZero"
        # 不同公司財年結束月不同，財季標籤（FY2024Q3）字串排序無法反映真實時間，
        # 缺值期間 Excel 預設會直接連到下一個有值的點造成誤導折線，兩者一起處理：
        # X 軸類別改用期末結算日（絕對日期）取代財季標籤，缺值處顯示為斷點不連線
        chart.display_blanks = "gap"

        # F3（2026-08-20 CTH 截圖回報 5 項，2026-08-21 確認最終方案）：
        # 尺寸拉一倍（openpyxl 預設 15cm×7.5cm 偏小）、圖例移到下方橫排
        # （右側直排在公司數一多時會被擠進繪圖區，下方橫排可隨公司數換行）
        chart.width = 30
        chart.height = 15
        chart.legend.position = "b"
        # 同一類「openpyxl 沒寫、Excel 就當不確定狀態」的坑：legend 沒有
        # overlay 屬性時，圖例會跟 X 軸標題/刻度標籤擠在同一條窄帶、疊在
        # 一起看不清楚（實測畫面：圖例文字直接蓋在 X 軸日期上）。原生 Excel
        # 輸出一定帶 `overlay="0"`（不要疊加，另外保留專屬空間），這裡明講。
        chart.legend.overlay = False

        # Y 軸數字格式跟 Compare_Data 儲存格同一套規則（不要另外定義一套會
        # 漂移的格式）；金額類指標軸標題帶單位，百分比類指標格式本身已經
        # 看得出是 %，標題維持指標名稱就好，不重複講
        fmt, _ = unit_format_for(metric_name)
        chart.y_axis.numFmt = fmt
        chart.y_axis.title = f"{metric_name} ($mm)" if fmt == FMT_FINANCIAL else metric_name

        # 三個標題都要明講 layout/overlay，理由見 _pin_title_layout()。
        # 一定要放在三個 title 都設完之後——openpyxl 每次指派 title 都會
        # 重新造一個 Title 物件，先設屬性會被後面的指派蓋掉。
        _pin_title_layout(chart.title)
        _pin_title_layout(chart.x_axis.title)
        _pin_title_layout(chart.y_axis.title)

        data_ref = Reference(
            data_ws, min_col=1, max_col=last_col, min_row=data_start, max_row=data_end
        )
        chart.add_data(data_ref, titles_from_data=True, from_rows=True)

        categories_ref = Reference(
            data_ws, min_col=2, max_col=last_col, min_row=end_date_row, max_row=end_date_row
        )
        chart.set_categories(categories_ref)
        # openpyxl 的 set_categories() 不管儲存格實際內容是什麼，永遠寫成
        # <cat><numRef>（數值參照）。期末結算日是文字（"20240331"，寫檔時
        # 刻意存成文字給 Snapshot 的 MATCH／SUMPRODUCT 用），Excel 拿到指向
        # 文字儲存格的數值參照解析不出來，類別軸整個讀不到值、連帶把圖例／
        # 座標軸擠壓變形（CTH 截圖回報「中間斷線＋圖例被吃」的真正原因）。
        # 手動把每個 series 的 cat 換成 strRef，內容不變，只是型別對上。
        str_categories_ref = StrRef(f=str(categories_ref))
        for series in chart.series:
            series.cat = AxDataSource(strRef=str_categories_ref)

        # 接上 D0-1 Q4 合成後跨公司比較的時間跨度可以拉到 60-70 欄，全部日期
        # 標籤硬擠在 X 軸上會疊字看不清楚（這是修復帶出的新問題，原本 F3 5
        # 項清單沒有涵蓋）。目標大約 15 個可視標籤，跳著顯示。
        n_periods = last_col - 1  # 扣掉 A 欄（公司名），B 欄開始才是期間
        tick_skip = max(1, n_periods // 15)
        chart.x_axis.tickLblSkip = tick_skip
        chart.x_axis.tickMarkSkip = tick_skip

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
