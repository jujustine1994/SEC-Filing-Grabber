# 交接文件：跨公司比較 `Chart_<指標>` 圖表殘留問題

**專案**：SEC Financial Fetcher（`C:\Users\CTH\Documents\Code\SEC Financial Tools`）
**日期**：2026-08-22
**用途**：這份文件給另一個 AI model review 用，把目前圖表功能的完整脈絡、已修復項目、還沒解決的兩類問題（資料缺漏、美編排版）講清楚，附上實際程式碼、資料、截圖層級的觀察。

---

## 1. 功能背景

「跨公司比較」是這個工具的一個功能：使用者選 2~N 家美股公司（ticker）+ 若干財務指標（如 Revenue、Gross Margin (%)），程式對每家公司抓 SEC EDGAR 的 XBRL 財報資料，重組成一份 Excel，裡面有：

- `Compare_Data`：原始資料表，每個指標一個區塊，區塊內每家公司一列，欄位是期間
- `Snapshot` / `Snapshot_Manual`：單一時間點的橫向切面表
- `Chart_<指標>`：每個指標一張折線圖，X 軸是期間、Y 軸是數值，一條線一家公司

本文件只談 `Chart_<指標>` 這部分的兩個殘留問題。

---

## 2. 資料流與相關程式碼

### 2.1 抓資料：`src/fetcher_gaap.py`

- `fetch_gaap_statements(ticker, identity, max_filings=80, max_annual_filings=20, fetch_quarterly, fetch_annual, ...)` 是核心抓取函式，逐份解析公司的 10-Q（季報）與 10-K（年報）XBRL，回傳 `StatementTable` 清單（依 sheet_name 分：`Data_Financials(Q)`、`Data_Financials(Y)` 等）。
- `max_filings=80` 上限（`fetcher_gaap.py:2066` 附近），註解寫「~20 年」。AMD/INTC 從 2009 年至今約 65~68 季，理論上在上限內，**但沒有針對這次跨公司比較的實際抓取結果逐一核對是否真的每家都抓好抓滿，這是本文件請 review 的重點之一（見 4.3）**。
- **Q4 合成邏輯**（`_synthesize_q4()`，`fetcher_gaap.py` 內）：SEC 的 10-Q 只涵蓋 Q1/Q2/Q3，公司的「日曆 Q4」財報數字要從年報（10-K）反推——流量項（IS/CF）用「年報－Q1－Q2－Q3」相減，存量項（BS）直接取年報值。**前提是 Q1/Q2/Q3 三季原始資料都要有，年報也要有**，任何一段缺，Q4 就合成不出來（見 4.2 的因果鏈）。

### 2.2 跨公司重組：`src/comparison.py`

```python
# comparison.py:107-119
with report_progress(progress_cb):
    tables = fetch_gaap_statements(
        ticker, identity, max_filings=max_filings,
        max_annual_filings=max_annual_filings,
        fetch_quarterly=(frequency == "quarterly"),
        fetch_annual=True,   # 季度模式也強制抓年報，見下方說明
    )
```

- `build_comparison()` 對每個 ticker 呼叫一次 `fetch_gaap_statements()`，把結果依指標名重組成 `{指標: {ticker: {period_label: value}}}`（`ComparisonResult.metrics`）與 `{ticker: {period_label: 期末結算日}}`（`ComparisonResult.period_ends`）。
- **`fetch_annual=True` 是這次 session 修過的一個 bug**：原本 `fetch_annual=(frequency == "annual")`，選季度比較時完全不抓年報，Q4 合成沒有材料，日曆年結束的公司（AMD/INTC）Q4 整欄空白。改成永遠抓年報後，Q4 缺洞大幅減少（但沒有變成 0，見下方 4.2）。
- 單一公司抓取失敗（`except Exception`）只記錄成 `CompanyFetchError`，不中斷其他公司（`comparison.py:120-122`）。

### 2.3 寫圖表：`src/comparison_writer.py::write_chart_sheets()`

```python
# comparison_writer.py:220-312（節錄，完整見檔案）
def write_chart_sheets(wb, metric_names, block_ranges):
    data_ws = wb["Compare_Data"]
    for metric_name in metric_names:
        data_start, data_end = block_ranges[metric_name]
        end_date_row = data_start - 1
        last_col = data_ws.max_column

        chart = LineChart()
        chart.title = metric_name
        chart.x_axis.title = t("gui.compare.period")   # 中文「期間」
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.x_axis.axPos = "b"
        chart.y_axis.axPos = "l"
        chart.x_axis.tickLblPos = "nextTo"
        chart.y_axis.tickLblPos = "nextTo"
        chart.x_axis.crosses = "autoZero"
        chart.y_axis.crosses = "autoZero"
        chart.display_blanks = "gap"

        chart.width = 30
        chart.height = 15
        chart.legend.position = "b"
        chart.legend.overlay = False

        fmt, _ = unit_format_for(metric_name)
        chart.y_axis.numFmt = fmt
        chart.y_axis.title = f"{metric_name} ($mm)" if fmt == FMT_FINANCIAL else metric_name

        data_ref = Reference(data_ws, min_col=1, max_col=last_col,
                              min_row=data_start, max_row=data_end)
        chart.add_data(data_ref, titles_from_data=True, from_rows=True)

        categories_ref = Reference(data_ws, min_col=2, max_col=last_col,
                                    min_row=end_date_row, max_row=end_date_row)
        chart.set_categories(categories_ref)
        # set_categories() 預設寫 numRef，這裡手動換成 strRef（見下方修復紀錄）
        str_categories_ref = StrRef(f=str(categories_ref))
        for series in chart.series:
            series.cat = AxDataSource(strRef=str_categories_ref)

        n_periods = last_col - 1
        tick_skip = max(1, n_periods // 15)
        chart.x_axis.tickLblSkip = tick_skip
        chart.x_axis.tickMarkSkip = tick_skip

        chart_ws = wb.create_sheet(_chart_sheet_name(metric_name))
        chart_ws.add_chart(chart, "B2")
```

用的套件是 **openpyxl**（純 Python 寫 xlsx，不呼叫真正的 Excel）。

---

## 3. 已經修復、驗證過的項目（背景，供 review 判斷哪些不用重查）

這次 session 用 **PowerShell 呼叫本機安裝的 Excel COM 自動化**（`New-Object -ComObject Excel.Application`）做了兩件事：① 把 openpyxl 產出的圖表匯出成 PNG 肉眼比對；② 建一個「Excel 原生 `ChartObjects.Add()` + `SetSourceData()`」的對照組圖表，把兩份底層 XML（`xl/charts/chart1.xml`）逐項比對，找出 openpyxl 沒寫但 Excel 原生一定會寫的屬性。找到並修好 3 個：

1. **`set_categories()` 永遠寫 `<c:numRef>`（數值參照），不管儲存格是不是文字。** 我們的期末結算日欄位是文字（`"20240331"`，openpyxl 寫入時是純字串，不是 Excel 日期型別——刻意這樣存是為了給 `Snapshot` 分頁的 `SUMPRODUCT`/`MATCH` 公式比對用）。Excel 拿到指向文字儲存格的數值參照解析不出來，類別軸整個讀不到值。**修法**：`chart.set_categories()` 呼叫後，手動把每個 `series.cat` 換成 `AxDataSource(strRef=StrRef(f=...))`。
2. **兩個軸都沒有 `<c:delete>` 元素。** OOXML 規格上沒寫預設值是 `false`，但實測 Excel 渲染時對「沒寫」和「明講 `delete=0`」處理不同——沒寫時保守地不畫任何刻度標籤（無論 X 軸日期或 Y 軸數字）。**修法**：`chart.x_axis.delete = False`、`chart.y_axis.delete = False`。
3. **`Legend` 沒有 `overlay` 屬性。** 沒寫時圖例會跟 X 軸標題/刻度標籤擠進同一條窄帶、直接疊字。原生 Excel 輸出一定帶 `overlay="0"`。**修法**：`chart.legend.overlay = False`。

同時補上 `axPos`（`catAx` 該是 `"b"`，openpyxl 預設是 `"l"`）、`tickLblPos="nextTo"`、`crosses="autoZero"`，這些也是原生輸出必有、openpyxl 不主動寫的欄位。

**這三個修復已確認有效**：用 INTC/NVDA/AMD 真實資料重新產出 Excel，Excel COM 匯出圖片後 Y 軸數字（0 ~ 90,000.0，帶 `$mm`）、X 軸日期、圖例都正常顯示，不再是空白/疊字。程式測試（`tests/test_comparison_writer.py`）全數通過（975 passed, 7 skipped，整個專案）。

**但 CTH 回報最新截圖顯示，還有兩類問題沒解決**（本文件正文）。

---

## 4. 問題一：資料缺漏／斷線

### 4.1 CTH 的觀察

截圖裡 AMD／INTC／NVDA 三條線在時間軸上有好幾處看起來「斷開」，不是連續平滑的線。CTH 懷疑「資料可能缺漏，常有斷線」。

### 4.2 目前已確認的缺漏原因（來自這次 session 的實測，非猜測）

直接讀 `output/_compare/INTC_NVDA_AMD_TSM_v3.xlsx` 的 `Compare_Data` 分頁，逐欄列出三家公司在每個期間的值，缺值（`None`）的期間如下：

| 期間 | 缺哪家 | 已知原因 |
|---|---|---|
| FY2009Q2, FY2009Q3, FY2010Q1 | AMD、NVDA | **SEC 端限制，不是程式 bug**：直接用 edgartools 呼叫這幾份 10-Q 的 `filing.xbrl()`，回傳空——這幾份是純 HTML 申報，SEC 對 XBRL 的強制申報是 2009-2011 分階段上路，這幾份早於 AMD 被要求的時間點，檔案本身就沒有結構化數字可抓 |
| FY2009Q4 | 全部 | 尚無 10-K 年報可合成（公司歷史更早的年報未涵蓋到，或年報抓取範圍未觸及） |
| FY2010Q4, FY2011Q4, FY2012Q4（NVDA） | NVDA（AMD 同理） | **連鎖失敗**：Q4 合成需要「年報－Q1－Q2－Q3」三季都在，FY2010/2011/2012 的 Q1 原始資料本身就缺（同上一列的 SEC 端限制延伸到後續兩年），連帶讓這三年的 Q4 合成不出來。`_synthesize_q4()` 的判斷式在 `fetcher_gaap.py:1691`：`if not (q1 in q_idx and q2 in q_idx and q3 in q_idx): continue` |
| **FY2017Q4, FY2022Q4, FY2023Q4（NVDA）** | NVDA | ⚠️ **2026-08-22 新查證：不是 SEC 缺資料，是我們自己的程式 bug，而且已經定位到確切那一行**——詳見下方 4.2.1，獨立成一節，因為這是這次最重要的發現 |
| FY2026Q3, FY2026Q4, FY2027Q1 | AMD、INTC（NVDA 有值） | **不是缺漏，是還沒發生**：今天是 2026-08-22，NVDA 財年結束在 1 月，這幾期對 NVDA 是已公布的歷史數字；AMD/INTC 財年跟日曆一致，這幾期是「未來」，本來就不該有數字 |

### 4.2.1 ⚠️ 新發現的真正 bug：quarter label 碰撞導致資料被靜默丟棄（非 SEC 端限制）

**這是 2026-08-22 CTH 直接要求「用我們目前邏輯去 call API 看是否空白」之後查出來的，是本文件最重要的新發現，推翻了原本「可能是已知限制」的假設。**

**第一步，確認 SEC 端資料其實都在**：直接呼叫 `fetch_gaap_statements('NVDA', ..., fetch_quarterly=True, fetch_annual=True)`，印出所有期間的 Revenue 值，`FY2017Q4`／`FY2022Q4`／`FY2023Q4` 確實是 `None`，但同時發現這三個財年**各自的 Q1 也整個不見**（`quarter_labels` 清單裡直接跳過，不是有欄位但值是 None，是整欄不存在）——這正是 Q4 合成失敗的原因（見上表）。

**第二步，直接查證 Q1 的原始 XBRL 是否真的存在**。用 edgartools 直接抓 NVDA 2016-05-25 申報的那份 10-Q（`accession_no='0001045810-16-000275'`，`period_of_report='2016-05-01'`——這正是 FY2017 應該存在的 Q1）：

```python
xbrl = filing.xbrl()
print(xbrl is not None)          # True —— 有 XBRL，資料真的存在
print(xbrl.period_of_report)     # 2016-05-01
```

**確認 XBRL 完整存在，不是 SEC 端沒資料。** 2021-05-26（FY2022 Q1）、2022-05-27（FY2023 Q1）兩份也同樣測過，`xbrl() is not None` 都是 `True`。

**第三步，找到真正的原因：`_col_to_quarter_label()` 算出來的財季編號跟實際財季對不上，導致 label 碰撞、後到的資料被吃掉。** 直接對這份 2016-05-01 的 10-Q 呼叫我們自己的解析函式：

```python
df = fin.income_statement().to_dataframe()
print(df.columns)
# ['concept', 'label', 'standard_concept',
#  '2016-05-01 (Q2)', '2015-04-26 (Q1)', ...]
q_col = _current_q_col(df)          # -> '2016-05-01 (Q2)'
label = _col_to_quarter_label(q_col, fy_end_month=1)   # -> 'FY2017Q2'
```

**edgartools 自己把這份 10-Q 的當期欄位標成 `(Q2)`，不是 `(Q1)`！** 我們的 `_col_to_quarter_label()`（`fetcher_gaap.py:431`）只是原樣採信 edgartools 欄名裡的 `(Qn)` 標記去算財季編號，沒有自己交叉驗證這個編號是否跟財年起始月份、期末日期算出來的財季吻合。這份 2016-05-01（NVDA 財年 2 月開始，這期落在 2-4 月，理論上是 FY2017 的 **Q1**）被 edgartools 標成 `Q2`，我們的程式全盤接受，算出 `FY2017Q2`。

**接下來是致命的一步：`_build_is_table()` 裡的 dedup 邏輯（`fetcher_gaap.py:864`）**：

```python
label = _col_to_quarter_label(q_col, fy_end_month)
if label in periods:
    continue          # ← 靜默跳過，不記錄、不警告、不寫 log
```

如果**另一份**真正的 FY2017Q2 10-Q（期末日期實際落在 2016 年 7-8 月，`period_end='2016-07-31'`，這是我們實測拿到的真正 FY2017Q2 資料）**也**被算成 `label == "FY2017Q2"`，兩者剛好撞名。`filings` 是新到舊排序，**先處理到的那份會被存進 `periods` 字典，後處理到的那份因為 `label in periods` 直接 `continue` 跳過**，資料就這樣不見了——不是抓不到，是抓到了兩份、留一份丟一份，而且丟的時候完全沒有任何 log 或警告，使用者／開發者都看不出來曾經抓到過。

**這解釋了 FY2017Q4／FY2022Q4／FY2023Q4 為什麼缺**：不是那三年的 Q4 本身有問題，是那三年的 **Q1** 因為 label 碰撞被靜默丟棄，連帶讓 `_synthesize_q4()` 判斷「Q1/Q2/Q3 都要有」時因為缺 Q1 而跳過合成。

**還沒查證、留給 review 的部分**：
1. 為什麼 edgartools 會把 NVDA 這份 2016-05-01 的 10-Q 欄位標成 `(Q2)` 而不是 `(Q1)`？是 edgartools 本身的 bug／對 NVDA 52/53 週財年處理不一致，還是這份 10-Q 的 XBRL 原始資料本身在申報時的 `dei:DocumentFiscalPeriodFocus` 標籤就標錯／標成 `Q2`（若是後者，那是申報公司/SEC 端的資料品質問題，不是 edgartools 或我們的 bug）？需要直接檢查這份 filing 的原始 XBRL `dei:DocumentFiscalPeriodFocus` 值
2. 這個「兩份不同期間的 10-Q 算出同一個 label」的碰撞是不是只發生在 NVDA（52/53 週財年、非 12 月結算，`fy_end_month` 換算灑輯較複雜），還是其他公司也會發生只是這次沒踩到
3. **修法方向的建議**（未實作，供評估）：`_build_is_table()` 的 dedup 判斷式（`fetcher_gaap.py:864`）發生碰撞時，至少應該印一行警告到 stderr／log（現在完全靜默），讓使用者知道「這裡有兩份 filing 撞到同一個標籤，其中一份被丟棄」——就算不修碰撞本身的根因，至少不要靜默失敗。更根本的修法可能是 `_col_to_quarter_label()` 不能隻信任 edgartools 的 `(Qn)` 標記，要自己用 `_col_to_period_end()` 解析出的實際日期反推正確的財季編號，兩者不一致時要有明確的仲裁規則（例如以日期反推為準，edgartools 標記僅供參考）

### 4.3 給 review 的具體問題

1. **最優先**：4.2.1 節的 label 碰撞問題影響範圍多大？除了 NVDA 這 3 個財年，其他公司（尤其非 12 月結算的公司，如 AAPL 9 月結算）有沒有同樣的碰撞？建議寫一個一次性稽核腳本，對現有測試覆蓋的公司清單，比對「filings 數量」跟「最終 `quarter_labels` 數量」的差距，數量對不上就代表發生過靜默丟棄
2. `max_filings=80`（`fetcher_gaap.py` 預設值）是否真的足夠涵蓋 AMD/INTC 從 2009 到 2026 的完整範圍？（這點還沒排除，但 4.2.1 的發現讓它的優先度下降——目前找到的缺洞都能歸因到更具體的原因了）
3. `display_blanks = "gap"` 讓缺值處"不連線"這個設計本身是對的（不能用假造的連線騙使用者），但既然 4.2.1 證實至少一部分缺洞是可修的程式 bug 而非資料源限制，優先順序應該是先修 bug 把缺洞補起來，而不是停在「顯示斷點」這個治標的處理

---

## 5. 問題二：美編／排版問題

### 5.1 CTH 最新截圖裡看到的兩個具體位置

（原始截圖：3 家公司 Revenue 折線圖，X 軸 2009~2027，Y 軸 0~90,000）

1. **Y 軸標題「Revenue ($mm)」（旋轉 90 度的直式文字，貼在 Y 軸數字左邊）跟 Y 軸的「50,000.0」這個刻度數字視覺重疊**——兩段文字疊在一起看不清楚，其他刻度數字（10,000.0、20,000.0...）沒有這個問題，只有大概在垂直置中位置的那一格被蓋到。
2. **X 軸標題「期間」（`chart.x_axis.title = t("gui.compare.period")`，中文「期間」兩個字）沒有乖乖待在整排 X 軸日期標籤下方置中，而是跑到日期標籤那一排的「中間」，跟其中一個日期標籤（視覺上接近 "20181229" 那格）擠在一起、部分重疊。**

### 5.2 目前的假設（未驗證，需要 review 協助判斷或提出更好的診斷方法）

- **Y 軸標題重疊**：懷疑是 Excel 自動計算「Y 軸標題應該離刻度數字多遠」時，用的是**文字框固定位置**而不是「動態量測刻度數字最長字串寬度後再往外推」，openpyxl 沒有寫任何 `layout`／`manualLayout` 讓標題明確定位，所以 Excel 用了某個預設偏移量，剛好在這組數字/這個圖表尺寸下跟置中的刻度數字打架。
- **X 軸標題位置錯亂**：懷疑原因是**圖表寬度 30cm 塞了很多刻度標籤**（`tickLblSkip` 讓 69 期只顯示約 17 個，但 17 個日期字串已經佔滿一整排寬度），Excel 原本應該把 X 軸標題放在「整條刻度標籤帶」的正下方置中，但可能因為 openpyxl 沒有給 `<c:layout>` 明確定位，Excel 的 auto-layout 在標籤帶很寬、很滿的情況下算錯了標題該放哪一列，於是掉進標籤帶內部而不是下方。這個假設**還沒有像前面三個 bug（numRef/delete/overlay）一樣做「跟原生 Excel 圖表 XML 逐項比對」去確認**，所以只是假設，不是實測結論。

### 5.3 已經試過但沒解決這兩個問題的東西

前面第 3 節列的三個修復（strRef／delete=False／overlay=False）解決的是「完全看不到文字」跟「圖例整條疊在日期上」這兩個**更嚴重**的問題。這次新發現的兩個重疊，是那三個修復生效**之後**才顯現出來的**次一層**排版問題——某種意義上是「東西終於畫出來了，但畫的位置還沒對齊」。

### 5.4 給 review 的具體問題

1. openpyxl 的 `openpyxl.chart.axis` / `openpyxl.chart.title` / `openpyxl.chart.layout` 有沒有辦法明確指定「軸標題」的位置或偏移量（例如 `Layout(manualLayout=ManualLayout(...))`)，讓它不要跟刻度標籤打架？有沒有人遇過同樣的 openpyxl 圖表軸標題位置問題並有已知解法？
2. 用「跟原生 Excel 圖表 XML 逐項比對」這個方法（本文件第 3 節的手法）能不能繼續用來找這兩個新問題？也就是說：拿掉 X 軸標題（`chart.x_axis.title` 設 `None`）或縮短刻度標籤數量，看重疊是不是消失，藉此縮小觸發條件。
3. 有沒有可能單純是「圖表高度 15cm 對這麼多標籤/標題/圖例元素來說仍然不夠高」，把 `chart.height` 從 15 拉到例如 18~20cm 就能讓 Excel 的 auto-layout 有足夠空間正確擺放，不用手動指定 layout？（這是最簡單、風險最低的嘗試方向，但還沒試過）

---

## 6. 相關檔案清單

| 檔案 | 內容 |
|---|---|
| `src/fetcher_gaap.py` | 抓取單一公司 XBRL 財報，含 Q4 合成邏輯 `_synthesize_q4()` |
| `src/comparison.py` | 跨公司抓取協調、`build_comparison()` |
| `src/comparison_writer.py` | 寫 Excel（`Compare_Data`／`Snapshot`／`Chart_<指標>`），本文件問題集中在 `write_chart_sheets()` |
| `tests/test_comparison_writer.py` | 對應的單元測試，含這次新增的 delete/overlay/strRef 測試 |
| `docs/TODO.md` | F3 條目記錄了完整的修復歷程與時間軸 |
| `docs/CHANGELOG.md` | 同上，含每次修復的詳細技術說明 |
| `output/_compare/INTC_NVDA_AMD_TSM_v3.xlsx` | 目前這次 session 產出的實際檔案，CTH 最新截圖就是這份檔案的 `Chart_Revenue` 分頁 |

---

## 8. 附帶研究：8-K／10-K 抓取功能是否正常

CTH 額外要求確認一下 8-K（法說會新聞稿）與 10-K（年報）的抓取功能現在還能不能正常運作（跟本文件主題「跨公司比較圖表」不同功能，但同一個 codebase，一併記錄）。**這裡的測試都是真的打 SEC EDGAR API，不是 mock。**

### 8.1 10-K（年報）抓取

直接呼叫 `fetcher_gaap.fetch_gaap_statements(ticker, identity, fetch_quarterly=False, fetch_annual=True)` 對 AAPL 實測：

```
periods: ['FY2023', 'FY2024', 'FY2025']
period_ends: ['2023-09-30', '2024-09-28', '2025-09-27']
Revenue: [383285000000.0, 391035000000.0, 416161000000.0]
```

三個年度都正確抓到，Revenue 數字跟 Apple 公開財報數量級吻合（FY2023 ~$383B、FY2024 ~$391B、FY2025 ~$416B）。**結論：10-K 抓取功能正常**。這也是本文件第 2.1／2.2 節提到的 Q4 合成邏輯的資料來源——Q4 合成失敗如果要排查，10-K 這一段抓取本身已經確認沒問題，問題若存在應該在合成/科目對照那一層，不是抓不到年報本身。

### 8.2 8-K（法說會新聞稿）抓取

這條路線走的不是 GAAP 主流程，是 `fetcher_nongaap.py` + `press_release_tables.py`（TODO B3，確定性表格解析，不靠 AI）。用 `cli.py press-release` 子指令對 ARLO 實測：

```
venv/Scripts/python.exe src/cli.py press-release ARLO --years 2025-2026 --tables --json
```

成功抓到 8-K 的表格資料（U.S. retail channel、Deferred revenue、Cumulative paid accounts、ARR、Headcount、Non-GAAP diluted shares 等多個表格，每個表格橫跨 5 期），JSON 輸出結尾 `"skipped": []`——代表這個 ticker/年份範圍內沒有任何一份 8-K 被跳過。**結論：8-K 抓取功能正常**。

### 8.3 補充：這條功能目前在 GUI 是關閉的

`main.py:377`：`NONGAAP_ENABLED = False`。Non-GAAP／8-K 相關功能目前**只能透過 `cli.py` 呼叫**，GUI（Tab1 單一公司抓取畫面）不會顯示、不會執行這段邏輯（`main.py` 裡兩處 `if fetch_nongaap and NONGAAP_ENABLED:` 判斷式擋住）。這不是功能壞掉，是刻意關閉——`docs/TODO.md` 的 B 段（B2）記錄了「Non-GAAP 改走 skill 端抽取」的後續規劃，這條底層抓取邏輯本身是給那個 skill 呼叫用的介面，不是要接回 GUI。

### 8.4 這兩個結論跟本文件正題（跨公司比較圖表）的關聯

- 10-K 抓取沒問題，代表本文件第 4 節「Q4 合成缺洞」如果要繼續查，排查方向應該放在 `_synthesize_q4()` 的合成邏輯本身、科目對照，而不是懷疑年報抓不到
- 8-K 抓取沒問題，但這條路線目前跟「跨公司比較」功能完全沒有交集（跨公司比較走的是 10-Q/10-K GAAP 主流程，不會用到 8-K），純粹是這次順便確認的附帶資訊，**不影響**前面第 4、5 節的問題判斷

## 9. 環境資訊

- Python 3.13，openpyxl（版本見 `requirements.txt`），純程式產出 xlsx，過程中**不會**打開真正的 Excel
- 這次除錯用的驗證方式：本機（Windows）剛好裝有 Microsoft Office，改用 PowerShell 呼叫 `New-Object -ComObject Excel.Application` 做自動化，可以真的用 Excel 開啟、匯出圖片、讀取渲染後的屬性（`Chart.PlotArea.InsideWidth` 等）、比對 XML。這個方法在沒有裝 Office 的環境下不能用，但可以用純 XML 檢查（`zipfile` 讀 `xl/charts/chart1.xml`）驗證「寫了什麼」，只是驗證不了「Excel 實際怎麼渲染」
