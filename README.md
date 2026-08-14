規則檔: windows-tool.md

# SEC Financial Fetcher

股票分析師工具：從 SEC EDGAR 抓取美國上市公司 GAAP 財報並存成 Excel。

## 執行方式

雙擊 `啟動器.bat`

### 指令列（給 skill 用，不經 GUI）

```bash
# GAAP 三表 + 比率 + segment → Excel（與 GUI 產的逐格相同）
./venv/Scripts/python.exe src/cli.py gaap AAPL --years 2023-2026 --xlsx out.xlsx

# 8-K 新聞稿的 Non-GAAP 調節表（已解析、已篩過）→ JSON
./venv/Scripts/python.exe src/cli.py press-release ARLO --years 2025-2026 --tables --json
```

兩個子指令都**不呼叫任何 AI API**，只打 SEC EDGAR。共通參數：`--years`
（`2023-2026` 或 `2024`）、`--identity`、`--max-filings`、`--json`（不給路徑
就印到 stdout）、`--lang`（產出 Excel 的顯示語言：`zh_tw` / `zh_cn` / `en` /
`ja`，不給就跟 GUI 用同一個設定；只影響 B 欄與 Index 版面，A 欄機器鍵與 C 欄
公司原文不變）。`gaap` 另有 `--xlsx` / `--quarterly-only` / `--annual-only`，
`press-release` 另有 `--raw`（改吐新聞稿全文，除錯用）。

`press-release` 吐的是**解析後的表格**不是原文：ARLO 一季原文 450K 字元，
篩完 4.4K。

⚠ **季度標籤看 `fiscal_label`，不要看 `label`。** `label` 是用 8-K 的
`period_of_report`（＝發布日）換算的，有系統性 off-by-one（偏 −3 到 +1 季，
見 `docs/8k-period-off-by-one.md`），為了不破壞既有介面而保留原值並帶著
`label_warning`。`fiscal_label` 是從新聞稿表格裡的**期末日**（`period_end`）
加公司財年結束月（`fy_end_month`，payload 頂層）算出來的，與 `Data_Q` 的
財季同一套慣例，兩邊對得起來。抓不到期末日時 `fiscal_label` 留空，不會用
發布日硬算。15 家 120 份實測全部抓得到。

⚠ **`--years` 篩的是發布日不是財期**：篩選發生在下載之前，那時還讀不到
期末日。非 12 月結算的公司在年份邊界可能差到 3 季，要精確就把範圍放寬一年，
再自己用 `fiscal_label` 篩。

## 系統需求

- Windows 10/11
- 需要網路連線（首次安裝 + 每次抓取資料）

## 首次設定

1. 雙擊 `啟動器.bat`，按照提示完成套件安裝
2. 程式啟動後點「進階設定」，填入 SEC EDGAR Identity（姓名 + 信箱）
3. 若要使用 Non-GAAP 功能，在進階設定填入 AI API Key
4. 要換介面語言的話，「進階設定」最上方的 `Language` 選單有繁體中文／简体中文／English／日本語。選完會跳一個英文視窗問要不要重啟，按 Restart 就直接換好

## Excel 結構

每間公司一個 `.xlsx`，存於 `output/` 資料夾。

| Sheet | 說明 |
|-------|------|
| `Data_Financials(Q)` | **季報三表**（IS + BS + CF，from 10-Q）。表頭 3 列為期間標籤，三表各有專屬底色，公司特有科目集中在底部 `Other (as reported)` |
| `Data_Financials(Y)` | **年報三表**（from 10-K），結構同上 |
| `Data_Ratios` | 37 個常見比率（Python 計算，**零 AI**）。A 欄英文列名（含 `(%)` / `(x)` / `(days)` / `($)` 單位後綴）、B 欄說明、C 欄算法 |
| `Data_Segments` | 營收／費用分類細項，長格式（各軸合併於一張） |
| `Data_Meta` | 申報資訊（Ticker、公司名、抓取日期、季數、財年結束月） |
| `Index` | 第一頁：公司抬頭、**財年起始月輸入格**、sheet 清單、品質明細 |

> **列位跨公司固定**：NVDA／AAPL／PLTR／AVGO 實測（財年結束月分別是 1／9／12／11 月），`Revenue` 都在第 8 列、`Gross Profit` 10、`Operating Income` 17、`Net Income` 24、`Cash` 38、`Total Assets` 51、`Operating Cash Flow` 98、`Capex` 99、`Free Cash Flow` 114。這是因為公司特有科目（overflow）集中在底部，不再插在 section 之間。跨檔案公式可以直接用固定儲存格參照。

**欄位說明（Data_Financials）：**

| 位置 | 內容 |
|---|---|
| A 欄 | 標準指標名稱（**永遠英文**，程式一律用這欄比對，跨檔案 `MATCH` 也吃這欄） |
| B 欄 | 說明（**跟著介面語言走**：繁中／简中／English／日本語） |
| C 欄 | Original Item（**永遠是公司財報上的英文原文**，如 AAPL 的 `Net sales`。拿它去 10-Q 裡 Ctrl+F 核對用，所以不翻譯） |
| D 欄起 | 各期數據（舊→新） |
| 第 1 列 | 期間標籤（`FY2026Q1`）**← 公式** |
| 第 2 列 | 申報日期 |
| 第 3 列 | 財季（`FY2026FQ1`，公司財年基準）**← 公式** |
| 第 4 列 | 日曆季（`2026Q1`，日曆年基準）**← 公式** |
| 第 5 列 | 期末結算日（XBRL 真實日期）**← 靜態，是上面三列公式的錨** |

## 財年起始月：程式猜錯時自己改

財年結束月是程式從 10-K 自動判讀的，**會出錯**。所以 `Index!B4`（黃底那格）是可以改的：

```
Index
  B4  = 10        ← 財年起始月，AAPL 是 10 月
```

改這一格，`Data_Financials(Q)/(Y)` 第 1、3、4 列的期間標籤會**全部自動更新**（公式引用定義名稱 `FY_START_MONTH`）。

**怎麼核對**：看第 5 列的期末結算日——那是 XBRL 的真實日期，一定正確。拿它對照公司財報上寫的財季。例如 AAPL 的 `2025-12-27` 公司叫它 FY2026 Q1，第 1 列就該顯示 `FY2026Q1`。

換算會先把期末日**往前推 15 天**再取年月。美股多用 52/53 週制，期末日在月底前後浮動最多 6 天（WDC 的 FY2026 Q2 結束在 `2026-01-02`），直接看月份會整整差一季。

> ⚠ 不會跟著變的：`Index` 表格的「最早／最新期間」、`Data_Ratios`、`Data_Meta`。這三個是 Python 算好寫死的。

**Section header 行：**
`Data_Financials` 內有三段分隔行（`Income Statement` / `Balance Sheet` / `Cash Flow`），資料值全為空。

分析用的自訂 Sheet 請命名為 `My_*`（如 `My_IS`），Python 不會碰這些 Sheet。

## 模板行數

| 報表 | 行數 | 說明 |
|------|------|------|
| Income Statement | 22 | 含 D&A/SBC/Minority Interest/Total Non-op |
| Balance Sheet | 42 | Assets 14 行、Liabilities 17 行、Equity 11 行（含期末流通股數） |
| Cash Flow | 25 + 1 | Operating/Investing/Financing + FCF 衍生 |

沒有資料的項目顯示空白（None），不影響其他行。

## 已知限制

- **`Data_Financials(Q)` 沒有 Q4**：Q4 沒有 10-Q，數字在 10-K 裡。連帶 TTM 類比率（ROE／ROA／FCF per Share／淨負債EBITDA）湊不到連續四季，多半是空的。
- **多股別公司抓不到流通股數**：PLTR／GOOGL／META 有 Class A/B/C，封面頁的 `dei:EntityCommonStockSharesOutstanding` 按股別分開標，`company.get_facts()` 取不到。連帶 BVPS、FCF per Share、流通股數 YoY 空白。
- 金融股（GS、JPM 等）：BS/IS 結構與一般公司不同，部分項目會是空白。金融股模板尚未實作（計畫中）。
- `Investment Proceeds`：XBRL 沒有單一加總行，取第一筆（已知缺陷）。

## 寫跨公司模板

以前這一段講的是 `Data_Std`——那張 sheet 已經在 2026-08-03 的輸出精簡裡**併回 `Data_Financials(Q)` 並刪除**了。列位固定與期間標籤現在直接長在三表本身，不需要另一張表。

三個保證：

1. **固定 sheet 名稱**：`Data_Financials(Q)` / `Data_Financials(Y)`
2. **固定列位**：公司特有科目（overflow）全部集中在最底部，不再插在 section 之間，所以 `Revenue` 永遠在第 8 列（見上方對照表）
3. **三種期間標籤各佔一列**，用途不同（見下）

### 三種期間標籤，用途不同

| 列 | 內容 | 基準 | 什麼時候用 |
|---|---|---|---|
| 1 | `FY2026Q1` | **公司財年** | 主要欄位鍵，`MATCH` 用這列 |
| 3 | `FY2026FQ1` | **公司財年** | 同上，加 `FQ` 標記避免與日曆季混淆 |
| 4 | `2026Q1` | **日曆年** | 跨產業比同一個日曆期間、對總經數據 |
| 5 | `2026-03-29` | 實際期末日 | 精確對齊；也是判斷兩家是否真的同期的唯一依據 |

12 月結算的公司三者一致；非 12 月結算的差很多：

| 公司 | 結算月 | 財季（列 1/3） | 日曆季（列 4） | 期末結算日（列 5） |
|---|---|---|---|---|
| PLTR | 12 月 | 2026Q2 | 2026Q2 | 2026-06-30 |
| AAPL | 9 月 | **2026Q1** | **2025Q4** | 2025-12-27 |
| NVDA | 1 月 | **2027Q1** | **2026Q2** | 2026-04-26 |

**期末結算日是真實日期不是月底**——美股多用 52/53 週制，AVGO 的 FY2026Q2 結束在 05-03 而不是 04-30。這個日期直接來自 XBRL，不是推算的。

### 公式怎麼寫

比財季（同業比較）：

```excel
=INDEX('[AAPL.xlsx]Data_Financials(Q)'!$D8:$AZ8,
       MATCH("FY2026Q1",'[AAPL.xlsx]Data_Financials(Q)'!$D$1:$AZ$1,0))
```

比日曆季（跨產業／對總經）把 `$1` 換成 `$4`（值變成 `2026Q1` 不含 FY）。`$D8:$AZ8` 的 8 是營收列，換 `$D38` 就是現金。

> ⚠ 第 1、3、4 列現在是**公式**（由 `Index!B4` 驅動）。同一個活頁簿內、或來源檔開著時，`MATCH` 照樣比對得到公式算出來的結果。

> ⚠ **跨檔案讀取、而且來源檔關著時，改用第 5 列當 key**。公式沒有快取值（openpyxl 不算公式），來源檔關著時外部參照只讀得到檔案裡的值——第 1、3、4 列在那裡是空的，`MATCH` 會回 `#N/A`。第 5 列（期末結算日）是靜態文字，永遠讀得到：
>
> ```excel
> =INDEX('C:\...\output\_final\[AAPL.xlsx]Data_Financials(Q)'!$D8:$AZ8,
>        MATCH("2025-12-27",'C:\...\output\_final\[AAPL.xlsx]Data_Financials(Q)'!$D$5:$AZ$5,0))
> ```
>
> 或者把來源檔在 Excel 開一次再存檔，快取值就寫進去了，之後關著也能用 `FY2026Q1` 當 key。
