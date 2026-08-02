規則檔: windows-tool.md

# SEC Financial Fetcher

股票分析師工具：從 SEC EDGAR 抓取美國上市公司 GAAP 財報並存成 Excel。

## 執行方式

雙擊 `啟動器.bat`

## 系統需求

- Windows 10/11
- 需要網路連線（首次安裝 + 每次抓取資料）

## 首次設定

1. 雙擊 `啟動器.bat`，按照提示完成套件安裝
2. 程式啟動後點「進階設定」，填入 SEC EDGAR Identity（姓名 + 信箱）
3. 若要使用 Non-GAAP 功能，在進階設定填入 AI API Key

## Excel 結構

每間公司一個 `.xlsx`，存於 `output/` 資料夾。

| Sheet | 說明 |
|-------|------|
| `Index` | 索引頁：所有 sheets 清單、時間範圍、完成度欄（9 個 key rows 的 ✓/✗ 明細） |
| `Data_Financials(Q)` | IS + BS + CF 三表合一（季報，from 10-Q），固定行數萬能模板 |
| `Data_Financials(Y)` | IS + BS + CF 三表合一（年報，from 10-K），固定行數萬能模板 |
| `Data_Financials_NG(Q/Y)` | Non-GAAP overflow 行（含 "adjusted"/"non-gaap" 等 label 的 XBRL 行），有資料才產生 |
| `Data_Seg_*` | 各收入/費用的地區/業務分類細項 |
| `Data_NonGAAP` | AI 從 8-K press release 提取的 Non-GAAP 指標（勾選 Non-GAAP 時產生） |
| `Data_Std` | **跨公司標準表**：列位完全固定、含日曆季標籤與機器鍵，寫通用模板參照這張 |
| `Data_Segments` | 各分類軸合併的長格式表（一張表涵蓋所有軸，給程式讀；寬格式見 `Data_Seg_*`） |
| `Data_Ratios` | 常見財務比率（37 項，自 `Data_Financials(Q)` 計算，B 欄寫算法） |
| `Data_Meta` | 申報資訊（Ticker、公司名、抓取日期、季度數） |

**欄位說明（Data_Financials）：**
- A 欄 = 標準指標名稱（Std Name）
- B 欄 = Original Item（公司的 XBRL 原始標籤）
- C 欄起 = 各季數據（舊→新）
- 第 1 列 = 季度標籤（如 FY2024Q1）
- 第 2 列 = 申報日期

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

- 金融股（GS、JPM 等）：BS/IS 結構與一般公司不同，部分項目會是空白。金融股模板尚未實作（計畫中）。
- `Investment Proceeds`：XBRL 沒有單一加總行，取第一筆（已知缺陷）。

## 用 `Data_Std` 寫跨公司模板

其他 sheet 的版面會隨公司變動（sheet 數 10～30 張、`Cash` 落在第 28～56 列、季度欄 4～50 欄），沒辦法用固定儲存格參照。`Data_Std` 是為此而生的：**同一個科目在每家公司都在同一列**。

```
      A            B                        C          D          E
1   ARLO                                FY2025Q1   FY2025Q2   FY2025Q3   ← 財季（公司自己的）
2                                       2025-05-08 2025-08-07 2025-11-06 ← 申報日
3   日曆季        META.CALENDAR_QUARTER    2025Q1     2025Q2     2025Q3   ← 跨公司對齊用這列
4   期末年月      META.PERIOD_END         2025-03    2025-06    2025-09
5   資料版本      META.SCHEMA              STD_V1
6   Income Statement
7   Revenue      IS.REVENUE                119.1      129.4      139.5
...
30  Cash         BS.CASH                   84.0       71.2       86.0
98  Free Cash Flow  CF.FREE_CASH_FLOW      28.1        5.5       19.1
108 毛利率 (%)    RATIO.毛利率              44.3%      44.9%      40.5%
```

**⚠ 第 1 列是財季不是日曆季。** 非 12 月結算的公司差很多：

| 公司 | 結算月 | `FY2025Q1` 實際期間 | 日曆季 |
|---|---|---|---|
| ARLO | 12 月 | 2025 年 3 月底 | 2025Q1 |
| AAPL | 9 月 | 2024 年 12 月底 | **2024Q4** |
| NVDA | 1 月 | 2024 年 4 月底 | **2024Q2** |

**跨公司比同一個日曆期間，一律對第 3 列比對**，不要用第 1 列，也不要假設欄位置一樣：

```excel
=INDEX('[ARLO.xlsx]Data_Std'!$C7:$AZ7,
       MATCH("2025Q3",'[ARLO.xlsx]Data_Std'!$C$3:$AZ$3,0))
```

想比同一家公司的營運週期（例如零售業的旺季）就改對第 1 列的財季標籤。

**列號 vs 機器鍵**：直接用列號（`C7` = 營收）最快；要更保險就用 B 欄機器鍵做 `MATCH`，即使日後調整列序也不會壞。列號本身有 `FROZEN_ROW_NUMBERS` 的測試釘住，任何改動都會讓測試紅掉。
