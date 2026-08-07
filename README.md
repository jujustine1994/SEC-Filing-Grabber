規則檔: windows-tool.md

# SEC Financial Fetcher

股票分析師工具：從 SEC EDGAR 抓取美國上市公司 GAAP 財報並存成 Excel。

## 執行方式

雙擊 `啟動器.bat`

### 指令列（給 skill 用，不經 GUI）

```bash
# GAAP 三表 + 比率 + segment → Excel（與 GUI 產的逐格相同）
./venv/Scripts/python.exe cli.py gaap AAPL --years 2023-2026 --xlsx out.xlsx

# 8-K 新聞稿的 Non-GAAP 調節表（已解析、已篩過）→ JSON
./venv/Scripts/python.exe cli.py press-release ARLO --years 2025-2026 --tables --json
```

兩個子指令都**不呼叫任何 AI API**，只打 SEC EDGAR。共通參數：`--years`
（`2023-2026` 或 `2024`）、`--identity`、`--max-filings`、`--json`（不給路徑
就印到 stdout）。`gaap` 另有 `--xlsx` / `--quarterly-only` / `--annual-only`，
`press-release` 另有 `--raw`（改吐新聞稿全文，除錯用）。

`press-release` 吐的是**解析後的表格**不是原文：ARLO 一季原文 450K 字元，
篩完 4.4K。⚠ 每一季的 `label` 有已知的 off-by-one（見
`docs/8k-period-off-by-one.md`），JSON 裡帶著 `label_warning`，實際財期
請看表格的期間表頭。

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
| `Data_Financials(Q)` | **季報三表**（IS + BS + CF，from 10-Q）。表頭 3 列為期間標籤，三表各有專屬底色，公司特有科目集中在底部 `Other (as reported)` |
| `Data_Financials(Y)` | **年報三表**（from 10-K），結構同上 |
| `Data_Ratios` | 37 個常見比率（Python 計算，**零 AI**），B 欄寫算法 |
| `Data_Segments` | 營收／費用分類細項，長格式（各軸合併於一張） |
| `Data_Meta` | 申報資訊（Ticker、公司名、抓取日期、季數、財年結束月） |

> **列位跨公司固定**：AAPL／NVDA／META／AVGO 實測，`Revenue` 都在第 8 列、`Cash` 第 34 列、`Free Cash Flow` 第 106 列。這是因為公司特有科目（overflow）集中在底部，不再插在 section 之間。跨檔案公式可以直接用固定儲存格參照。

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

其他 sheet 的版面會隨公司變動（sheet 數 10～30 張、`Cash` 落在第 28～56 列、季度欄 4～50 欄），沒辦法用固定儲存格參照。`Data_Std` 是為此而生：**同一個科目在每家公司都在同一列**。

```
      A              B                        C            D            E
1   ARLO                                  FY2025Q1     FY2025Q2     FY2026Q1    ← 財季標籤
2                                         2025-05-08   2025-08-07   2026-05-07  ← 申報日
3   財季           META.FISCAL_QUARTER      2025Q1       2025Q2       2026Q1     ← 公司財年基準
4   日曆季         META.CALENDAR_QUARTER    2025Q1       2025Q2       2026Q1     ← 日曆年基準
5   期末結算日     META.PERIOD_END        2025-03-30   2025-06-29   2026-03-29   ← 真實期末日
6   資料版本       META.SCHEMA              STD_V1
7   Income Statement
8   Revenue        IS.REVENUE                119.1        129.4        150.4
31  Cash           BS.CASH                    84.0         71.2        152.6
99  Free Cash Flow CF.FREE_CASH_FLOW          28.1          5.5         27.4
109 毛利率 (%)      RATIO.毛利率               44.3%        44.9%        48.3%
```

### 三種期間標籤，用途不同

| 列 | 內容 | 基準 | 什麼時候用 |
|---|---|---|---|
| 3 | `2026Q1` | **公司財年** | 比同一家的營運週期；或比同業（同業財年通常一致） |
| 4 | `2026Q1` | **日曆年** | 跨產業比同一個日曆期間、對總經數據 |
| 5 | `2026-03-29` | 實際期末日 | 精確對齊；也是判斷兩家是否真的同期的唯一依據 |

12 月結算的公司三者一致；非 12 月結算的差很多：

| 公司 | 結算月 | 財季（列 3） | 日曆季（列 4） | 期末結算日（列 5） |
|---|---|---|---|---|
| ARLO | 12 月 | 2026Q1 | 2026Q1 | 2026-03-29 |
| AAPL | 9 月 | **2026Q1** | **2025Q4** | 2025-12-27 |

**期末結算日是真實日期不是月底**——美股多用 52/53 週制，ARLO 的 Q1 2025 結束在 03-30 而不是 03-31。這個日期直接來自 XBRL，不是推算的。

### 公式怎麼寫

比財季（同業比較）：

```excel
=INDEX('[ARLO.xlsx]Data_Std'!$C8:$AZ8,
       MATCH("2026Q1",'[ARLO.xlsx]Data_Std'!$C$3:$AZ$3,0))
```

比日曆季（跨產業／對總經）把 `$3` 換成 `$4`。`$C8:$AZ8` 的 8 是營收列，換 `$C31` 就是現金。

**列號 vs 機器鍵**：直接用列號最快；要更保險就用 B 欄機器鍵做 `MATCH`，即使日後調整列序也不會壞。列號本身有 `FROZEN_ROW_NUMBERS` 測試釘住，任何改動都會讓測試紅掉。
