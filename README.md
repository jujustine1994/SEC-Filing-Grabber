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
