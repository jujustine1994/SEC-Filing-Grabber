# Changelog

## 現狀

- Phase 1 (GAAP)：萬能模板完成 ✅
- Phase 2 (Non-GAAP)：完成 ✅
- Phase 3 (Excel 美化)：完成 ✅

## 功能清單

### 已完成
- [x] config.json 搬到 %APPDATA%\SEC Financial Tools\（不進 git，啟動時自動 migrate）
- [x] Watchlist 每間公司獨立輸出路徑（📁 按鈕，存於 watchlist item output_dir）
- [x] Excel 自動美化（深藍色 header、交替底色、section 分隔、subtotal 粗體）
- [x] 財務數字自動 ÷1M（EPS 除外），套用千分位格式
- [x] Index sheet（第一頁，列出所有 sheet 用途 + 最早/最新期間）
- [x] Data_Financials(Q)（季報）+ Data_Financials(Y)（年報 10-K）雙 sheet
- [x] Per-ticker output path memory（ticker_paths in config.json）
- [x] Non-GAAP fetching from 8-K press releases（Data_EPS_Recon + Data_NonGAAP）
- [x] nongaap_cache.json 增量快取（每季 AI 呼叫結果本機快取）
- [x] 單一公司 GAAP 財報抓取
- [x] Excel 輸出（Data_Financials 三表合一 + Data_Seg_* + Data_Meta）
- [x] 批量更新 (Watchlist)
- [x] Watchlist 管理 popup
- [x] 進階設定 popup（AI config, identity, output dir）
- [x] max_filings 設定（Advanced Settings，預設 80 筆 = 約 20 年）
- [x] Ticker 標識（每個 sheet A1）
- [x] IS 固定 22 行模板（含 D&A/SBC/Minority Interest/Total Non-op）
- [x] BS 固定 41 行模板（完整 Assets / Liabilities / Equity）
- [x] CF 固定 25 行模板 + Free Cash Flow 衍生計算
- [x] B 欄 Original Item（公司的 XBRL 原始標籤）
- [x] ProfitLoss fallback（BA、TSLA、XOM、WMT 用 ProfitLoss 報 Net Income）
- [x] D&A label fallback（TSLA std_concept = nan 情況）
- [x] Total Non-op DERIVED fallback（Pre-tax − Operating Income）
- [x] GOOGL encoding fix（非 ASCII 字元 NFKC normalize）
- [x] match="first"|"last" + label_hint 精確比對（解決 BS 重複 std_concept 問題）

### 待辦
- [ ] 實機測試（GAAP）：AAPL、TSLA、BA、XOM 確認 Data_Financials 正確
- [ ] 實機測試（Non-GAAP）：AAPL、NVDA 確認 Data_EPS_Recon + Data_NonGAAP + nongaap_cache.json
- [ ] main.py 舊名稱掃描：確認無 Data_IS/BS/CF 殘留參照
- [ ] 金融股模板（GS/JPM）：UI 自動偵測 + 警告（已設計，延後實作）
- [ ] 批量更新（Tab 2）加入 Non-GAAP 支援

---

## 更新記錄

### 2026-04-17（Session 5）

**Config 搬家 + Watchlist 路徑管理**
- config.json 移到 `%APPDATA%\SEC Financial Tools\config.json`，啟動時自動 migrate 舊檔
- Watchlist 管理介面每行新增 📁 按鈕，可為每間公司設定獨立輸出資料夾
- 路徑存於 watchlist item `output_dir` 欄位，優先順序：watchlist `output_dir` → `ticker_paths`（向後相容）→ 全域 `output_dir`

---

### 2026-04-17（Session 4）

**Excel 自動美化（Phase 3）**
- 新增 `excel_formatter.py`：format_workbook() 在存檔前自動套用所有格式
- 欄寬修正（A=22, B=24, 資料欄=13）：解決科學記號顯示問題
- 深藍色 header 列（Row 1/2）、藍色 section header、灰色分隔列、交替底色、subtotal 粗體
- 財務數字自動 ÷1M，套用 `#,##0.0` 千分位格式；EPS 保留原值用 2 位小數；Shares ÷1M 整數
- Index sheet 自動插入第一頁：列出所有 Data_* sheet 用途、最早/最新期間
- 凍結窗格 C3（Rows 1–2 + Cols A–B 固定）
- 新增 Data_Financials(Y)（年報 10-K），原 Data_Financials 更名為 Data_Financials(Q）
- 新增 72 個 unit tests，全數通過（總計 106 tests）

---

### 2026-04-17（Session 3）

**Per-Ticker Output Path Memory**
- config.json 新增 ticker_paths 欄位
- 確認公司後自動帶出已記憶路徑
- 瀏覽選資料夾後自動儲存至 ticker_paths

**Non-GAAP Fetching（Phase 2）**
- fetcher_nongaap.py 完整實作
- 8-K Item 2.02 篩選，EPS reconciliation（edgartools 原生）
- AI 從 EX-99.1 press release 提取 Non-GAAP 指標（Google / OpenAI / Anthropic）
- nongaap_cache.json 增量快取，只對新季度呼叫 AI
- 輸出：Data_EPS_Recon + Data_NonGAAP sheet

---

### 2026-04-15（Session 2）

**萬能模板實作**
- IS_TEMPLATE 從 18 行擴展至 22 行（新增 D&A、SBC、Minority Interest、Total Non-op）
- 新增 BS_TEMPLATE（41 行）、CF_TEMPLATE（25 行 + FCF 衍生）
- 模板 tuple 從 4-tuple 升級為 6-tuple，加入 `match` 和 `label_hint` 欄位
- `_match_is_row` 新增第三層 label fallback（解決 TSLA D&A nan 問題）
- 三表合一：IS + BS + CF 合併輸出為單一 `Data_Financials` sheet，section header 分隔
- `StatementTable` 新增 `labels: list[str]` 欄位（B 欄 Original Item）
- `excel_writer.py` 改為 A=Std Name / B=Original Item / C+=季度數據
- Post-processing fallbacks：ProfitLoss（Net Income）、DERIVED（Total Non-op）、label "depreciation"（D&A）
- GOOGL encoding fix：`unicodedata.normalize("NFKC")` 處理非 ASCII 標籤
- 新增 53 個 unit tests，全數通過

**GUI 設定**
- Advanced Settings 加入 max_filings 調整（Spinbox，from=4 to=320，預設 80）

### 2026-04-13（Session 1）

- 完成 Phase 1：GAAP fetcher + Excel writer + 完整 Tkinter GUI
- 初始化專案
