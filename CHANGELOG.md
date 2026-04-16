# Changelog

## 現狀

- Phase 1 (GAAP)：萬能模板完成 ✅
- Phase 2 (Non-GAAP)：完成 ✅

## 功能清單

### 已完成
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
- [ ] 實機測試：對真實公司（AAPL、TSLA、BA、XOM）跑一次驗證模板
- [ ] main.py 確認：輸出改成 Data_Financials，需確認 GUI 無誤
- [ ] 金融股模板（GS/JPM）：UI 自動偵測 + 警告（已設計，延後實作）
- [ ] Excel Template 著色功能：使用者自訂顏色 template.xlsx，工具只填值不改格式

---

## 更新記錄

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
