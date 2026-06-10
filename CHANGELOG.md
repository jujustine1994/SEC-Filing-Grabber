# Changelog

## 現狀

- Phase 1 (GAAP)：萬能模板完成 ✅
- Phase 2 (Non-GAAP)：完成 ✅
- Phase 3 (Excel 美化)：完成 ✅

## 功能清單

### 已完成
- [x] B1 Overflow Rows：IS/BS/CF 三表各自追加未被模板消耗的 XBRL rows，確保不遺漏任何財務數字
- [x] Non-GAAP overflow 自動分流至 `Data_Financials_NG(Q/Y)` 獨立 sheet（label 含 "adjusted"/"non-gaap"/"excluding" 等）
- [x] Auto-repair override engine（新 ticker 首次 fetch 自動診斷並修復缺失 key rows）
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
- [x] Gross Profit DERIVED fallback（Revenue − COGS，修復 COHR）
- [x] GOOGL encoding fix（非 ASCII 字元 NFKC normalize）
- [x] match="first"|"last" + label_hint 精確比對（解決 BS 重複 std_concept 問題）
- [x] `_match_is_row` 串接優先級（cascading）：label_hint 不符時跳至下一優先級而非誤匹配
- [x] CF Net Income fallback regex OR（`NetIncomeLoss|ProfitLoss`）
- [x] CF 彙總行 label_hint 改為 `^net cash|^cash`（同時支援 AAPL 與 XOM 排除）
- [x] 實機測試（GAAP）：AAPL ✅ TSLA ✅ BA ✅ XOM ✅ NVDA ✅ COHR ✅（2026-04-23）

### 待辦
- [x] 實機測試（Non-GAAP）：NVDA ✅（正規化後每季 6 指標整齊；修前 34 個稀疏 row）；AAPL ⚠️ 少報屬預期行為（2026-04-19）
- [x] main.py 舊名稱掃描：確認無 Data_IS/BS/CF 殘留參照 ✅（2026-04-19）
- [x] Non-GAAP：nongaap_cache.json ticker 隔離（多公司共用 output_dir 安全）✅（2026-04-25）
- [x] CF：Q2/Q3 YTD overflow 跨 filing 減法 — overflow 行與模板行同邏輯 ✅（2026-04-26）
- [ ] Non-GAAP：Data_EPS_Recon 從未產生（edgartools eps_reconciliation 對 AAPL/NVDA 均回傳空）
- [x] Non-GAAP：NVDA 指標名稱正規化（剝除期間 token、去重比較期間、過濾展望指標）✅（2026-04-26）
- [x] CF：Q2/Q3 YTD→季度換算（Q2 = Q2_YTD − Q1；Q3 = Q3_YTD − Q2_YTD）✅（2026-04-23）
- [x] 金融股警告（GS/JPM 等）：fetch 完成後 log 顯示模板限制警告 ✅（2026-04-26）
- [x] 批量更新（Tab 2）加入 Non-GAAP 支援：新增 checkbox + `_worker_batch` 支援 Non-GAAP ✅（2026-04-26）

---

## 更新記錄

### 2026-06-10
- 修正：`winget install Python` 加入 `--override "/quiet PrependPath=1 Include_pip=1"`，確保靜默安裝後 Python 自動加進 PATH

### 2026-05-05（Session 18）

**Index Sheet 品質檢測**

設計文件：`docs/superpowers/specs/2026-05-03-index-quality-check-design.md`

- **`excel_formatter.py`**：
  - 新增 4 個顏色常數：`QUALITY_GREEN / ORANGE / MISS_BG / MISS_FG`
  - 新增 `ALL_KEY_ROWS`（9 個 key rows：IS 4 + BS 3 + CF 2）
  - 新增 `_compute_quality(tables)` helper：找 `Data_Financials(Q)`，呼叫 `check_key_rows` for IS/BS/CF，回傳 `(score, total, missing_set)` 或 `None`
  - `_build_index_sheet()` 更新：
    - Header 合併範圍 `A1:D1` → `A1:E1`（同 A2）
    - Row 4 新增第 5 欄「完成度」
    - 每個 sheet 列的 E 欄：`Data_Financials(Q)` 顯示 `"9/9 ✓"`（綠）或 `"N/9 ⚠"`（橘），其餘顯示「—」
    - 表格下方加「品質明細」區塊：section header + 9 行逐一顯示 ✓ / ✗，缺失行底色淺橘
- **`tests/test_excel_formatter.py`**：新增 14 個 tests（`_compute_quality` × 4、E 欄 × 5、明細區塊 × 5）

**全套測試：250/250 PASSED**

**Commits**

```
ecddafd  refactor: clean up unused unpacking and duplicate import
f9aab47  refactor: use QUALITY_GREEN constant instead of hardcoded colour in detail loop
da128bb  feat: add quality detail section to Index sheet
b5714d0  feat: add quality score column E to Index sheet
b82de46  refactor: move override_engine import to top-level; import ALL_KEY_ROWS in tests
ee19aa1  feat: add _compute_quality helper for Index sheet quality check
```

---

### 2026-05-03（Session 17）

**日期區間反轉驗證**

- **`main.py`**：
  - `_run_single()`：解析 `start_year` / `end_year` 後新增 guard，`start_year > end_year` 時彈錯誤訊息並中止，訊息帶實際數值（如「起始年份（2025）不可大於結束年份（2020）」）
  - `_run_batch()`：同上

**全套測試：236/236 PASSED**

**Commits**

```
385a67c  fix: guard against start_year > end_year in Tab 1 and Tab 2
```

---

### 2026-04-29（Session 16）

**UI 日期區間 + 報表類型 + Sheet 預覽 + FY Month Fix**

設計文件：`docs/superpowers/specs/2026-04-29-ui-date-range-design.md`

**後端（`fetcher_gaap.py`）**

- **`_filter_filings_by_year()`**（新 helper）：依起訖年份過濾 filing 列表，支援 `date` 物件與 ISO 字串兩種格式
- **`fetch_gaap_statements()` 新增 5 個參數**：
  - `start_year` / `end_year`：只抓指定年份區間的 filings（`None` = 全部，現有行為不變）
  - `fetch_quarterly` / `fetch_annual`：控制是否抓 10-Q / 10-K（預設兩者皆抓）
  - `excluded_sheets`：跳過指定 sheet 名稱，不寫入 Excel
- **`preview_sheets(ticker, identity)`**（新公開函式）：只抓最新一份 10-Q，回傳預期 sheet 名稱清單（~5–15 秒），用於執行前快速預覽
- **FY month fix**：當 `fetch_annual=False` 時，後端自動偷抓 1 份 10-K 僅用於偵測財年結束月份，不產生 `Data_Financials(Y)`；確保 AAPL（9 月）等非 12 月財年公司的季報欄位標籤正確

**後端（`fetcher_nongaap.py`）**

- **`_filter_nongaap_by_year()`**（新 helper）：依 label 字串中的年份（如 `FY2021Q2`）過濾
- **`fetch_nongaap_statements()` 新增 `start_year` / `end_year` 參數**

**UI（`main.py`）**

- **Tab 1（單一公司）**：
  - 「快速掃描 ▶」按鈕（Row 0）：在執行前偵測該 ticker 有哪些 `Data_Seg_*` sheets；固定 sheet（Financials/Meta）不可取消
  - 「▶ 進階設定」收折區（預設隱藏）：展開後可個別勾選季報（10-Q）/ 年報（10-K）
  - 日期區間 Spinbox（起 / 迄年份，留空 = 全部）
- **Tab 2（批量更新）**：
  - 「▶ 進階設定」收折區：同上，可選報表類型
  - 日期區間 Spinbox

**測試（`tests/`）**

- `test_fetcher_gaap.py`：新增 14 個測試（filter、fetch_gaap 新參數、preview_sheets、FY probe）→ 共 **110 tests**
- `test_fetcher_nongaap.py`：新增 4 個 filter 測試 → 共 **41 tests**
- 全套 **151/151 PASSED**

**Commits（2026-04-29）**

```
555193d  feat: move report type to inline adv settings; fix FY month probe for Q-only mode
adf2336  fix: guard scan against running fetch; hide panel on empty sheet list
5dd7cdd  feat: add quick scan button and sheet selection panel to Tab 1
63cd654  fix: align Tab 2 row sticky to ew matching Tab 1 pattern
bbb7d43  feat: add report type and date range to Tab 2 batch UI
d15f1d9  feat: add report type (10-Q/10-K) and date range to Tab 1 UI
ab9d697  feat: add start_year/end_year filter to fetch_nongaap_statements
a9caeec  feat: add preview_sheets() for quick segment detection
016845b  feat: add start_year/end_year/fetch_quarterly/fetch_annual/excluded_sheets to fetch_gaap_statements
9260596  feat: add _filter_filings_by_year helper with year-range support
```

---

### 2026-04-28（Session 15）

**GAAP Fetcher 精度修復（Tasks 1–8）+ Batch Smoke Test**

設計文件：`docs/superpowers/plans/2026-04-26-fetcher-gaap-fixes.md`

- **Task 1：pre-XBRL early exit** — IS/BS/CF/Seg 四個 filing loop 在 `filing_date < 2008-01-01` 時 `break`，AMD 等老公司抓取速度大幅改善
- **Task 2：Dividends Paid bug** — CF_TEMPLATE Dividends Paid 的 `std_concept` 從 `DistributionsToMinorityInterests`（錯誤）改為 `None`，fallback regex 正確命中
- **Task 3：Net Income fallback 順序** — IS 先嘗試 `NetIncomeLossAttributableToParent`，再 `ProfitLoss`，避免少數股東損益被誤計入母公司 Net Income
- **Task 4：Revenue fallback 擴展** — 新增 `Revenues`、`SalesRevenueNet`、`SalesRevenueGoodsNet`，涵蓋更多公司的 XBRL 命名
- **Task 5：Total Non-op DERIVED guard** — 當 discontinued operations 行存在時跳過衍生計算，避免雙重計入
- **Task 6：Investment Proceeds 多概念加總** — 新增 `_sum_matching_rows()` helper，將 `ProceedsFromSaleOfInvestments`、`ProceedsFromMaturitiesOfInvestments` 等多個概念加總，而非只取第一筆
- **Task 7：Debt Proceeds/Repayments 多概念加總** — 同上，LT 債 + ST 債分別加總，不再只抓第一個命中的概念
- **Task 8：FY Label 對齊公司財年** — 新增 `_detect_fy_end_month(filings_k)` 從 10-K 推算財年結束月份；`_col_to_quarter_label()` 根據 `fy_end_month` 調整 Q 序號，確保 AAPL（Sep）、MSFT（Jun）、NVDA/WMT（Jan）的季度標籤正確

**`tests/test_fetcher_gaap.py`**：新增對應 Tasks 1–8 的 unit tests；**95/95 全數通過**

---

**Batch Live Smoke Test（新腳本）**

- **`scripts/smoke_test_10.py`（新檔）**：
  - 對 10 間公司（AAPL/MSFT/TSLA/AMD/NVDA/GOOGL/META/WMT/COHR/AMZN）呼叫真實 EDGAR API
  - 自動檢查 7 個 key rows：Revenue / Gross Profit / Operating Income / Net Income / Operating Cash Flow / Capex / Free Cash Flow
  - 輸出每間公司的 OK/NONE 狀態 + 末尾彙總表，列出有問題的 ticker
  - Windows cp950 encoding 修復：`sys.stdout = io.TextIOWrapper(..., encoding="utf-8")`

- **`scripts/README.md`（新檔）**：腳本索引，符合 CLAUDE.md 規則（新增腳本必須更新此表）

**實機測試結果（2026-04-28）：9/10 OK**
- AAPL/MSFT/TSLA/AMD/NVDA/GOOGL/META/WMT/AMZN：7 個 key rows 全部 ✅
- COHR：Revenue=NONE、Operating Income=NONE — 已知限制（2022 三方合併導致非標準 XBRL）
- FY label 驗證：AAPL FY2026Q1 ✅ / MSFT FY2026Q2 ✅ / NVDA FY2026Q3 ✅ / WMT FY2026Q3 ✅

**所有 commits 已 push 至 GitHub remote**（`0508dc6`）

---

### 2026-04-26（Session 14）

**CF YTD Overflow 修復 + 測試套件（COHR/LITE/AAPL/NVDA/GOOGL）**

- **`fetcher_gaap.py`**：
  - `_build_cf_table`：移除 `if not is_ytd: _collect_overflow(...)` 限制
  - 新增 `overflow_per_filing: dict[str, dict]` — 對所有 filing（含 YTD）收集原始 overflow 值
  - Filing loop 結束後，與模板行相同的跨 filing 減法計算出 Q2/Q3 standalone overflow
  - Q2 overflow standalone = Q2_YTD_overflow − Q1_overflow；Q3 = Q3_YTD − Q2_YTD
  - 若前一季無對應 concept，Q2/Q3 overflow 保持 None（保守策略，與模板行一致）

- **`tests/test_fetcher_gaap.py`**：
  - 新增 4 個 CF overflow YTD unit tests：
    - `test_cf_overflow_q1_standalone_captured` — Q1 overflow 正常收集
    - `test_cf_overflow_q2_ytd_subtracted` — Q2 overflow = Q2_YTD − Q1
    - `test_cf_overflow_q3_ytd_subtracted` — Q3 overflow = Q3_YTD − Q2_YTD
    - `test_cf_overflow_q2_without_q1_is_none` — 前一季缺失時結果為 None
  - **73/73 unit tests 全數通過**

- **`tests/test_live_snapshots.py`**：
  - 新增 `CF_OVERFLOW_TICKERS = ["COHR", "LITE", "AAPL", "NVDA", "GOOGL"]`
  - 新增 `cf_overflow_tables` fixture（module scope，同 `all_tables`）
  - 3 個 `@pytest.mark.cf_overflow` live tests：
    - `test_cf_overflow_rows_exist` — COHR/LITE 至少有 1 個 CF overflow row
    - `test_cf_overflow_multi_quarter_coverage` — 至少 1 個 overflow row 有 ≥2 季數據（驗證 YTD 減法有效）
    - `test_cf_overflow_no_all_none_rows` — 無全 None 的 overflow row
  - 執行：`pytest -m "slow and cf_overflow"` 實測約 5 分鐘

- **`conftest.py`**：新增 `cf_overflow` marker 定義

**Live 驗證結果（2026-04-26）：15/15 PASSED**
- COHR ✅ LITE ✅ AAPL ✅ NVDA ✅ GOOGL ✅（5:05 分鐘）
- 三項驗證全過：overflow rows 存在、≥2 季有數據、無全 None rows

---

**Non-GAAP 指標名稱正規化（NVDA 稀疏表格修復）**

問題：NVDA press release 的表格包含多個比較期間，AI 回傳如：
- `"Non-GAAP Q4 FY26 Gross margin"` / `"Non-GAAP Q3 FY26 Gross margin"` / `"Non-GAAP Q4 FY25 Gross margin"` — 同一指標三個版本
- `"Non-GAAP FY2026 Gross margin"` — 全年版（重複）
- `"Expected Non-GAAP Gross margin (Q1 FY27)"` — 展望指標（不應存）
- 修前 Q4 FY26 filing 回傳 34 個 metrics；修後 6 個

- **`fetcher_nongaap.py`**：
  - 新增 `_clean_metric_name(name)` — 用 `_PERIOD_TOKEN_RE` 移除所有期間 token（`Q4 FY26`、`FY2026`），再移除空括號和噪音標籤
  - 新增 `_normalize_nongaap_metrics(raw)` — 兩段式去重：quarterly token（或無 token）優先，FY-only token 補缺；同類別內第一次出現的值優先（press release 以最新期為首，比較期自動被丟棄）
  - `_call_ai()` 回傳前呼叫 `_normalize_nongaap_metrics(result)`

- **`tests/test_fetcher_nongaap.py`**：新增 12 個正規化 unit tests（37/37 通過）

**Unit tests：110/110 通過**

---

**金融股警告 + Tab 2 Non-GAAP 批量支援**

- **`main.py`**：
  - 新增 `_FINANCIAL_SECTOR_TICKERS = frozenset({"GS", "JPM", "BAC", "C", "WFC", "MS", "BLK", "BX", "KKR"})`
  - `_worker_single` / `_worker_batch`：fetch 完成後若 ticker 在金融股集合內，log 警告「BS/IS 部分欄位可能為空」
  - Tab 2 新增「同時抓取 Non-GAAP」Checkbutton + API Key 警告 label
  - `_on_batch_nongaap_toggle()`：切換 checkbox 時顯示/隱藏 API Key 警告
  - `_run_batch()`：讀取 Non-GAAP checkbox 狀態，未設 API Key 時 error dialog
  - `_worker_batch(fetch_nongaap: bool)`：若啟用，對每個 ticker 呼叫 `fetch_nongaap_statements` 並合併 tables

**Data_EPS_Recon 說明**：edgartools `eps_reconciliation` 對 NVDA/AAPL/MSFT 均回傳 None（非 XBRL tagged），無法從 edgartools 取得。此功能保留為已知限制，待 edgartools 改善或改用 AI 解析方案。
- COHR ✅ LITE ✅ AAPL ✅ NVDA ✅ GOOGL ✅（5:05 分鐘）
- 三項驗證全過：overflow rows 存在、≥2 季有數據、無全 None rows

---

### 2026-04-25（Session 13）

**B1 Overflow Rows — 確保所有 XBRL 財務數字都不遺漏**

設計文件：`docs/superpowers/specs/2026-04-25-overflow-rows-design.md`

核心概念：每次 filing 建表時追蹤模板已消耗的 XBRL row indices（`consumed: set[int]`），未被消耗的 rows 以 overflow 形式追加在模板行之後。

- **`fetcher_gaap.py`**：
  - 新增 `_NONGAAP_KEYWORDS: frozenset` + `_is_nongaap_label(label)` — keyword-based 分類，label 含 "adjusted"/"non-gaap"/"excluding"/"excl."/"ex-" 的視為 Non-GAAP
  - 新增 `_collect_overflow(df, consumed, data_col, quarter_label, gaap_out, ng_out)` — 從未消耗的 XBRL rows 中收集數值，分流至 GAAP / NG 兩個 output dict；跳過 abstract / breakdown / dimension rows（沿用 `_consolidated_mask`）
  - `_build_is_table` → 回傳 `tuple[StatementTable, StatementTable]`：
    - IS df 的 consumed 追蹤；CF-source rows 只追蹤 IS df consumed（CF overflow 由 `_build_cf_table` 負責）
    - `gaap_tbl`（`Data_IS`）= 22 個模板行 + GAAP overflow 行；`ng_tbl`（`Data_IS_NG`）= NG overflow 行
  - `_build_bs_table` → 回傳 `tuple[StatementTable, StatementTable]`（`Data_BS` / `Data_BS_NG`）
  - `_build_cf_table` → 回傳 `tuple[StatementTable, StatementTable]`（`Data_CF` / `Data_CF_NG`）；只對非 YTD filings（Q1/FY）收集 overflow，避免需要跨 filing 減法
  - `fetch_gaap_statements` → 所有 6 次 build 呼叫改用 tuple unpacking；若任一段有 NG overflow rows，則建 `Data_Financials_NG(Q)` / `Data_Financials_NG(Y)` 並加入輸出清單

- **`tests/test_fetcher_gaap.py`**：
  - 新增 16 個 helper tests（`_is_nongaap_label` × 8，`_collect_overflow` × 8）
  - 新增 3 個 smoke tests（`_build_is/bs/cf_table` 空 filings 回傳 tuple）
  - 所有現有 `_build_is_table` / `_build_cf_table` tests 改為 `gaap_tbl, _ = ...` 解包

已知限制（待辦，不在本次範圍）：CF Q2/Q3 YTD overflow 需要跨 filing 減法，目前跳過；地雷十二仍存在。

**69/69 unit tests 全數通過**

---

### 2026-04-24（Session 11）

**Auto-repair Override Engine**

新增 `override_engine.py`，fetch 完成後自動偵測並修復關鍵欄位缺失，無需人工介入。

設計文件：`docs/superpowers/specs/2026-04-23-auto-repair-design.md`

- **`override_engine.py`（新檔）**：
  - `check_key_rows()`：檢查 9 個 key rows（Revenue / Operating Income / Net Income / EPS Diluted / Total Assets / Total Liabilities / Total Equity / OCF / Capex）最近 4 期是否全為 None
  - `e1_fuzzy_match()`：Rule-based，用 `SYNONYM_MAP` 對 EDGAR DataFrame 做子字串比對（免 API 費用）
  - `e2_llm_diagnose()`：E1 失敗時呼叫現有 AI API，提供概念清單給 LLM 識別正確 std_concept 或確認為 structural_absence
  - `load_overrides()` / `save_overrides()`：永久記錄診斷結果至 `%APPDATA%/SEC_Financial_Tools/ticker_overrides.json`，下次同 ticker 不重診斷
  - `run_diagnosis()`：串接 E1 → E2，診斷完自動存檔

- **`fetcher_gaap.py` 整合**：
  - `_apply_row_override()`：新 helper，按 override 的 fix_type（concept_override / structural_absence）從 DataFrame 取值
  - `_build_is_table` / `_build_bs_table` / `_build_cf_table`：各加 `*_overrides` 參數，filing loop 開頭先套用 override，不再呼叫 `_match_is_row`
  - `fetch_gaap_statements()`：加 `ai_config` 參數；三表建完後自動 check_key_rows → run_diagnosis；有新 override 時重建三表（當次 fetch 直接出正確結果）

- **`main.py`**：兩處 `fetch_gaap_statements` 呼叫改傳 `ai_config=self.cfg.get("ai", {})`

**Bug 修復：E2 LLM response 解析**

- 舊邏輯 `response.upper() == "ABSENT"` 只能抓完全等於 "ABSENT" 的回應
- 新邏輯：`re.search(r'\bABSENT\b', ..., re.IGNORECASE)` 處理 ABSENT 嵌在句子中的情況
- 新增垃圾回應防護：含空格或長度 >100 → 回傳 None，不存入 override

新增 **28 個 unit tests**（`tests/test_override_engine.py`）+ 4 個 override 整合測試（`tests/test_fetcher_gaap.py`）；**總計 151 tests，全數通過**

---

### 2026-04-24（Session 12）

**Live Snapshot Tests（自動化實機驗證）**

新增 `tests/test_live_snapshots.py`，對真實 EDGAR API 抓資料，斷言 key rows 不全為 None。

- **`tests/test_live_snapshots.py`（新檔）**：
  - 24 個 `@pytest.mark.slow` tests（8 tickers × IS/BS/CF）
  - Tickers：MSFT、AMZN、META、GOOGL、NVDA、JPM、GS、JNJ
  - `all_tables` fixture（module-scoped）：8 間公司只抓一次（max_filings=8）
  - 金融股（GS/JPM）IS 允許 Operating Income 缺失；JPM CF 允許 Capex 缺失
  - 實跑結果：**24/24 PASSED**（總耗時約 36 分鐘）

- **`conftest.py`（新檔）**：註冊 `slow` marker，避免 PytestUnknownMarkWarning

- **`override_engine.py` bug 修正**：
  - KEY_ROWS 三個命名錯誤（與 StatementTable concept 名稱不符）
  - `"EPS Diluted"` → `"Diluted EPS"`；`"Total Equity"` → `"Total Equity — Parent"`；`"Capital Expenditures"` → `"Capex"`
  - 同步更新 SYNONYM_MAP key 名稱
  - 影響：原本 3/9 key rows 的 override 自動診斷為無效 no-op，現已修復

- **`tests/test_override_engine.py`**：更新 IS_CONCEPTS mock 及 e1_fuzzy_match 測試名稱

- **`PITFALLS.md` 地雷十六**：JPM CF Capex structural absence（銀行類公司使用不同 XBRL 概念名）

執行方式：`pytest -m slow`（~36 分鐘）、`pytest -m "not slow"`（~9 秒）

**總計 175 tests（151 unit + 24 live），全數通過**

---

### 2026-04-23（Session 10）

**CF YTD→季度換算**
- `fetcher_gaap.py`：新增 `_ytd_col()` — 偵測 `(YTD)` 欄位（edgartools 對 Q2/Q3 CF 的標記格式）
- `fetcher_gaap.py`：新增 `_prev_quarter_label()` — 將 "FY2025Q2" 對應到 "FY2025Q1"
- `fetcher_gaap.py`：改寫 `_build_cf_table()` — 不再呼叫 generic `_build_template_table`；改為：
  - Q1/FY：直接讀 standalone 欄位（行為不變）
  - Q2/Q3：讀 YTD 欄位 + 借 IS 取 quarter label（同 BS 做法）；儲存原始 YTD 值，最後做 `Q2 = Q2_YTD − Q1`、`Q3 = Q3_YTD − Q2_YTD` 換算
  - 無前期資料時保留 YTD 原值（best-effort）
- `fetcher_gaap.py`：在 `_match_is_row` 上方加 TODO — 提醒驗證 fallback 涵蓋 ≥10 間公司（MSFT/AMZN/META/GOOGL/NVDA/JPM/GS/JNJ 待測）
- 新增 13 個 unit/integration tests（總計 119 tests，全數通過）

---

### 2026-04-23（Session 10）

**核心修復：`_match_is_row` 串接優先級邏輯（Cascading Priority）**
- 舊邏輯：`label_hint` 找不到符合行時，仍從未過濾的候選行取 first/last（可能拿到錯誤項目）
- 新邏輯：`label_hint` 不符合 → 傳回 None → 呼叫端跳至下一優先級（std_concept → fallback_suffix → label_fallback）
- 避免「hint 失效後偷跑到錯誤行」的問題

**IS 修復**
- `Cost of Revenue`：新增 `label_hint="cost"`，避免 XOM "Sales-based taxes" 誤匹配（其 std_concept 也是 CostOfRevenue）
- `Operating Income`：fallback_suffix 改為 `"OperatingIncomeLoss"`，避免比對到 `us-gaap_OtherOperatingIncomeExpenseNet`（舊 "OperatingIncome" 是後者的子字串）；同時移除已破損的 `label_hint="operation"`（"Operating income" 含 "operat" 不含 "operation"）
- `Gross Profit` 新增 DERIVED fallback：若無直接匹配，從 `Revenue − COGS` 衍生（修復 COHR）

**CF 修復**
- `Net Income`：fallback_suffix 改為 `"NetIncomeLoss|ProfitLoss"`（regex OR），修復 XOM CF Net Income = None（XOM 用 ProfitLoss 而非 NetIncome）
- `Change in Receivables`：新增 `label_hint="receivable"`，搭配 cascading 修復 COHR（ChangeInReceivables std_concept 誤貼在 "Income taxes" 行）
- `Cash Taxes Paid`：label_hint 從 `"income tax"` 改為 `"paid"`，排除 "Deferred income taxes"（不含 "paid"）
- `Cash Interest Paid`：label_hint 從 `"interest paid"` 改為 `"paid"`，修復 COHR "Cash paid for interest"（詞序不符）
- `Operating/Investing/Financing Cash Flow`：label_hint 從 `"net cash"` 改為 `"^net cash|^cash"`（正則 starts-with），修復 AAPL "Cash generated by operating activities" 同時繼續排除 XOM "Noncash right of use assets..."

**實機測試結果（AAPL / TSLA / XOM / NVDA / COHR）**
- AAPL：IS 正確 ✅；CF OCF/ICF/FCF Q1 正確（FY2025Q1 OCF=$53.9B）✅
- TSLA：IS/CF 正常 ✅
- XOM：Revenue 正確；Gross Profit / Operating Income = None（XOM 不報此兩行，預期行為）✅；CF OCF FY2025Q1=$13.0B ✅
- NVDA：IS/CF 正常 ✅
- COHR：Gross Profit DERIVED 正確 ✅；Operating Income = None（COHR 無此獨立行，正確）✅

---

### 2026-04-18（Session 9）

**Bug 修復（實機測試發現）**
- `fetcher_gaap.py`：CF 三大彙總行（Operating/Investing/Financing Cash Flow）新增 `label_hint="net cash"`，避免 `match="last"` 因相同 `standard_concept` 拿到 trailing noncash 項目（如 XOM 的 ROU lease 調整項）
- `fetcher_gaap.py`：FCF 計算改為 `OCF − abs(Capex)`，修正 XOM 等以負數回報 Capex 的公司 FCF 加倍的問題

**實機測試結果**
- TSLA、BA：三表輸出正常，IS/BS 數值與公開財報吻合 ✅
- XOM：OCF 從 $6M（誤）修正為 $12,953M，FCF 從 $5,904M（偶然正確）修正為 $7,055M ✅
- 已知限制：Q2/Q3 CF 全為 None（XOM/TSLA/BA 均以 YTD 格式回報，`_current_q_col` 跳過）

---

### 2026-04-17（Session 8）

**Bug 修復**
- `fetcher_nongaap.py`：`fetch_nongaap_statements` 新增 `max_filings` 參數（預設 80），首次抓取不再無限往回到 2004 年
- `main.py`：呼叫 `fetch_nongaap_statements` 時傳入 `max_filings`，與 GAAP 上限保持一致

**程式碼文件化**
- `main.py`：補上 SECFetcherApp class docstring（說明 thread/queue 架構）及 27 個 method docstring（涵蓋所有原本空白的方法）
- 重點標注非顯而易見的行為：`_wl_draft` 暫存模式、cache-first 查詢邏輯、Tkinter thread 安全機制（msg_queue + _poll_queue）、double-run 防護等
- 其他五個檔案（config.py、excel_formatter.py、excel_writer.py、fetcher_gaap.py、fetcher_nongaap.py）原本已有完整 docstring，無需異動

---

### 2026-04-17（Session 7）

**Bug 修復（程式碼審查後）**
- `config.py`：修正 config 值為非 dict 時 `dict.update()` 會 crash 的問題（如舊 config 格式不相容）
- `main.py`：修正 `_wl_add()` 中重複呼叫 `wl_group_var.get()` 的冗餘程式碼
- `fetcher_nongaap.py`：修正 `_call_ai()` 中 AI 回傳非數字字串時 `float()` 整批失敗的問題；改為逐項 try/except

**UI 改善**
- 移除「確認公司」按鈕，Enter 鍵觸發查詢即可
- 輸出設定改為可收合（▼/▶ toggle），預設展開
- 視窗支援縮放（resizable），Log 區域隨視窗高度延伸
- Advanced Settings 新增「預設模板 / 自訂模板」Radio 選擇

**Excel 自訂模板**
- `excel_writer.py` 新增 `template_path` 參數
- 自訂模板模式：保留所有 cell 格式，只寫入數值；超出模板欄數時複製最後一欄格式
- 新增 `_write_sheet_template()` 及 `_copy_cell_format()` helper
- 兩種模式（預設自動著色 / 自訂模板）可在 Advanced Settings 切換

**Watchlist 群組管理**
- 支援股票分群（群組 CRUD：新增、重命名、刪除）
- 刪除群組時自動把 ticker 移至「未分類」
- 群組依字母排序，「未分類」固定最後
- Watchlist popup 改為「儲存關閉 / 放棄關閉」（Escape = 放棄），操作前先建立 deep copy draft

---

### 2026-04-17（Session 6）

**BS 抓取修復**
- `_build_bs_table` 改為獨立實作；BS 欄位為 instant（bare date），`_current_q_col` 無法識別導致全空白
- 修法：從同一 filing 的 IS 借用 quarter label，BS 取第一個非 meta 欄位讀值

**Watchlist popup 改善**
- 「目前 Watchlist」區加入捲軸（固定高度 160px）
- Ticker 輸入框：Enter → 查詢，查到後自動加入（單次 Enter 完成）

**Tab 2 批量更新改善**
- Watchlist 改為捲軸顯示（固定高度 150px），只顯示 ticker 代號，3 個一列

---

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
