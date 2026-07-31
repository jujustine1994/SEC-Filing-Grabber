# SEC Financial Fetcher — Architecture

## File Map

| File | Role |
|------|------|
| 啟動器.bat | 薄 BAT，呼叫 launcher.ps1 |
| launcher.ps1 | 環境檢查、uv venv、安裝套件、啟動 main.py |
| main.py | Tkinter GUI，兩個 tab + 兩個 popup |
| config.py | load_config() / save_config() |
| fetcher_gaap.py | edgartools XBRL 抓取 → StatementTable 列表 |
| fetcher_nongaap.py | 8-K press release 抓取 → EPS Recon + Non-GAAP StatementTable |
| excel_writer.py | 寫 Data_* sheets 至 output/TICKER.xlsx，並呼叫 excel_formatter |
| excel_formatter.py | 寫 Index sheet（品質明細）、設欄寬、凍結窗格 |
| override_engine.py | 自動修復缺失 key rows（E1 fuzzy + E2 LLM） |
| config.json | 使用者設定（gitignored） |
| config.example.json | 範本（committed） |
| company_cache.json | Ticker → 公司名快取（committed） |
| output/ | 輸出的 Excel 檔（gitignored） |
| nongaap_cache.json | 各公司輸出資料夾內，Non-GAAP 快取（runtime，非 git） |

## Data Flow

```
使用者輸入 Ticker（Tab 1）或從 Watchlist 選取（Tab 2）
    ↓
fetcher_gaap.py
    ├─ _build_is_table()  → (gaap_tbl, ng_tbl)  IS 22-row 模板 + GAAP/NG overflow
    ├─ _build_bs_table()  → (gaap_tbl, ng_tbl)  BS 41-row 模板 + GAAP/NG overflow
    ├─ _build_cf_table()  → (gaap_tbl, ng_tbl)  CF 26-row 模板 + GAAP/NG overflow
    │
    ├─ override_engine.check_key_rows()  → 找全 None 的 key rows
    ├─ override_engine.run_diagnosis()   → E1 fuzzy + E2 LLM → save_overrides()
    │   （有新 override 時重跑三個 build 函式）
    │
    ├─ _merge_financials(is, bs, cf)    → Data_Financials(Q)    ← 主輸出
    ├─ _merge_financials(ng_is, ng_bs, ng_cf) → Data_Financials_NG(Q)  ← 有 NG overflow 時
    ├─ _merge_financials(annual...)     → Data_Financials(Y)    ← 10-K
    ├─ _merge_financials(ng annual...)  → Data_Financials_NG(Y) ← 有 NG overflow 時
    ├─ _build_segment_tables()          → Data_Seg_* (多個)
    └─ _build_meta_table()              → Data_Meta

fetcher_nongaap.py（勾選 Non-GAAP 時，完全獨立於 GAAP fetcher）
    ├─ _list_earnings_filings()   → 在 SEC 申報清單階段以 items（2.02）+ period_of_report 篩選、去重、套年份與 max_filings，零下載
    ├─ _find_missing_quarters()   → 偵測季度序列缺口
    ├─ _recover_missing_quarters()→ 只對缺季區間逐筆 obj() 深掃，用 has_earnings 找回未標 2.02 的財報
    ├─（以上皆不下載；filing.obj() 移進 fetch_nongaap_statements() 的逐季迴圈，
    │   只對 nongaap_cache.json 尚未收錄的季度下載，且包在 try 內——下載失敗只損失該季）
    ├─ _extract_eps_recon()       → edgartools eps_reconciliation
    ├─ _extract_nongaap_metrics() → AI 解析 EX-99.1 press release
    │       └─ _normalize_nongaap_metrics() → 剝除期間 token、去重比較期間、過濾展望指標
    ├─ _build_eps_recon_table()   → Data_EPS_Recon
    └─ _build_nongaap_table()     → Data_NonGAAP
    ↓（中間結果寫入 nongaap_cache.json，增量更新）
    ↓
excel_writer.py
    → 全量改寫所有 Data_* sheets，不碰 My_* 等其他 sheets
    → format_workbook() (excel_formatter.py)
        ├─ _build_index_sheet()  → Index sheet（sheet 清單 + 完成度欄 + 品質明細區塊）
        ├─ _apply_column_widths()
        └─ _set_freeze_panes()
    → output/TICKER.xlsx（或 ticker_paths[ticker] 指定路徑）
```

## Key Config Variables (config.json)

| 鍵 | 說明 |
|----|------|
| `identity` | SEC EDGAR 身份字串（必填，格式：名字 空格 信箱） |
| `output_dir` | Excel 輸出路徑（預設 "output"） |
| `ticker_paths` | `{TICKER: absolute_path}` 各公司輸出資料夾記憶 |
| `max_filings` | 最多抓幾筆 10-Q（預設 80，約 20 年） |
| `watchlist` | [{ticker, name}, ...] 清單 |
| `ai.provider` | "google" / "openai" / "anthropic" |
| `ai.model` | 模型名稱 |
| `ai.api_key` | API Key（gitignored） |

## Excel Sheet Layout

### Data_Financials(Q) / Data_Financials(Y)（主要輸出）

```
A1=ticker  B1=空  C1=FY2024Q1  D1=FY2024Q2  ...
A2=空      B2=空  C2=2024-02-01 D2=2024-05-03 ...
A3=Income Statement  (section header, all values = None)
A4=Revenue  B4=Net sales  C4=100.0  ...   ← 22 IS 固定模板行
...
A25=〔IS overflow 行〕  B25=us-gaap_ConceptXxx   ← 可能 0 到 N 行
A26=空  (blank separator)
A27=Balance Sheet  (section header)
A28=Cash  B28=Cash and cash equivalents  ...   ← 41 BS 固定模板行
...
A70=〔BS overflow 行〕                          ← 可能 0 到 N 行
A71=空
A72=Cash Flow  (section header)
A73=Net Income  B73=Net income  ...            ← 26 CF 固定模板行（含 FCF DERIVED）
...
A99=〔CF overflow 行〕                          ← 可能 0 到 N 行（僅 Q1/FY filings）
```

- **Col A** = 標準指標名稱（template 行）或 XBRL 原始 label（overflow 行）
- **Col B** = Original Item（XBRL 原始標籤，overflow 行為 concept name）
- **Col C+** = 季度數據（oldest → newest）

### Data_Financials_NG(Q) / Data_Financials_NG(Y)（有 Non-GAAP overflow 時才產生）

格式與 Data_Financials 完全相同，但每個 section 只含 Non-GAAP overflow 行（無固定模板行）。  
Non-GAAP 判斷：label 含 "adjusted"/"non-gaap"/"non gaap"/"excluding"/"excl."/"ex-"。  
**與 Data_EPS_Recon / Data_NonGAAP 完全獨立**（來源不同：這裡是 XBRL，那裡是 8-K press release）。

### Data_Seg_*

每個有 segment breakdown 的 IS 概念一張 sheet，格式同上但沒有 B 欄 labels。

### Data_EPS_Recon

EPS 調和表（GAAP EPS → 調整項 → Non-GAAP EPS）。B 欄為空（無 XBRL labels）。

### Data_NonGAAP

AI 從 8-K press release 提取的所有 Non-GAAP / Adjusted / Excluding 指標。跨季取聯集，缺的季填 None。

> ⚠️ **欄位標籤已知不準（2026-07-31 查證，未修，見 TODO 第 2 項）**：季度標籤由 `_period_to_quarter_label()` 依 8-K 的 `period_of_report` 推算，但該欄在 Item 2.02 財報 8-K 上存的是**發布日**而非財期結束日，故標籤普遍比數字實際所屬財季**晚約一季**（INTC `20260723` 標成 `FY2026Q3`，實報 FY2026 Q2）。同一根因下，同一日曆季內發布兩份財報 8-K 時（如 WDC 2025-01-10 與 2025-01-29）兩者標籤相同，去重「留最舊」會丟掉較新那份。此為長期行為，`Data_Financials` 走 XBRL 不受影響。

## StatementTable（fetcher_gaap.py 的輸出合約）

```python
@dataclass
class StatementTable:
    sheet_name:     str           # "Data_Financials", "Data_Seg_Revenue", ...
    quarter_labels: list[str]     # Row 1, col C+
    filing_dates:   list[str]     # Row 2, col C+
    concepts:       list[str]     # Col A, Row 3+
    values:         list[list]    # values[concept_idx][quarter_idx]
    ticker:         str = ""      # Col A1
    labels:         list[str]     # Col B, Row 3+ (original XBRL labels)
```

## Template Matching Logic（_match_is_row）

3 層查找 + 2 個修飾參數：

```
Priority 1: standard_concept == std_concept
Priority 2: concept 欄位包含 fallback_suffix（case-insensitive）
Priority 3: label 欄位包含 label_fallback（case-insensitive）

label_hint: 在 candidates 中優先選 label 含 hint 的行
match:      "first"（預設）= 最早那行；"last" = 最後那行（用於 CF 彙總行）
```

## Template 行數摘要

| 報表 | 行數 | 格式 |
|------|------|------|
| IS_TEMPLATE | 22 | 6-tuple (label, std_concept, fallback_suffix, source, match, label_hint) |
| BS_TEMPLATE | 41 | 同上 |
| CF_TEMPLATE | 26 | 同上（第 26 行為 Free Cash Flow，DERIVED = OCF − |Capex|） |

source 欄位值：`"IS"` / `"BS"` / `"CF"` = 從哪個 DataFrame 取值；`"DERIVED"` = 不做 XBRL 比對，由 post-processing 計算。

## B1 Overflow Rows（三表 GAAP / NG 分流）

每個 `_build_*_table` 函式現在回傳 `tuple[StatementTable, StatementTable]`：`(gaap_tbl, ng_tbl)`。

**Consumed tracking：**
- 每個 filing 的 template 比對迴圈維護 `consumed: set[int]`
- `_match_is_row()` 找到 index 時 → `consumed.add(idx)`
- IS 的 CF-source rows（D&A 等）消耗 `cf_df` 的 index，不計入 IS `consumed`（由 `_build_cf_table` 各自追蹤）
- ProfitLoss fallback 也計入 IS `consumed`

**Overflow 收集（`_collect_overflow`）：**
```python
_collect_overflow(df, consumed, data_col, quarter_label, gaap_out, ng_out)
```
- 套用 `_consolidated_mask`（排除 abstract / breakdown / dimension rows）
- 跳過 `consumed` 中的 index
- `_is_nongaap_label(label)` 為 True → 進 `ng_out`；否則進 `gaap_out`
- all-None 的 overflow row 最終不追加（build 函式末段過濾）

**CF YTD overflow（2026-04-26 修復）：**
Q2/Q3 overflow 使用與模板行相同的跨 filing 減法：
- Filing loop 內：對所有 filing（含 YTD）收集原始 overflow 值至 `overflow_per_filing[label]`
- Loop 結束後：非 YTD 季 → 直接使用原始值；YTD 季 → `raw[q] - raw[prev_q]`
- 若前一季無對應 concept → 保持 None（與模板行行為一致）
- 驗證：`pytest -m "slow and cf_overflow"` 15/15 PASSED（COHR/LITE/AAPL/NVDA/GOOGL）

## IS Post-processing Fallbacks

在每個 filing 的 row_vals 計算完後執行：

1. **Total Non-op**：若 XBRL None → `Pre-tax Income − Operating Income`
2. **Net Income**：若 NetIncome None → 試 `ProfitLoss`（BA/TSLA/XOM/WMT）
3. **D&A**：若 DepreciationExpense None → label fallback `"depreciation"`（TSLA）

## CF Post-processing

- **Free Cash Flow** = `Operating Cash Flow − Capex`（每季計算）

## Override Engine（自動修復缺失指標）

spec: `docs/superpowers/specs/2026-04-23-auto-repair-design.md`

```
fetch_ticker(ticker)
    ↓ 三表跑完
check_key_rows()       ← 找「最近 4 期全為 None」的 key rows（約 9 個）
    ↓ 有缺失
load_overrides()       ← APPDATA/ticker_overrides.json，已診斷過就直接套用
    ↓ 無 override
E1: fuzzy_match()      ← rule-based，無 API 費用
    ↓ E1 未命中
E2: llm_diagnose()     ← 呼叫現有 AI API（需 api_key）
    ↓
save_overrides()       ← 診斷結果永久寫入，下次同 ticker 不重跑
```

**Override 套用時機**：每個 filing 的 row_vals 計算前（loop 開頭），不是 post-processing。  
**新增檔案**：`override_engine.py`

## 測試分層

| 指令 | 時間 | 測試數 | 用途 |
|------|------|--------|------|
| `python -m pytest tests/ --ignore=tests/test_live_snapshots.py` | ~13 秒 | 250 | Unit tests（每次改 code 後跑） |
| `pytest -m "slow and b1"` | ~12 分鐘 | 24 | B1 overflow live 驗證（8 tickers） |
| `pytest -m "slow and cf_overflow"` | ~5 分鐘 | 15 | CF YTD overflow 驗收（COHR/LITE/AAPL/NVDA/GOOGL） |
| `pytest -m slow` | ~25 分鐘 | 全部 slow | 完整 live 驗收 |

**Markers：**
- `slow` — 需要網路，排除於預設 CI
- `b1` — B1 overflow-row tests（slow 子集）
- `cf_overflow` — CF YTD overflow 正確性測試（slow 子集）

## 待辦功能

### ✅ 已完成

- **Non-GAAP 指標名稱正規化**（2026-04-26）：`_normalize_nongaap_metrics()` 剝除期間 token、去重比較期間、過濾展望指標
- **金融股警告**（2026-04-26）：`_FINANCIAL_SECTOR_TICKERS` + fetch 後 log 警告
- **Tab 2 Non-GAAP 批量支援**（2026-04-26）：Checkbutton + `_worker_batch(fetch_nongaap)`
- **Session 15 修復（2026-04-26）**：pre-XBRL early exit、Dividends Paid bug、Net Income / Revenue fallback、Total Non-op guard、Investment Proceeds / Debt Proceeds / Debt Repayments 多概念加總、FY Label 對齊公司財年
- **日期區間 / 報表類型 / 快速掃描 UI（Session 16，2026-04-29）**：Tab 1 / Tab 2 加入起始年、結束年、報表類型（Q/Y/Both）、快速掃描下拉選單，inline 進階設定
- **start_year > end_year 驗證（Session 17，2026-05-03）**：Tab 1 + Tab 2 各加 guard，避免區間反轉
- **Index Sheet 品質檢測（Session 18，2026-05-05）**：`_compute_quality()` + `ALL_KEY_ROWS`；Index 新增 E 欄完成度分數（`9/9 ✓` / `N/9 ⚠`）與表格下方品質明細區塊

## Known Issues（已知限制，暫不修）

- **Investment Proceeds**：XBRL 沒有單一加總行，取 first match
- **金融股（GS/JPM）**：現行模板 BS/IS 部分空白，待獨立模板（已有 UI 警告）
- **NG 分類誤判**：keyword-based 分類，label 含 "excluding" 的 GAAP 行可能誤進 NG sheet（可接受方向）
- **Data_EPS_Recon 從未產生**：edgartools `eps_reconciliation` 對 NVDA/AAPL/MSFT 均回傳 None，非 XBRL-tagged 公司無解；待 edgartools 改善或改用 AI 解析方案
