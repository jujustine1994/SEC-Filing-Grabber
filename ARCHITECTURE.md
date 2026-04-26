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
| excel_writer.py | 寫 Data_* sheets 至 output/TICKER.xlsx |
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
    ├─ _get_earnings_filings()    → 8-K Item 2.02 清單
    ├─ _extract_eps_recon()       → edgartools eps_reconciliation
    ├─ _extract_nongaap_metrics() → AI 解析 EX-99.1 press release
    ├─ _build_eps_recon_table()   → Data_EPS_Recon
    └─ _build_nongaap_table()     → Data_NonGAAP
    ↓（中間結果寫入 nongaap_cache.json，增量更新）
    ↓
excel_writer.py
    → 全量改寫所有 Data_* sheets，不碰 My_* 等其他 sheets
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

**CF 限制：** YTD filings（Q2/Q3）的 overflow 不收集，因為 YTD overflow 需要跨 filing 減法（地雷十二延伸）。

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

## 待辦功能

### 🟡 中優先
1. **Non-GAAP cache ticker 隔離**：多公司共用同一 output_dir 時 `nongaap_cache.json` 互蓋
2. **CF Q2/Q3 YTD overflow**（地雷十二延伸）：YTD filings 的 overflow 目前跳過；需跨 filing 減法
3. **Non-GAAP：NVDA 指標名稱帶期間後綴**：同指標跨季是不同 row，表格超稀疏

### 🟢 低優先
4. **金融股模板**（GS/JPM）：UI 自動偵測 + 警告（已設計，延後實作）
5. **批量更新（Tab 2）的 Non-GAAP 支援**：目前批量只跑 GAAP
6. **Non-GAAP：Data_EPS_Recon 從未產生**：edgartools eps_reconciliation API 對主要公司回傳空

## Known Issues（已知限制，暫不修）

- **Investment Proceeds**：XBRL 沒有單一加總行，取 first match
- **金融股（GS/JPM）**：現行模板 BS/IS 部分空白，待獨立模板
- **CF Q2/Q3 overflow 跳過**：YTD filings overflow 需跨 filing 減法（見 待辦 2）
- **NG 分類誤判**：keyword-based 分類，label 含 "excluding" 的 GAAP 行可能誤進 NG sheet（可接受方向）
