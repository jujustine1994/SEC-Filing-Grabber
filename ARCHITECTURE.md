# SEC Financial Fetcher — Architecture

## File Map

| File | Role |
|------|------|
| 啟動器.bat | 薄 BAT，呼叫 launcher.ps1 |
| launcher.ps1 | 環境檢查、uv venv、安裝套件、啟動 main.py |
| main.py | Tkinter GUI，兩個 tab + 兩個 popup |
| config.py | load_config() / save_config() |
| fetcher_gaap.py | edgartools XBRL 抓取 → StatementTable 列表 |
| fetcher_nongaap.py | Non-GAAP stub（Phase 2） |
| excel_writer.py | 寫 Data_* sheets 至 output/TICKER.xlsx |
| config.json | 使用者設定（gitignored） |
| config.example.json | 範本（committed） |
| company_cache.json | Ticker → 公司名快取（committed） |
| output/ | 輸出的 Excel 檔（gitignored） |

## Data Flow

```
使用者輸入 Ticker（Tab 1）或從 Watchlist 選取（Tab 2）
    ↓
fetcher_gaap.py
    ├─ _build_is_table()    → IS 22-row 固定模板
    ├─ _build_bs_table()    → BS 41-row 固定模板
    ├─ _build_cf_table()    → CF 25-row 固定模板 + FCF 衍生
    ├─ _merge_financials()  → 合併成 Data_Financials
    ├─ _build_segment_tables() → Data_Seg_* (多個)
    └─ _build_meta_table()  → Data_Meta
    ↓
excel_writer.py
    → 全量改寫所有 Data_* sheets，不碰 My_* 等其他 sheets
    → output/TICKER.xlsx
```

## Key Config Variables (config.json)

| 鍵 | 說明 |
|----|------|
| `identity` | SEC EDGAR 身份字串（必填，格式：名字 空格 信箱） |
| `output_dir` | Excel 輸出路徑（預設 "output"） |
| `max_filings` | 最多抓幾筆 10-Q（預設 80，約 20 年） |
| `watchlist` | [{ticker, name}, ...] 清單 |
| `ai.provider` | "google" / "openai" / "anthropic" |
| `ai.model` | 模型名稱 |
| `ai.api_key` | API Key（gitignored） |

## Excel Sheet Layout

### Data_Financials（主要輸出）

```
A1=ticker  B1=空  C1=FY2024Q1  D1=FY2024Q2  ...
A2=空      B2=空  C2=2024-02-01 D2=2024-05-03 ...
A3=Income Statement  B3=空  C3..=None  (section header)
A4=Revenue  B4=Net sales  C4=100.0  D4=105.0  ...
...
A26=Balance Sheet  (section header)
A27=Cash  B27=Cash and cash equivalents  ...
...
A69=Cash Flow  (section header)
A70=Net Income  B70=Net income  ...
```

- **Col A** = 標準指標名稱（Std Name）
- **Col B** = Original Item（公司 XBRL 原始標籤，section header 行為空）
- **Col C+** = 季度數據（oldest → newest）

### Data_Seg_*

每個有 segment breakdown 的 IS 概念一張 sheet，格式同上但沒有 B 欄 labels。

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
| CF_TEMPLATE | 26 | 同上（含 Free Cash Flow DERIVED 行） |

## IS Post-processing Fallbacks

在每個 filing 的 row_vals 計算完後執行：

1. **Total Non-op**：若 XBRL None → `Pre-tax Income − Operating Income`
2. **Net Income**：若 NetIncome None → 試 `ProfitLoss`（BA/TSLA/XOM/WMT）
3. **D&A**：若 DepreciationExpense None → label fallback `"depreciation"`（TSLA）

## CF Post-processing

- **Free Cash Flow** = `Operating Cash Flow − Capex`（每季計算）

## 待辦功能（下一個 AI 看到請提醒使用者）

### 🔴 優先（下次開始前必做）
1. **實機測試**：對 AAPL、TSLA、BA、XOM 跑一次，確認 Data_Financials 輸出正確
2. **main.py 確認**：fetch 改輸出 Data_Financials，確認 GUI 沒有參照舊的 Data_IS/BS/CF sheet 名稱

### 🟡 中優先
3. **金融股模板**（設計已完成，待實作）：
   - GS/JPM 用不同的 IS/BS 模板
   - 自動偵測：BS 含 `TotalDeposits` std_concept → 金融股
   - UI 警告：偵測到金融股時彈出提示
4. **Excel Template 著色功能**：
   - 使用者建立 `template.xlsx`，預先著色各行
   - 工具開啟 template 後只填值（不改格式），利用 openpyxl 保留 cell formatting
   - 季度欄擴充時複製前一欄格式

### 🟢 低優先
5. **Non-GAAP 抓取**（Phase 2）：EPS reconciliation、adjusted metrics

## Known Issues

- **Investment Proceeds**：XBRL 沒有單一加總行，取 first match（已知限制，不修）
- **金融股（GS/JPM）**：現行模板 BS/IS 大量空白，待獨立模板實作
