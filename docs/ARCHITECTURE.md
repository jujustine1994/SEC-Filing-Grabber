# SEC Financial Fetcher — Architecture

## File Map

> 2026-08-12 目錄結構整理：17 個 `.py` 全部搬進 `src/`（下表路徑已更新），
> `conftest.py`／`company_cache.json`／`config.example.json` 判斷後留根目錄不動。

| File | Role |
|------|------|
| 啟動器.bat | 薄 BAT，呼叫 launcher.ps1 |
| launcher.ps1 | 環境檢查、uv venv、安裝套件、啟動 src/main.py |
| src/main.py | Tkinter GUI，兩個 tab + 兩個 popup |
| src/cli.py | 指令列介面（給外部 skill）：`gaap` / `press-release` 兩個子指令，薄封裝，零 AI |
| src/output_tables.py | `append_ratio_table()`：決定最後寫進 Excel 的 sheet 清單。GUI 與 CLI 共用同一份 |
| src/fiscal_input.py | Index 上「財年起始月」可編輯輸入格 + 由它驅動的期間標籤公式（定義名稱 `FY_START_MONTH`） |
| src/press_release_tables.py | 8-K 新聞稿表格的確定性解析（`pandas.read_html` + Workiva 版面規則），零 AI |
| src/config.py | load_config() / save_config()。`language` 欄位靠 merge-with-defaults 自動補，舊 config.json 零遷移 |
| src/fetcher_gaap.py | edgartools XBRL 抓取 → StatementTable 列表 |
| src/fetcher_nongaap.py | 8-K press release 抓取 → EPS Recon + Non-GAAP StatementTable |
| src/excel_writer.py | 寫 Data_* sheets 至 output/TICKER.xlsx，並呼叫 excel_formatter |
| src/excel_formatter.py | 寫 Index sheet（品質明細）、設欄寬、凍結窗格、數值分類（÷1M／百分比／每股） |
| src/ratios.py | `Data_Ratios`：37 個常見比率，值 + B 欄算法文字 + 列名單位後綴 |
| src/nongaap_layout.py | `Data_NonGAAP` 固定模板版面（Core／調節／overflow／年度分區） |
| src/segments.py | `Data_Segments`：把 `Data_Seg_*` 寬表彙成單一長格式表 |
| src/metric_rules.py | **Non-GAAP 指標名稱規則表（唯一可調整處）**：期間 token、guidance 詞、中英對照、Excel 數值分類關鍵字 |
| src/override_engine.py | 自動修復缺失 key rows（E1 fuzzy + E2 LLM） |
| src/errsafe.py | `_exc_status()`：從例外物件安全萃取 HTTP status，main / fetcher_* 共用 |
| src/zh_labels.py | 薄 wrapper：`zh_label()` / `ratio_label()` / `meta_label()` / `axis_label()`，查 `locales/` 的 `acct.*` / `ratio.*` / `meta.*` / `axis.*`。譯文本體已遷入 locale |
| src/i18n.py | **多語言核心**：`LANGUAGES` 登錄表（代號／顯示名／Excel 字型）、`set_lang()` / `t()` / `excel_font()`。`t()` 的 fallback 鏈為 目標語言 → 繁中 → key 本身 |
| src/locales/*.py | 四份字串表（zh_tw／zh_cn／en／ja），各 341 條。繁中是母表，其餘從它翻出來 |
| conftest.py | pytest 探索設施（rootdir 定位 + slow/b1/cf_overflow marker 註冊），留根目錄不進 `src/` |
| config.json | 使用者設定（gitignored） |
| config.example.json | 範本（committed，留根目錄） |
| company_cache.json | Ticker → 公司名快取（committed，留根目錄——程式主動讀寫的執行期狀態，非純靜態資料） |
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
    │       └─ _dedupe_by_label() → 同標籤只留一份：有 Item 9.01（＝有新聞稿附件）優先，其次取最新。仍只讀 listing metadata
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
| `language` | 介面與 Excel 顯示語言：`zh_tw` / `zh_cn` / `en` / `ja`。重開程式生效。**預設是空字串＝「使用者還沒選過」**，首次啟動據此決定要不要跳選語言視窗；不另開 `language_chosen` 布林值，兩個欄位描述同一件事遲早不同步 |
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
     A（機器鍵，永遠英文）  B（隨語言）      C（公司原文，永遠英文）  D, E, F...
 1   AAPL                                                       =公式 → FY2026Q1
 2                                                              2026-01-29        ← 申報日
 3   Fiscal Quarter        公司財年基準的季度                     =公式 → FY2026FQ1
 4   Calendar Quarter      日曆年基準的季度                       =公式 → 2025Q4
 5   Period End            該期實際結束日                         2025-12-27        ← XBRL 真實日期（靜態）
 6   (空)
 7   Income Statement      損益表                                                   ← section header
 8   Revenue               營業收入         Net sales             124300.0
...
31   Diluted Shares        稀釋加權平均股數
32-36 (空 5 列)
37   Balance Sheet        資產負債表
38   Cash                 現金及約當現金   Cash and cash equiv.  30000.0
...
85   Cash Flow            現金流量表
98   Operating Cash Flow  營業活動現金流
114  Free Cash Flow       自由現金流（衍生）
...
     Other (as reported)  ← overflow 一律在最底部，所以上面的列號跨公司固定
```

- **Col A** = 標準指標名稱（英文；程式一律用這欄比對）
- **Col B** = 中文說明（表在 `zh_labels.py`，改它不影響任何計算）
- **Col C** = Original Item（公司的 XBRL 原始標籤）
- **Col D+** = 各期數據（oldest → newest）

**第 1、3、4 列是 Excel 公式**，由 `Index!B4`（定義名稱 `FY_START_MONTH`）驅動——
財年結束月是程式判讀的，會出錯，使用者改那一格就能修正整本活頁簿的期間標籤。
第 5 列是公式的錨，永遠是 XBRL 的靜態值。細節見 `fiscal_input.py`。

> **實測列位（AAPL，2026-08-12 隨 BS 新增 Total Non-current Assets/Liabilities 重驗）**：
> `Revenue` 8、`Gross Profit` 10、`Operating Income` 17、`Net Income` 24、`Cash` 38、
> `Total Assets` 52、`Operating Cash Flow` 100、`Capex` 101、`Free Cash Flow` 116。
> `Total Assets` 之後（含）的列位比 BS 改動前全部 +1～+2，改跨檔案 `MATCH` 公式的人要注意。

### Data_Financials_NG(Q) / Data_Financials_NG(Y)（有 Non-GAAP overflow 時才產生）

格式與 Data_Financials 完全相同，但每個 section 只含 Non-GAAP overflow 行（無固定模板行）。  
Non-GAAP 判斷：label 含 "adjusted"/"non-gaap"/"non gaap"/"excluding"/"excl."/"ex-"。  
**與 Data_EPS_Recon / Data_NonGAAP 完全獨立**（來源不同：這裡是 XBRL，那裡是 8-K press release）。

### Data_Segments（長格式，各軸合併於一張）

A 欄 `{XBRL 概念} — {成員}`、B 欄維度軸的中文分類、C 欄原始軸名、D 欄起各期數值。

**公司改變分類時怎麼呈現**（MSFT FY2025 改過營收分類，實測）：

```
                                     2023 →→→→→→→→→→→→→→→→→→→→ 2026
Office Products and Cloud Services   11.8  12.4  13.1  13.5  13.9
Windows                               4.8   5.3   5.6   5.3   5.9
Devices                               1.4   1.3   1.1   1.3   1.1
Microsoft 365 Commercial                                       20.4  21.1  ...
Microsoft 365 Consumer                                          1.7   1.8  ...
Windows and Devices                                             4.3   4.5  ...
```

**新舊分類各自成列，各自只在存在的期間有值，不硬接成一條線。** MSFT 的新舊分類不是一對一（Office → M365 Commercial + M365 Consumer），硬接等於替使用者做判斷，而且會錯。要接是使用者（或下游 skill）的工作，工具只負責照實落地。

改名的情況同理：`Search and News Advertising` 與 `Search Advertising` 是兩列。

**維度軸為什麼一定要標**：XBRL 的分類細項掛在不同的軸上，只看成員名稱會混進根本不是 segment 的東西——MSFT 實測有 `Retained earnings`（權益項目軸）與 `Service Life`（固定資產耐用年限軸）。B/C 欄把軸標出來讓使用者自行篩選；**不過濾、不丟棄**，軸表在 `zh_labels.AXIS_LABELS`，沒收錄的標成「其他維度」。

### Data_EPS_Recon

EPS 調和表（GAAP EPS → 調整項 → Non-GAAP EPS）。B 欄為空（無 XBRL labels）。

### Data_NonGAAP

AI 從 8-K press release 提取的所有 Non-GAAP / Adjusted / Excluding 指標。跨季取聯集，缺的季填 None。

**名稱正規化與合併（2026-08-01 起）**：規則表在 `metric_rules.py`，作用在**讀取快取**階段而非寫入階段，因此改規則表後重跑即生效，不必刪 `nongaap_cache.json` 重呼叫 AI。流程為
`_normalize_nongaap_metrics()`（剝期間 token、丟 guidance、當季優先於年度）
→ `_canonicalize_metric_name()`（中文詞彙換英文、同義名合併）
→ `_metric_merge_key()`（忽略大小寫與標點的跨季合併鍵）。
Excel 數值分類（每股 → 百分比 → 股數 → 金額）同樣讀 `metric_rules.py` 的關鍵字表；ASCII 關鍵字一律以詞界比對（`Operations` 含 `ratio`、`Corporate` 含 `rate`）。

> ⚠️ **欄位標籤已知不準（2026-08-07 完整量化，未修）**：季度標籤由 `_period_to_quarter_label()` 依 8-K 的 `period_of_report` 推算，但該欄在 Item 2.02 8-K 上存的是**發布日**而非財期結束日。實測 16 家 128 份、成功比對 119 份，**只有 13% 標對**，偏移量 −3 到 +1 季且由財年結束月決定（NVDA/CRM −3、ORCL/QRVO −2、MSFT/MU/COST/PANW/WDC −1、ARLO/AMD/INTC/NOW +1、AAPL/AVGO 0）。同根因下 dedupe「留最舊」實測撞到 2 次、兩次都留錯（WDC 整季消失、QRVO 拿到 preliminary）。完整報告見 `docs/8k-period-off-by-one.md`。`Data_Financials` 走 XBRL 不受影響。
>
> ✅ **2026-08-09 已處理**：dedupe 改成「有 Item 9.01 優先、其次最新」（`_dedupe_by_label()`）；標籤採方案 B+——`_period_to_quarter_label()` 的 `label` 保留原值不動（它只用於列清單分組與年份篩選），`cli.py press-release` 另外吐 `period_end` 與正確的 `fiscal_label`。15 家 120 份實測期末日 120/120、偏移全對。`--years` 篩的仍是發布日，年份邊界可能差到 3 季。

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
| BS_TEMPLATE | 44 | 同上（含 Total Non-current Assets/Liabilities，2026-08-12 新增） |
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

## 多語言（i18n）

2026-08-14 導入。四種語言：繁體中文、简体中文、English、日本語。

### 分工

```
src/i18n.py          LANGUAGES 登錄表 + set_lang() / t() / excel_font()
src/locales/zh_tw.py 繁中母表（341 條），其餘語言從它翻出來
src/locales/zh_cn.py 简中
src/locales/en.py    English
src/locales/ja.py    日本語
```

`t(key, **fmt)` 的查找順序是 **目標語言 → 繁中 → key 本身**。查不到絕不
raise、絕不回空字串——最壞情況畫面顯示 `gui.btn.run` 這串 key，一眼看得出
哪裡漏翻；回空字串會變成看不見的按鈕。

key 命名空間：`gui.*`（介面）、`acct.*`（三表科目）、`ratio.*`
（Data_Ratios 列名）、`meta.*`（Data_Meta 欄位）、`axis.*`（segment 維度
軸）、`xls.*`（Index 版面）、`err.*`（使用者可見的錯誤）。

### 什麼跟著語言變、什麼不變

| 位置 | 是否隨語言變 | 原因 |
|---|---|---|
| Excel A 欄（`Revenue`、`Gross Margin (%)`） | ❌ 永遠英文 | 下游跨檔案 `MATCH` 的機器鍵；`Data_Ratios` 的單位後綴還兼任數字格式判斷依據 |
| Excel B 欄 | ✅ | 純顯示，改錯不影響任何計算 |
| Excel C 欄（`Net sales`） | ❌ 永遠公司原文 | 開放集合（七家公司的 Revenue 有六種寫法），且用途是拿回 10-Q Ctrl+F 核對 |
| Excel D 欄起的值 | ❌ | 數字 |
| Index sheet 版面文字 | ✅ | |
| GUI | ✅ | 重開程式生效 |
| `logs/app.log` | ❌ 永遠繁中 | 給維護者除錯用，跟著使用者語言變等於自廢。`_write_log*` 落檔，`self._log()` 預設只推 UI 所以照翻 |
| `cli.py` 的主控台訊息與 argparse help | ❌ 繁中 | 給 skill 與維護者的開發者介面。它產出的 Excel 語言由 `--lang` 控制 |
| Watchlist 的「未分類」群組名 | ❌ 存的值固定 | 那是寫進 config.json 的**資料**。顯示走 `_group_display()`、存回走 `_group_stored()`；少了這層，日文使用者會長出第二個空群組 |

### Excel 字型隨語言

`i18n.LANGUAGES` 第三欄。微軟正黑體缺日文假名字形，日文用 Yu Gothic、
简中用 Microsoft YaHei。`excel_formatter._font()` 與 `fiscal_input._font()`
都是**呼叫時**才解析——綁在 import 當下會凍結在預設語言。

### 首次啟動選語言

`main._pick_language_on_first_run()`，在建主視窗之前跑。判斷依據是
`config.json` 的 `language` **不是合法代號**（空字串、缺鍵、舊版怪值）。

視窗刻意**不翻譯**：這時候還不知道使用者要哪個語言，用任一種當說明都在賭。
只有一個英文抬頭，其餘全是各語言的自稱，看得懂哪個就點哪個。按鈕由
`i18n.LANGUAGES` 生成。

直接關掉視窗＝接受第一個選項並**照樣存檔**——需求是「選完就記住不要再跳」，
關掉還一直跳才是煩人。選錯了在「進階設定」隨時能改。

### 新增語言

兩步，不必碰 `main.py` 或任何功能程式碼：

1. 複製一份 `locales/en.py` 改譯文
2. `i18n.LANGUAGES` 加一行 `("ko", "한국어", "Malgun Gothic")`

下拉選單、Excel 字型、`tests/test_i18n.py` 的漏 key 檢查全部自動涵蓋。

### 三道防線（`tests/test_i18n.py`，38 條）

1. **四語言 key 集合必須完全一致**——新增語言時漏翻幾條是必然，靠人眼比對
   341 條不可能可靠
2. **placeholder 必須一致**——譯文把 `{name}` 打錯不會 crash，只會靜默吐出
   未格式化的字串，特別容易漏
3. **`src/` 不得再出現寫死的中日文字面**。這條是**永久**的，擋的是下一次而
   不是這一次：新增功能時順手寫個中文按鈕標籤最自然不過，沒有它三個月後就
   又回到全部寫死的狀態。豁免清單在測試檔裡，每條都要寫理由

另外釘住兩件不屬於上述三類、但錯了會很難發現的事：

4. **Watchlist 群組名稱的顯示／儲存往返**（`_group_display` ↔ `_group_stored`）。
   這是整個 i18n 唯一會污染使用者資料的地方
5. **財年區間 Excel 公式的引號逸出與月份格式**。譯文含撇號會把公式切碎
   （Excel 顯示 `#NAME?` 或拒絕開檔，而 Python 一點錯都不會報）；月份格式碼
   要跟著語言走，寫死 `"m"` 的話英文會得到 `FY 10 – 9` 而不是 `FY Oct – Sep`

### 改動 Excel 相關程式碼前先存基準

`scripts/excel_golden.py` 把 `output/_final/*.xlsx` 讀回來、走**真正的**寫檔
流程重產，dump 每一格的值／數字格式／字型／粗體／底色。不打網路。

```
./venv/Scripts/python.exe scripts/excel_golden.py make  base
# ...改 excel_writer / excel_formatter / ratios / fiscal_input...
./venv/Scripts/python.exe scripts/excel_golden.py make  new
./venv/Scripts/python.exe scripts/excel_golden.py check base new
```

單元測試驗的是邏輯，這支驗的是「產出來那份 xlsx 有沒有變」——「÷1M 沒套到」
「百分比格式掉了」「字型混到別的」這幾種最常見的排版回歸，值都是對的，
只有逐格比對抓得到。

## 測試分層

| 指令 | 時間 | 測試數 | 用途 |
|------|------|--------|------|
| `python -m pytest -m "not slow"` | ~25 秒 | 725 | Unit tests（每次改 code 後跑） |
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
