# SEC Financial Fetcher — Architecture

## File Map

> 2026-08-12 目錄結構整理：17 個 `.py` 全部搬進 `src/`（下表路徑已更新），
> `conftest.py`／`company_cache.json`／`config.example.json` 判斷後留根目錄不動。

| File | Role |
|------|------|
| 啟動器.bat | 薄 BAT，呼叫 launcher.ps1 |
| launcher.ps1 | 環境檢查、uv venv、安裝套件、啟動 src/main.py。**不依賴系統 Python 也不用 winget**——`uv venv venv --python 3.13` 找不到直譯器時 uv 自己下載（見 docs/CHANGELOG.md 2026-08-17） |
| src/main.py | Tkinter GUI，三個 tab（單一公司／批量更新／進階設定）+ Watchlist popup。另有 `work_area()` / `fit_geometry()`（視窗置中，見「視窗擺放」） |
| src/cli.py | 指令列介面（給外部 skill）：`gaap` / `press-release` 兩個子指令，薄封裝，零 AI |
| src/output_tables.py | `append_ratio_table()`：決定最後寫進 Excel 的 sheet 清單；`has_any_data()`：一期都沒抓到就別寫檔。GUI 與 CLI 共用同一份 |
| src/net_retry.py | 網路層韌性：`is_network_error()`（沿 `__cause__` 走訪）、`with_retry()`（只對網路類退避重試）、`sec_reachable()`（實際戳一次 SEC）、`configure_timeouts()` |
| src/fetch_ledger.py | `FetchLedger`：記下哪幾期沒抓到、是網路還是資料造成的，產出給人看的摘要。見「抓取缺漏」 |
| src/fiscal_input.py | Index 上「財年起始月」可編輯輸入格 + 由它驅動的期間標籤公式（定義名稱 `FY_START_MONTH`） |
| src/press_release_tables.py | 8-K 新聞稿表格的確定性解析（`pandas.read_html` + Workiva 版面規則），零 AI |
| src/config.py | load_config() / save_config()。`language` 欄位靠 merge-with-defaults 自動補，舊 config.json 零遷移 |
| src/fetcher_gaap.py | edgartools XBRL 抓取 → StatementTable 列表 |
| src/fetcher_nongaap.py | 8-K press release 抓取 → EPS Recon + Non-GAAP StatementTable |
| src/excel_writer.py | 寫 Data_* sheets 至 output/TICKER.xlsx，並呼叫 excel_formatter |
| src/excel_formatter.py | 寫 Index sheet（含**資料完整度**區塊，見下方「缺漏判斷」）、設欄寬、凍結窗格、數值分類（÷1M／百分比／每股） |
| src/data_quality.py | **抓取結果的缺漏判斷**（2026-08-22）：季度斷層／整欄稀疏／中間有洞／整列全空且矛盾。純函式，不打網路、不看別家公司。見下方「缺漏判斷」 |
| src/comparison.py | **跨公司比較的抓取協調**：對每個 ticker 呼叫 `fetch_gaap_statements()`，重組成 `{指標: {ticker: {日曆季: 值}}}`。單一公司失敗不中斷其他家 |
| src/comparison_writer.py | 跨公司比較的 Excel 輸出：`Compare_Data`／`Snapshot`／`Snapshot_Manual`／`Chart_<指標>` |
| src/fetcher_facts.py | **走 SEC companyfacts API 的平行取數路徑（G11 spike，尚未接上主流程）**。見下方「companyfacts 平行路徑」 |
| src/facts_mapping.py | 模板列 → us-gaap concept 對照表。**不是手填的**，是拿現行路徑當答案卷對 52 家反推出來的，每列附證據註解 |
| src/ratios.py | `Data_Ratios`：37 個常見比率，值 + B 欄算法文字 + 列名單位後綴 |
| src/nongaap_layout.py | `Data_NonGAAP` 固定模板版面（Core／調節／overflow／年度分區） |
| src/segments.py | `Data_Segments`：把 `Data_Seg_*` 寬表彙成單一長格式表 |
| src/metric_rules.py | **Non-GAAP 指標名稱規則表（唯一可調整處）**：期間 token、guidance 詞、中英對照、Excel 數值分類關鍵字 |
| src/override_engine.py | 自動修復缺失 key rows（E1 fuzzy + E2 LLM） |
| src/errsafe.py | `_exc_status()`：從例外物件安全萃取 HTTP status，main / fetcher_* 共用 |
| src/zh_labels.py | 薄 wrapper：`zh_label()` / `ratio_label()` / `meta_label()` / `axis_label()`，查 `locales/` 的 `acct.*` / `ratio.*` / `meta.*` / `axis.*`。譯文本體已遷入 locale |
| src/i18n.py | **多語言核心**：`LANGUAGES` 登錄表（代號／顯示名／Excel 字型）、`set_lang()` / `t()` / `excel_font()`。`t()` 的 fallback 鏈為 目標語言 → 繁中 → key 本身 |
| src/locales/*.py | 四份字串表（zh_tw／zh_cn／en／ja）。繁中是母表，其餘從它翻出來 |
| conftest.py | pytest 探索設施（rootdir 定位 + slow/b1/cf_overflow marker 註冊），留根目錄不進 `src/` |
| config.json | 使用者設定（gitignored） |
| config.example.json | 範本（committed，留根目錄） |
| company_cache.json | Ticker → 公司名快取，留根目錄。**2026-08-17 起 gitignored**——它是程式主動讀寫、且會自己重建的執行期狀態（405 KB），先前誤入版控 |
| output/ | 輸出的 Excel 檔（gitignored） |
| nongaap_cache.json | 各公司輸出資料夾內，Non-GAAP 快取（runtime，非 git） |
| scripts/打包.bat + pack.ps1 | **打包散布用 zip**：白名單複製 → 壓縮 → 12 項自我驗證，任一項沒過就刪 zip 並 exit 1。雙擊即可，是 `docs/PACKAGING.md` 的可執行版本（兩邊須同步） |
| docs/PACKAGING.md | 打包作業指示（給 AI 照著做的版本）：包含／排除清單、檔名規則、驗證步驟 |
| docs/RECIPIENT-README.txt | 給收件人看的說明，打包時改名為 `先讀我.txt` 放進 zip。只講兩件事：填 SEC EDGAR Identity、首次啟動選語言 |
| dist/ | 打包產物（gitignored） |

## Data Flow

```
使用者輸入 Ticker（Tab 1）或從 Watchlist 選取（Tab 2）
    ↓
fetcher_gaap.py
    ├─ _build_is_table(filings_k)  → is_ann（年報，10-K，先建——季報 Q4 要用它反推）
    ├─ _build_bs_table(filings_k)  → bs_ann
    ├─ _build_cf_table(filings_k)  → cf_ann
    │
    ├─ _build_is_table(filings_q)  → (gaap_tbl, ng_tbl)  IS 22-row 模板 + GAAP/NG overflow
    ├─ _build_bs_table(filings_q)  → (gaap_tbl, ng_tbl)  BS 41-row 模板 + GAAP/NG overflow
    ├─ _build_cf_table(filings_q)  → (gaap_tbl, ng_tbl)  CF 26-row 模板 + GAAP/NG overflow
    │
    ├─ override_engine.check_key_rows()  → 找全 None 的 key rows
    ├─ override_engine.run_diagnosis()   → E1 fuzzy + E2 LLM → save_overrides()
    │   （有新 override 時重跑三個 build 函式）
    │
    ├─ _synthesize_q4(is_tbl, is_ann, …)  → 補季報表的 Q4 欄（見下方「Q4 推算」）
    ├─ _synthesize_q4(bs_tbl, bs_ann, …)
    ├─ _synthesize_q4(cf_tbl, cf_ann, …)
    │
    ├─ _merge_financials(is, bs, cf)    → Data_Financials(Q)    ← 主輸出，含推算的 Q4
    ├─ _merge_financials(ng_is, ng_bs, ng_cf) → Data_Financials_NG(Q)  ← 有 NG overflow 時
    ├─ _merge_financials(is_ann, bs_ann, cf_ann) → Data_Financials(Y) ← 10-K
    ├─ _merge_financials(ng annual...)  → Data_Financials_NG(Y) ← 有 NG overflow 時
    ├─ _build_segment_tables()          → Data_Seg_* (多個)
    └─ _build_meta_table()              → Data_Meta

fetcher_nongaap.py（勾選 Non-GAAP 時，完全獨立於 GAAP fetcher）
    ├─ _list_earnings_filings(fiscal_year_end=MMDD) → 在 SEC 申報清單階段以 items（2.02）+ period_of_report 篩選、去重、套年份與 max_filings，零下載
    │       ├─ _label_for_listing() → 季度標籤：拿得到 fiscal_year_end 就走零下載規則（B5），否則逐份退回 _period_to_quarter_label()
    │       └─ _dedupe_by_label() → 同標籤只留一份：有 Item 9.01（＝有新聞稿附件）優先，其次取最新。仍只讀 listing metadata
    ├─ _find_missing_quarters()   → 偵測季度序列缺口
    ├─ _recover_missing_quarters(fiscal_year_end)→ 只對缺季區間逐筆 obj() 深掃，用 has_earnings 找回未標 2.02 的財報（label 與列清單同一套）
    ├─（以上皆不下載**文件**；B5 之後多一次 company 層級的 fiscal_year_end 查詢，
    │   一個 ticker 一次 submissions 請求，本來就要查財年結束月。
    │   filing.obj() 移進 fetch_nongaap_statements() 的逐季迴圈，
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

#### 跨公司模板公式怎麼寫

三個保證：固定 sheet 名稱、固定列位（overflow 全部集中在最底部）、三種期間標籤各佔一列（用途不同，見下表）。

| 列 | 內容 | 基準 | 什麼時候用 |
|---|---|---|---|
| 1 | `FY2026Q1` | 公司財年 | 主要欄位鍵，`MATCH` 用這列 |
| 3 | `FY2026FQ1` | 公司財年 | 同上，加 `FQ` 標記避免與日曆季混淆 |
| 4 | `2026Q1` | 日曆年 | 跨產業比同一個日曆期間、對總經數據 |
| 5 | `2026-03-29` | 實際期末日 | 精確對齊；也是判斷兩家是否真的同期的唯一依據 |

比財季（同業比較）：

```excel
=INDEX('[AAPL.xlsx]Data_Financials(Q)'!$D8:$AZ8,
       MATCH("FY2026Q1",'[AAPL.xlsx]Data_Financials(Q)'!$D$1:$AZ$1,0))
```

比日曆季（跨產業／對總經）把 `$1` 換成 `$4`（值變成 `2026Q1` 不含 FY）。`$D8:$AZ8` 的 8 是營收列，換 `$D38` 就是現金。

> ⚠ **跨檔案讀取、而且來源檔關著時，改用第 5 列當 key**。第 1、3、4 列是公式，
> openpyxl 不算公式也不寫快取值——來源檔開著時 Excel 會重算沒問題，關著時外部
> 參照只讀得到檔案裡的值，那三列在那裡是空的，`MATCH` 會回 `#N/A`。第 5 列是
> 靜態文字，永遠讀得到：
>
> ```excel
> =INDEX('C:\...\output\_final\[AAPL.xlsx]Data_Financials(Q)'!$D8:$AZ8,
>        MATCH("2025-12-27",'C:\...\output\_final\[AAPL.xlsx]Data_Financials(Q)'!$D$5:$AZ$5,0))
> ```
>
> 或者把來源檔在 Excel 開一次再存檔，快取值就寫進去了，之後關著也能用 `FY2026Q1` 當 key。

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

## edgartools 到底是什麼（2026-08-23 釐清，很容易誤會）

**它是裝在本機的 pip 套件，不是雲端服務。沒有任何人在遠端幫我們維護財報對應。**

```
Name     edgartools 5.29.0        Author   Dwight Gunning（個人）
License  MIT                      Source   github.com/dgunning/edgartools
```

它連的網域只有兩個，都是 SEC 自己的免費端點：`www.sec.gov`（下載 filing 文件）
與 `data.sec.gov`（官方 XBRL JSON）。**沒有第三方伺服器。**

**XBRL 是在我們的 CPU 上解的**：`venv/Lib/site-packages/edgar/xbrl/` 有 55 個
`.py`、33,078 行，全部在本機跑。這就是「每家 30 秒～2 分鐘」的來源——實測 ARLO
25 份 filing，XBRL 解析 19.9 秒 ＋ `to_dataframe()` 28.4 秒。慢的是解 XML，
不是網路。

### ⚠ 它的「標準科目對照表」很薄，不能當靠山

```
edgar/xbrl/standardization/concept_mappings.json   100 個標準名 → 163 個 us-gaap concept
edgar/xbrl/standardization/company_mappings/       公司專屬 override 共 3 家：BRK.A、MSFT、TSLA
```

美股幾千家上市公司、us-gaap 分類標準上萬個 element，對照表只有 163 筆、3 家特例。
**所以 `standard_concept` 這一欄薄而且會錯**，實測到的：

| 它給的 | 實際是什麼 | 後果 |
|---|---|---|
| `EquityExpenseIncome(BuybackIssued)` | 模板寫的名字沒有括號 | 優先層變成死碼 |
| `IncomeTaxes` ← `DeferredIncomeTaxExpenseBenefit` | 遞延所得稅費用 | 「現金稅」抓到完全不同的東西 |
| `NetCashFromFinancingActivities` ← `ProceedsFromDebtNetOfIssuanceCosts`（GOOGL） | 借款收入 | 語意完全不對 |
| `NonoperatingIncomeExpense` ← D&A（AMD/MRVL，concept 是 `OtherDepreciationAndAmortization`） | 折舊攤銷 | **兩層一起失守**：標錯的 std 擋掉第一層，concept 名字又不含舊的 `DepreciationDepletionAndAmortization`。2026-08-25（G10）把第二層放寬成 `Depreciation\w*Amortization`、補上第三層 `^depreciation` 才救回來 |

**這就是模板要寫成 `(std_concept, fallback, label_hint)` 三層的原因**：
`std_concept` 只是去碰它那張 163 筆的表，碰不到（或碰錯）才靠我們自己的 fallback。

### 結論：`IS/BS/CF_TEMPLATE` 才是這個專案真正的資產

edgartools 負責「把 XBRL 解成 dataframe」，**「哪個 tag 對應到我們的哪一列」
100% 是我們自己維護的**。沒有人會幫我們把它變好；反過來說，錯了我們也修得動
（2026-08-23 H3 一次修好 16 列）。評估任何「換掉 edgartools」的方案時，要分清楚
換掉的是**解析器**還是**對照表**——對照表換不掉，那是自製的。

## `Data_Financials` 的 A / B / C 三欄各是什麼

| 欄 | 內容 | 誰在用 |
|---|---|---|
| A | **模板列名（英文機器鍵）**，如 `Share Repurchases` | 下游靠**固定列位**取值；`_col_b()` 拿它當查表 key。**永遠英文、永遠不可改** |
| B | A 欄的介面語言翻譯（`zh_labels.py` 查表，查不到留白） | 人看的 |
| C | **公司自己在財報上印的那行字**（`df.loc[idx,"label"]`） | 稽核軌跡 |

C 欄常被誤以為是「edgartools 的名稱」，**它不是**——它是公司原文標籤。價值在於
它讓每一格都可回溯：C 欄寫「Common stock acquired」而 A 欄是 `Share Repurchases`，
就看得出我們把 XOM 的哪一行填進了哪一列。這是這個專案相對於黑箱資料源最大的優勢，
**動任何取數路徑之前先確認 C 欄還在**（companyfacts 沒有 presentation linkbase，
切過去這一欄會消失——見 G11 報告第 5 節）。

已知小限制：`row_labels[i]` 只在第一份命中的 filing 寫入（`if i not in row_labels`），
公司改過用詞的話 C 欄只顯示其中一種。

## Template Matching Logic（_match_is_row）

3 層查找 + 2 個修飾參數：

```
Priority 1: standard_concept == std_concept
Priority 2: concept 欄位包含 fallback_suffix（case-insensitive）
Priority 3: label 欄位包含 label_fallback（case-insensitive）

label_hint: 在該層的 candidates 裡**過濾**出 label 含 hint 的行
match:      "first"（預設）= 最早那行；"last" = 最後那行（用於 CF 彙總行）
```

⚠ **`label_hint` 不是「優先選」，是「濾掉」——濾空之後整個優先層被跳過，
不會退回去用 concept 比對。** 所以一個寫太窄的 hint 等於把那一列整個關掉。
2026-08-23（H3）掃 22 家最新 10-Q，光是這個成因就讓 `Cash Taxes Paid` 少 14 家、
`Deferred Revenue, current` 少 12 家、`Share Repurchases` 少 9 家、
`Accounts Receivable` 少 8 家。踩過的坑：hint 寫 `repurchas`，但多數公司的 label
是「Treasury stock purchases」；hint 寫複數 `inventories`，但公司寫單數
`Inventory`；hint 寫 `^net cash|^cash`，但 PG 三個小計都寫
「TOTAL OPERATING ACTIVITIES」。

**hint 只該用來擋掉會抓錯的鄰居，不該用來描述這一列長什麼樣。** 加 hint 之前
先確認它擋的是什麼（例：`Operating Cash Flow` 的 hint 是為了擋現金流量表最下面
的租賃補充揭露列，那個 hint 有存在的必要），並且用實測掃一遍再收工。

### H6（2026-08-25）：hint 太窄與太寬各是什麼下場

H3 那批 hint 是照 **22 家**調的，擴到 **201 家**重掃後有四條明顯太窄
（`scripts/diag_hintsweep.py`，killed 前 → 後）：

| 模板列 | 前 | 後 | 症狀 |
|---|---|---|---|
| Capex | 15 | 3 | 公司寫 Capital spending／investments／premises and equipment |
| Common Stock & APIC | 14 | 2 | 外國註冊公司寫 Ordinary shares、另一批寫 Common shares |
| Cash | 20 | 15 | 寫 Cash／Cash and cash items／and temporary investments |
| Cost of Revenue | 36 | 30 | 能源公用餐飲寫 Purchased crude oil／power／Food, beverage |

四條 hint 現在寫成 `fetcher_gaap` 頂端的具名常數（`_CAPEX_HINT`、`_CASH_HINT`、
`_COMMON_STOCK_HINT`、`_COGS_HINT`），每條旁邊註記它擋的是什麼、為什麼不能再寬。

**放寬的同時要擋住的鄰居（都是實測踩到的，不是假想）**：

| 模板列 | 一放寬就會吃到 | 為什麼不能吃 |
|---|---|---|
| Capex | `CapitalExpendituresIncurredButNotYetPaid`（UNP [32]／AMD／NEE） | 非現金揭露，**`std_concept` 同樣是 `CapitalExpenses`**，混進來會重複計算 |
| Common Stock & APIC | `TreasuryStockCommonValue`（LIN [28]／ABT [51]／AMP [77]／KR [37]） | 庫藏股，**std_concept 同樣是 `CommonEquity`** |
| Cost of Revenue | `LaborAndRelatedExpense`（銀行／鐵路的人事費） | 那 29 家概念上沒有 COGS（同 D8），**留空比填錯好** |
| Cash | `CashAndDueFromBanks`（銀行 7 家） | 口徑不同質，是產品決定不是 bug（TODO H6-1） |

**兩個排除條件本身也會出錯，寫的時候要挑判準**：
- 「label 含 treasury 就踢掉」會誤傷 NSC 的普通股列（它自己寫
  `Common stock, net of treasury shares`）。真正的庫藏股列都帶「at cost」或
  「in treasury」，用那個當判準才分得開
- Capex 的排除詞要放在 negative lookahead（`^(?!.*(?:accrued|not yet paid|payable))`），
  不是在正向 pattern 裡想辦法避開

**擋掉之後那一格會變空，而空白有時候才是對的**：NEE 的 Capex 在 H6 之前抓到的就是
「Accrued property additions」——**填的是錯的數字**。H6 之後那格留空。判斷一個 hint
改得對不對，不能只看「填滿的格子有沒有變多」。

### D&A：延伸 tag 是常態，不是特例（G10，2026-08-25）

`D&A` 這一列的第三層（`^depreciation`）**救回的家數比第二層還多**。201 家最新
10-Q 實測新增命中 13 家，其中 9 家用的是公司自訂延伸 tag：

```
msft_DepreciationAmortizationAndOther          csco_DepreciationAmortizationAndOther
tsla_DepreciationAmortizationAndImpairment     gm_DepreciationAmortizationAndImpairmentChargesOnProperty
acn_DepreciationAmortizationAndOther           mar_DepreciationAmortizationAndOther
odfl_DepreciationAndAmortizationIncludingDebtIssuanceCosts
schw_DepreciationAndAmortizationExcludingAmortizationOfIntangibleAssets
isrg_DepreciationandGainLossonDispositionofPropertyPlantEquipment
```

**MSFT／TSLA 這種規模的公司也在裡面**——延伸 tag 不是小公司才有的邊角情況。
剩下 4 家是 us-gaap 標準 tag 但名字不同（`OtherDepreciationAndAmortization` ←
AMD/MRVL、`UtilitiesOperatingExpenseDepreciationAndAmortization` ← AEP、
`amt_...IncludingDiscontinuedOperations` ← AMT）。

⚠ **第三層一定要 `^depreciation` 開頭錨定**：現金流量表上另外有「Amortization of
acquisition-related intangibles」（AMD/MRVL 都有獨立一列）、債務發行成本攤銷、
遞延佣金攤銷——只寫 `amortization` 會把無形資產攤銷當成 D&A，那是兩個不同科目。

**幾家抓到的是「口徑略寬」的行**，這是接受公司報表表面列示的必然結果，記在這裡
不是 bug：SCHW 那行明講 excluding intangibles、ISRG 那行含處分損益、GM／TSLA 含
減損。都是公司自己在現金流量表上列的那一行。

### 什麼時候該用第三層（label_fallback）而不是放寬 concept

`Cash` 那一列 2026-08-25 補了 `^cash and cash equivalents$`。成因是 ASU 2016-18：
**現金流量表**的期初期末總額必須含受限現金，**資產負債表沒有要求合併列示**，
但有些公司（INTC 2022~2025、PG、SBUX、GILD…）把 BS 那一行 tag 成 ASU 的合併 element
`CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents`（⚠ 名字裡**沒有
"And"**，所以模板的 `CashAndCashEquivalents` 兩層都比不中），而印出來的字仍然是
「Cash and cash equivalents」。

**判準：抓「公司印在報表表面的那一行」。** 所以走第三層 label 比對、正則要窄到
只吃那一行——真的把受限現金併進列示的公司會寫「Cash, cash equivalents and
restricted cash」，窄正則吃不到，口徑不同的自動排除。

副作用實測（201 家最新 10-Q）：**新增命中 11 家、換答案 0 家**。其中 BAC 拿到的是
它自己在 BS 上列的小計 `$229.7bn = 28.1（cash and due from banks）+ 201.6（存放同業）`
——**正好是 ASC 230 現金流量表定義的銀行現金**。JPM 沒列這條小計，所以仍是空的
（銀行 Cash 的完整解法要能加總兩列，屬於 D8）。

## Template 行數摘要

| 報表 | 行數 | 格式 |
|------|------|------|
| IS_TEMPLATE | 22 | 7-tuple (label, std_concept, fallback_suffix, source, match, label_hint, label_fallback) |
| BS_TEMPLATE | 44 | 同上（含 Total Non-current Assets/Liabilities，2026-08-12 新增） |
| CF_TEMPLATE | 26 | 同上（第 26 行為 Free Cash Flow，DERIVED = OCF − |Capex|） |

source 欄位值：`"IS"` / `"BS"` / `"CF"` = 從哪個 DataFrame 取值；`"DERIVED"` = 不做 XBRL 比對，由 post-processing 計算。

第 7 欄 `label_fallback` 是 2026-08-23（H4 第一步）加的第三層，**公司自訂延伸 tag
（`nvda_` / `nee_` 這種）唯一抓得到的方式**——那種 concept 名字每家自己取，只有 label
對得上。第三層後面沒有任何東西再擋它，所以寫得下去就要窄。

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

## 抓取缺漏（2026-08-17）

**規則：抓不到就留空，但一定要講出來。** 不中止、不寫空殼。

原本九個 `filing.obj()` 迴圈都是 `except Exception: continue`——網路在第 5 份
掛掉時那一季被當成「這期沒資料」默默跳過，程式照樣顯示完成，使用者拿到少
一季的 Excel 卻不知道。修正過程走過三版，中間那版「網路斷了就中止不寫檔」
被 CTH 否決（「不希望抓得太嚴格讓資料永遠抓不出來」），定案是本節。

| 狀況 | 行為 |
|---|---|
| 閃斷 | 退避重試 2/4/8 秒，多半救得回來 |
| 某幾期抓不到 | 那幾期留空，其餘照常產出，**主動報告缺了哪幾期** |
| 整個網路斷掉 | 連續 3 期失敗後停止重試（`give_up_retrying`），剩下快速跑完，一樣寫檔一樣提示 |
| 一期都沒抓到 | `has_any_data()` 擋下，**不寫檔**（空殼會蓋掉舊檔，唯一不可逆的傷害） |

**「是網路還是資料」怎麼判斷**：不猜例外類別名稱（那份名單得跟著 httpx 的
版本走，漏一個就誤判）。每一期本來就是一次 EDGAR API 呼叫，失敗當下直接戳
一次 SEC——戳得通代表伺服器有回應、問題在這份資料；戳不通就是網路。每趟
最多戳一次，用 `urllib` 不用 httpx（錯誤處理路徑上不該再依賴可能正是出問題
那一層的東西）。

> ⚠ **`NetworkDownError` 不可以走探測**。它代表「已退避重試三次都連不上」，
> 是最強的網路證據；而探測在事後跑，網路可能已恢復，於是斷網被報成「SEC
> 連得上，是資料問題」，方向完全相反。實機驗收踩過，已釘測試。

**警告顯示三處**：GUI 橘字、`logs/app.log`、Excel 的 `Index!A3`（橘底）。
Index 那份最重要——GUI 的 log 關掉就沒了，而使用者真正會搞混的時點是三天後
重開這份 Excel。原始值另存於 `Data_Meta` 的 `Fetch Gaps` 列。

> 缺漏警告刻意**不放進 Index 下方的「品質明細」區塊**：那一區講「這個科目 SEC
> 沒報」，是資料本身的性質；缺漏講「這次抓取沒拿到」，是這份檔案的狀態，重抓
> 可能就有了。混在一起使用者分不出哪個該重抓。

帳本用 `ContextVar` 串接（`fetcher_gaap.collect_gaps()`），九個抓取函式不必
改簽名。想拿到明細就在外面包一層；不包的話 `fetch_gaap_statements` 自己開
一本，結果寫進 `Data_Meta`。

## 視窗擺放（2026-08-17）

`main.work_area()` 走 Win32 `SPI_GETWORKAREA` 取得**扣掉工作列**的可用矩形，
`fit_geometry()` 算出保證落在裡面的座標。視窗比工作區大就縮到剛好塞得下——
寧可矮一點，也不要讓下緣的「開始抓取」按鈕看不到。

> 原本只呼叫 `geometry("700x650")` 不給座標，位置全交給 Windows，它用階梯式
> 落點（每開一個新視窗往右下挪），開久了下緣就掉到工作列底下。

尺寸固定 `900x720` **永不跳動**。舊版在可選 Sheet 面板展開時把視窗從 650 切到
800，掃描完成的瞬間視窗自己長高 150px。現在靠寬度 900（Sheet 面板 4 欄）與
面板容器 60px 把高度需求壓下來。設定頁的可捲動容器高度 `_TAB3_HEIGHT = 342`
是實測貼齊值——停在這裡 Notebook 維持 393px、log 保有約 10 行；拉到 360 就
頂到 410px、log 掉一行。**改任何一頁的版面後要重量**。

## `logs/app.log` 的語言與格式（2026-09-02 改）

**訊息一律英文。** `launcher.ps1`（PowerShell）與 `main.py` 兩邊寫同一個檔，
靠來源標籤區分；一個任務兩行（起始 `===` 行帶設定、結束行帶結果與耗時）。

```
=== 2026-09-02 22:42:13 Fetch NVDA | GAAP | 10-Q/10-K | max80 ===
[22:42:13] [OK   ] NVDA OK, elapsed 1m 54s
=== 2026-09-02 22:42:14 Compare 3 companies NVDA,MSFT,INTC | quarterly | 2020-2025 | 4 metrics ===
[22:50:26] [FAIL ] Compare FAILED (ReadTimeout), elapsed 8m 12s
```

為什麼英文：① 主控台是 cp950，中文＋符號在 PowerShell/Python 混寫這條路徑上
炸過；② 不跟著介面語言漂，`grep` 永遠是同一組關鍵字；③ 讀者是維護者與 AI。
**畫面上的訊息仍然走 `t()` 跟著介面語言**——兩者分工不同，共用的只有
`format_elapsed()` 產生的 `1m 54s` 這種與語言無關的字串。

⚠ 這與全域規則「log 檔固定用母語言」牴觸，是**本專案的例外**；全域規則檔沒有動。

## 測試分層

| 指令 | 時間 | 測試數 | 用途 |
|------|------|--------|------|
| `python -m pytest -m "not slow"` | ~20-35 秒 | 1288（2026-08-25，B5/H6/G10 後） | Unit tests（每次改 code 後跑） |
| `pytest -m "slow and b1"` | ~12 分鐘 | 24 | B1 overflow live 驗證（8 tickers） |
| `pytest -m "slow and cf_overflow"` | ~5 分鐘 | 15 | CF YTD overflow 驗收（COHR/LITE/AAPL/NVDA/GOOGL） |
| `pytest -m slow` | ~12-31 分鐘 | 58 passed / 7 skipped（2026-08-25） | 完整 live 驗收 |

> ⚠ **slow 紅燈先看是不是逾時。** 2026-08-25 同一天連跑三輪 201 家掃描之後 SEC 開始
> 限流，slow 出現 6 條 `httpx.ReadTimeout`（JPM 5 條 + ARLO 1 條），單獨重跑就全過。
> **失敗訊息裡是 `ReadTimeout` 而不是數字對不上，就不是回歸**——重跑那幾條確認即可。

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
- **Q4 推算（D0-1，2026-08-20）**：`Data_Financials(Q)` 補上原本永遠缺的 Q4 欄。細節見下方「Q4 推算」獨立章節

## Q4 推算（D0-1，2026-08-20）

**模組位置**：`src/fetcher_gaap.py` 的 `_synthesize_q4()` 函式（`_merge_financials()` 定義之後），
由 `fetch_gaap_statements()` 呼叫，是 GAAP fetcher 內部的一步，**不是獨立模組、沒有對外的
CLI 子指令**——外部想拿到含 Q4 的季度資料，一樣走既有的 `fetch_gaap_statements()` /
`cli.py gaap` 路徑即可，Q4 欄位會自動包含在 `Data_Financials(Q)` 裡，不需要另外呼叫。

**背景**：SEC 沒有 Q4 的 10-Q——公司只交 Q1/Q2/Q3，Q4 數字本來要嘛在 10-K 年報裡、要嘛
沒有。`Data_Financials(Q)` 過去因此永遠沒有 Q4 欄，連帶 TTM 類比率（ROE／ROA／FCF per
Share／淨負債EBITDA）湊不到連續四季，多半是空的。

**做法**：有年報（10-K）可用時，用年度值反推單季 Q4：
- IS/CF（流量項）：`Q4 = 年報 FY 值 − Q1 − Q2 − Q3`
- BS（存量項，資產負債表本來就是年底時點數字）：`Q4 = 年報 FY 值`，直接取用不相減

⚠ **現金流量表裡混著時點值**（2026-08-22 修，G12）：`Ending Cash` 是期末現金**餘額**，
不是本期發生額，做「年報 − Q1 − Q2 − Q3」會減成負數。`_CF_POINT_IN_TIME_IDX` 標出這些列，
`_synthesize_q4()` 的 `point_in_time_idx` 參數讓它們直接取年報值。同一個錯誤在
`_build_cf_table()` 的 YTD 拆算也有一份（本季 YTD − 上季 YTD 會把餘額減成變動額），
兩處都要跳過。實測 AAPL 的 `2026-03-28` 從 255,000,000 修成 45,572,000,000。
**日後在 `CF_TEMPLATE` 加時點值列時要記得加進 `_CF_POINT_IN_TIME_IDX`。**

只處理模板列（`IS_TEMPLATE`／`BS_TEMPLATE`／`CF_TEMPLATE` 涵蓋的固定科目），overflow 列
（公司特有科目）Q4 一律留 `None`——季報與年報兩邊 overflow 列的出現順序不保證對齊，沒有
可靠的列對應。`fetch_annual=False`（只抓季報不抓年報）時優雅跳過，不影響既有行為；
NG／segment 表不處理。Excel 上不特別標示 Q4 為推算值（跟既有 Q2/Q3 YTD 拆算的處理方式
一致，見 `Data_CF` 的 YTD 相減邏輯）。

**呼叫順序**（`fetch_gaap_statements()` 內）：年報 IS/BS/CF 表要先建好（`is_ann`／`bs_ann`／
`cf_ann`），季報表才能用它們反推 Q4，所以年報的 build 呼叫被移到季報之前，但輸出到
`tables` 列表的順序不變（`Data_Financials(Q)` 仍先於 `Data_Financials(Y)`）。

**測試**：`tests/test_fetcher_gaap.py` 的 `_synthesize_q4` 單元測試（6 個，涵蓋相減/直接取值/
季度不齊全跳過/不覆蓋既有值/無年報時原樣返回/overflow 留空）+ 1 個 `fetch_gaap_statements`
整合測試。全套 156 個 fetcher_gaap 測試、848 個非 live 測試、906 個含真連線 SEC EDGAR 的
live 測試（NVDA/AVGO/PLTR/ARLO 等）皆通過，0 failed。

**已驗證（2026-08-20）**：實測 NVDA/AVGO/PLTR 真連線 SEC EDGAR，`ratios.py::build_ratio_table()`
算出的 ROE (%)／ROA (%) 從第一個補上的 Q4 期開始正常出現非 None 值（三家皆 10~11/16 期），
不再是原本 D0-1 記錄的「整列全空」。附帶發現 FCF per Share 三家仍有缺口，但根因是
`Shares Outstanding` 資料本身缺期（PLTR 屬於下方「多股別公司抓不到流通股數」已知限制；
NVDA/AVGO 缺口原因待查），與本次 Q4 修復無關，不在本次範圍內處理。

## 期間標籤與日曆季（2026-08-22 重整）

整套系統有**三種期間概念**，混用會出錯，所以刻意用不同的名字與寫法：

| 概念 | 長相 | 意義 | 誰在用 |
|---|---|---|---|
| 財季 | `FY2026Q2` / `FY2026FQ2` | 公司自己財年裡的第幾季 | `Data_Financials` 第 1、3 列 |
| 結算季 | `2025Q3` | 這一季**結束**在哪個日曆季 | `Data_Financials` 第 4 列 |
| 對齊季 | `2025Q2` | 這一季的**多數天數**落在哪個日曆季 | 跨公司比較的欄位 |

### 財季編號一律由期末日反推

**不採信 edgartools 欄名裡的 `(Qn)`。** 那個標記對 52/53 週財年制的公司會標錯，
而且同一家公司相鄰兩季會撞號：

    NVDA 2010-08-01 → (Q3)   實際 FY2011Q2   ← 與下一行撞號
    NVDA 2010-10-31 → (Q3)   實際 FY2011Q3
    INTC 2023-07-01 → (Q3)   實際 2023Q2     ← 與下一行撞號

filings 是新到舊排序，`_build_*_table()` 的 dedup（`if label in periods: continue`）
會把撞號的舊那一季**靜默丟掉**；少了 Q1，`_synthesize_q4()` 也就合成不出來。
更嚴重的是沒被丟掉的資料也貼錯格。

現在 `(Qn)` 只用來分辨「季度欄 vs 年度欄」，編號一律交給
`fiscal_input.fiscal_quarter_of()` 用實際期末日算。

### 8-K 季度標籤：零下載規則（B5，2026-08-25）

Item 2.02 8-K 的 `period_of_report` 是**發布日**不是財期結束日，直接取它的日曆季
實測 31.5% 連年份都錯（`--years` 就篩錯）。修法是在**不下載任何文件**的前提下算對：

```
fiscal_input.quarter_label_from_announcement(發布日, fiscal_year_end MMDD, tol=21)
    候選季末 = MMDD 往前推 0/3/6/9 個月（該月沒那天 → 當月最後一天）
    取「不晚於 發布日 + tol」的最新候選，再套 fiscal_quarter_of()
```

- **界線變了**：`_list_earnings_filings()` 從「純 listing metadata」變成
  「listing metadata + **一次 company 層級查詢**」。仍然零文件下載，
  多的成本是每個 ticker 一次 submissions 請求（`Company.fiscal_year_end`），
  而財年結束月本來就要查
- 傳的是**完整 MMDD**（`"0703"`）不是月份。只給月份會退化到 79.8%
- 算不出來（沒有 MMDD／日期畸形／sanity check 沒過）→ **逐份**退回舊算法，
  不讓 EDGAR 少一個欄位變成整批失敗
- `_recover_missing_quarters()` 也吃同一個 MMDD：缺季比對的兩邊必須是同一個
  命名空間，否則整組都會被判成「缺季」
- 200 份實測，157 份基準可信全部與下載後算的 `fiscal_label` 一致（100%）
- **`fiscal_label` 沒動**，它仍然最準；`cli.py` 會把兩者比對後吐
  `label_agrees_with_fiscal_label`
- **已知風險**：EDGAR 只給現在的 `fiscal_year_end`，公司改過財年的舊申報會整段
  偏掉，目前沒有對策（交接文件原本提的 0~70 天 sanity check 是恆真式，攔不到）。
  細節見 `docs/8k-period-off-by-one.md`

### 結算季 vs 對齊季：一份實作、兩個具名基準點

`fiscal_input.calendar_quarter_of(period_end, *, basis)`：

- `basis="end"` 內縮 **15 天** → 回到該季最後一個月，吃掉 52/53 週的月底漂移
- `basis="span"` 內縮 **45 天** → 回到該季中點

**`basis` 刻意不給預設值**，不給就噴 `TypeError`。強迫每個呼叫端表態——以後只有
一個地方會算錯，而且每個使用點都看得出它要哪種語意。

為什麼跨公司要用期中點：NVDA 7 月底結束那季要跟 AMD/INTC 6 月底那季擺同一欄
（同一波財報），用期末日會把它推到跟 AMD 9 月那季同欄。**期中點是離日曆季邊界
最遠的位置，最穩**；期初日反而最不穩（13 週季的起訖日都落在邊界附近）。

**獨立驗證**：SEC 官方在每筆 fact 上標的 `frame`（如 `CY2025Q2`）是它自己的日曆季
正規化。拿它跟 `basis="span"` 比，24 家、59,564 筆**零例外全數一致**。

Excel 公式（`fiscal_input` 的 `_date_expr()` 等）是 Python 那份的規格複寫，
維持 `basis="end"` 的邏輯，兩邊註解互相指到對方。

## 缺漏判斷（`data_quality.py`，2026-08-22）

Index 上的「資料完整度」區塊。**取代**舊的「9 個關鍵列、最近 4 期全空才算缺」
——那個判準寬到形同虛設（NVDA 顯示 `9/9 ✓`，實際 95 個欄位有 27 個幾乎全空）。

四個判斷，可信度由高到低，**全部不需要「同業基準表」**：

| | 判斷 | 做法 | 誤判率 |
|---|---|---|---|
| A | 季度斷層 | `round(天數差 / 91) - 1` | 0 |
| D | 整欄稀疏 | 一欄有值的模板列 < 50% | 0 |
| B | 中間有洞 | 首末有值之間仍有空格 | 0 |
| C | 整列全空且矛盾 | 空白，但相關欄位顯示它應該要有 | 低 |

**為什麼不用同業普及率**：那會讓「公司真的沒有某個科目」被永遠標紅。C 改用
**同一家公司的相關欄位互相驗證**（`_COHERENCE` 表，全是會計上必然的關係）：
有負債餘額就會有借還款現金流；反過來若負債類欄位全空，沒有借還款紀錄是一致的，
不標紅。

**`_COHERENCE_EXCUSES`：空白的正當理由，優先於 `_COHERENCE`**（2026-08-23，H3）。
只有 `_COHERENCE` 一張表不夠——它只問「有沒有理由該有值」，不問「有沒有理由
可以空白」。`Current Portion of LT Debt` 就是這樣被誤判 25 家：多數美國公司的
資產負債表表面只有**一條**流動借款列（`us-gaap:DebtCurrent`），一年內到期的長期
負債併在裡面，而那條已經進了 `Short-term Debt`——資訊沒掉，不該標紅。所以
「Short-term Debt 有值」就是這一列空白的正當理由。改完 25 家 → 3 家。

加新規則到 `_COHERENCE` 之前先問一句：**這一列空白時，數字有沒有可能是被抓進
隔壁那一列了？** 有的話要一起加 excuse，不然就是在製造假警報。

三個實測踩出來、改動時不要弄掉的細節：

1. **A 不能用固定門檻**——52 家 1,482 對相鄰期間裡，111~150 天的 16 筆全部是
   COSTCO（16 週的第四季）。固定門檻會把它們全部誤判成缺季。
   **`missing_quarters()` 同時是 G6「缺季留空白欄」的判定來源**：單一公司
   `fetcher_gaap._with_gap_columns()` 與跨公司 `comparison_writer._fill_period_gaps()`
   都呼叫它，不要再寫第二份公式
2. **B 只看「首末有值之間」**——`Operating Lease ROU Assets` 只有 28/67 期是因為
   ASC 842 從 2019 才適用，前後空白不是漏抓
3. **B 要排除 overflow 列與整欄稀疏的期間**——不排除的話 NVDA 報 85 列有洞，
   排除後 14 列。合成 Q4 失敗時每個流量列都會在那一期留一個洞，那是**一個期間
   問題**不是 40 個列問題

**補出來的空白欄**（G6）：抓不到的季度不再整欄消失，保留欄位、內容全空。
單一公司那條線第 5 列退回由財季標籤反推的年月（`2025-06`，非完整 ISO，
`fiscal_input._apply_to_sheet()` 會保留靜態標籤不套公式）；跨公司那條線的期末
結算日列整格留空。因此同一件事會同時被判定 A（季度斷層，看真實期末日）與
判定 D（整欄稀疏，看新的空白欄）報出來——兩邊都抓得到，不會漏。

**模板不適用**：稀疏欄超過一半期數時，Index 直接顯示「這個模板不適用這家公司」，
不列出每一欄。金融股與 REIT 全數觸發。

基線資料：`docs/template-coverage-baseline-*.md`，用
`scripts/gen_template_coverage_baseline.py` 重跑（不打網路）。

## 解析快取（`_parse_cache_scope()`，2026-08-22）

IS/BS/CF/segments 四個 build pass 各自對**同一批 filing** 重新解析一次。實測
ARLO 25 份 filing 共 66 秒，`_filing_obj` 被呼叫 96 次（每份 3.8 次），其中
XBRL 解析 19.9 秒、`to_dataframe` 28.4 秒。**edgartools 不會跨呼叫快取**——
同一支 ticker 在同一個 process 連跑兩次是 64.5s vs 67.3s，完全沒變快。

`_parse_cache_scope()` 涵蓋一次抓取，`_filing_obj()` 與 `_financials_of()`
在範圍內以 accession number 為鍵快取。實測 64.5s → 33.6s（冷）／44.0s → 32.4s
（全熱），**輸出零變化**（逐格比對 5,678 格，0 格不同）。

兩個容易弄錯的地方：

- **快取的生命週期只能是一次抓取**。跨 ticker 殘留會吃到別家資料，跨執行殘留
  會拿到過期申報
- **範圍要包在 `fetch_gaap_statements()` 外層**（本體拆成 `_fetch_gaap_impl()`），
  不是放在 `_ledger() is None` 那個分支裡——`main.py`／`cli.py` 會自己先開
  `collect_gaps()`，那條路不會走到遞迴

## 本地 filing 快取（`filing_cache.py`，2026-09-03）

**卡在哪一層**：解析層與比對層之間。`_filing_obj()`（`fetcher_gaap.py`）解出
edgartools 的 income statement / balance sheet / cashflow statement 三張
DataFrame 之後，原封不動存進 `%APPDATA%\SEC Financial Tools\filing_cache\`；
`IS/BS/CF_TEMPLATE` 那套科目比對規則、hint regex、Q4 合成邏輯，永遠在快取
**之上**即時重跑，不進快取檔。所以以後改比對規則、加比率、調模板都不會讓
快取失效——**但 edgartools 升版會**，靠每份快取檔裡的 `edgartools_version`
欄位擋（見下方四道閘）。這跟同一節上面的「解析快取（`_parse_cache_scope()`）」
是兩層不同的東西：那層只活在一次執行的記憶體裡（G9），這層落地跨執行有效。

**儲存位置與檔案**：一個 ticker 一個資料夾（`filing_cache/ARLO/`），檔名是
SEC accession number（`0000866787-25-000123.json`）。**`<accession>.json`
是唯一的落地狀態**——查快取一律直接問這個檔案存不存在、內容能不能過四道閘，
不維護任何索引。GUI 想知道「哪些公司有快取、各佔多大」也不例外：
`list_cached_tickers()` 直接掃 `filing_cache/` 底下有哪些子資料夾、逐一
`stat()` 加總，不開額外的索引檔——公司數量最多幾十家，掃資料夾的成本可以
忽略，比維護一份索引跟磁碟隨時同步的心智成本低得多。

**四道閘**（`load_filing()`，任一沒過就當無快取、照舊打 SEC 重抓，不拋例外）：
1. JSON 能解析（檔案沒被中斷寫入或手動改壞）
2. `schema_version` 跟現在的 `filing_cache.SCHEMA_VERSION` 相符
3. `cik` 跟這次抓取的公司相符（防 ticker 換手撞名）
4. `edgartools_version` 跟現在裝的套件版本相符（`importlib.metadata.version()`
   讀出來，讀不到就整次不用快取）

**負向快取 vs 網路失敗**：`has_financials: false` 是負向快取，代表這份
filing 解析**成功**但沒有 XBRL financials（多半是 pre-XBRL 的舊申報）——
這種情況合法且穩定，快取起來下次不用重解。**網路失敗絕對不會走到這裡**：
`_filing_obj()` 下載失敗會直接往外拋，交給既有的 D11-B 缺漏帳本，下次照樣
重試，不留任何快取（正向負向都不留）。分不清楚「真的沒有 financials」還是
「取用 `.financials` 時炸出 `AttributeError`」的情況（XBRL 殘缺或格式異常）
也一律不寫快取，留給下次重試——寫錯負向快取的代價是把一期本來可能修好
再抓到的資料永久判死刑，遠比多重試幾次貴。

**替身物件的兩條隱性規則**（`_CachedFiling` / `_CachedFinancials` /
`_CachedStatement`）：命中快取時回傳的不是真正的 edgartools filing 物件，
是只實作 `.financials.income_statement().to_dataframe()` 這條鏈的替身。
① **刻意不定義 `__getattr__`**——存取未實作的屬性一律照 Python 預設拋
`AttributeError`，不兜底回 `None`。如果兜底，以後某個 builder 新用到
filing 物件的其他屬性，快取命中的路徑會安靜地把整份 filing 當成沒資料，
清快取重跑卻是好的——這種 bug 極難查，寧可讓它馬上炸出來。
② **`None` 不等於空 DataFrame**：`payload_to_df(None)` 回傳 `None`（代表
「這張表本來就不存在」），跟真的解出一張沒有列的空表是兩種狀態，下游判斷
`if stmt is None` 的邏輯才不會被快取路徑改變語意。

**⚠ edgartools 自己也有一層持久化 HTTP 快取，這個專案的清除動作碰不到它**：
`~/.edgar/_tcache`（Hishel-File，規則對 `/Archives/edgar/data` 是「快取到
天荒地老」）跟這裡的 `filing_cache.py` 是完全獨立的兩層。
`filing_cache.clear_ticker()` 只清這個專案自己 `%APPDATA%` 底下那份，
**不會**動到 `~/.edgar/_tcache`——如果只清這邊就量「冷跑」，量到的其實是
「本專案快取沒有、但 edgartools 自己的 HTTP 快取可能還在」的混合狀態，不是
真正對 SEC 重新握手一次。下面的數字是把兩層都清乾淨（`filing_cache.clear_ticker`
+ 手動刪 `~/.edgar/_tcache`）之後才量的，第一次量測（2026-09-03 稍早）踩過
這個坑，數字已作廢重量。

**這次量到的數字**（ARLO、GAAP、10-Q+10-K、`--max-filings` 預設 80，
33 份 filing，2026-09-03，SEC EDGAR 實測，重量版）：

| 跑法 | elapsed | 快取命中 |
|---|---|---|
| 冷跑（兩層快取都清空後第一次抓）| 55~110s | 0/33 |
| 熱跑（緊接著再抓一次，兩層都不清）| 10~17s | 33/33 |

⚠ **「快幾倍」不是一個穩定的數字，不要引用單一倍數。** 2026-09-03 當天在
同一台機器上量了三輪，冷跑落在 55.12 / 55.96 / 56.20 / **110.45**s，熱跑落在
7.89 / 10.17 / 11.66 / 12.69 / 13.39 / 15.54 / 15.69 / 16.01 / 17.15s，
換算出來的倍數從 **3.5 倍到 9.5 倍**都有。變異來源是兩邊各自的：冷跑取決於
SEC EDGAR 當下的反應速度（同樣把兩層快取清乾淨，55s 與 110s 都量到過），
熱跑取決於本機 CPU 當下的負載。**同一次連續執行內量出來的比值才有意義**
（例如 110.45s → 中位數 11.66s = 9.5 倍），跨時段互相比的數字不可靠。

**穩定、可引用的是這個**：快取確實把「下載 + XBRL 解析 + `to_dataframe()`」
那一段整個消掉了，`cache 0/33 → 33/33` 是精確的計數，不是估算。想跟人講
效益時講這個，不要講倍數。

**golden 逐格比對 0 格不同**
（`scripts/excel_golden.py check`，exit code 0，覆蓋 ARLO 全部 sheet；用的是
本節最早那組冷/熱跑產出的 xlsx 存的基準，比對的是輸出內容而非耗時，**不受
上面這次重量影響**，結論仍然有效，未重跑）——快取只改變耗時，沒有改變任何
輸出內容。

**冷跑變慢的取捨（已發生，低於但接近 +30% 門檻）**：miss 時
`_save_to_disk_cache()` 會為了落檔多呼叫一次三張表的 `to_dataframe()`，
builder 組表時還要再呼叫一次——同一份 filing 的 `to_dataframe()` 冷跑時被叫
兩次。第一次量這個數字時只清了本專案的快取、沒清 `~/.edgar/_tcache`，
且 `master` 與這個分支的兩輪比較沒有交錯執行順序，量到 +56%——**這個數字
已經作廢**。重量方法：`master`／這個分支交替執行（M→F→M→F），**每一輪都把
兩層快取清乾淨**，避免 SEC 端當天流量波動只影響其中一邊：

| 順序 | 分支 | elapsed | 備註 |
|---|---|---|---|
| 1 | master | 45.15s | |
| 2 | 這個分支 | 55.96s | cache 0/33 |
| 3 | master | 43.01s | |
| 4 | 這個分支 | 56.20s | cache 0/33 |

master 平均 44.08s（45.15s / 43.01s），這個分支冷跑平均 56.08s（55.96s /
56.20s），**慢了約 +27%**——比原先估計的 +34% 略低，也低於 +30% 的記錄門檻，
但因為前一版數字（+56%）明顯是量測方法有誤、且新數字已經很接近門檻，這裡
還是把交替測試的原始數字整組留下來，供以後複查。這是刻意的取捨：miss 時若
改回傳替身物件重用那三張表可以省掉這筆重複解析成本，但那樣「冷跑 vs 熱跑」
逐格比對就變成同一條程式碼路徑跟自己比，驗收會失去意義，所以不預設採用。
熱跑完全不受影響（這筆重複解析只發生在 miss 的那一份上）。⚠ 但這個 +27%
本身也繼承了上面那條變異問題：交替測試的四趟都在同一個時段跑完，所以彼此
可比；跟別的時段量到的冷跑（例如同日稍晚的 110.45s）不能混著算。

**`pytest -m slow`（65 條，打真實 SEC，2026-09-03）**：58 passed、7 skipped、
0 failed，耗時 12m 29s。這次改動的是四個 builder 共用的熱路徑，這層測試沒有
發現行為變化。

**新一期財報會被抓到——用等價路徑實測過（2026-09-03）**：機制上，
`_list_filings()` 每次抓取都重新對 SEC 查最新清單（不受快取影響，見下方
「不涵蓋」），新申報的 accession number 在快取資料夾裡找不到對應
`<accession>.json`，四道閘第一關就過不了，自然落回正常下載流程並寫入新
快取檔——不需要額外的「過期」判斷。

真實情境（等某家公司剛好發新一期）等不到，但**程式走的是同一條路徑**：
「清單裡有一份 accession，本機沒有對應檔案」。所以直接把 ARLO 快取裡
**最新那一份**（`0001736946-26-000060`，2026-08-06 的 10-Q）刪掉再抓一次，
實測結果：

- log 記到 `cache 32/33`——33 份裡只有被刪的那一份是 miss，其餘全部命中，
  **沒有整批重抓**
- 那一份被重新解析並寫回快取，總份數回到 33
- 重新解析出來的 JSON 與原本只差 `cached_at` 時間戳，`scripts/excel_golden.py`
  逐格比對 **IDENTICAL**

也就是說「補抓缺的那一份、不動已有的」這個行為是驗過的。唯一還沒被真實
新申報覆蓋到的，只剩「SEC 清單本身會不會即時出現新 filing」，而那一段
完全不經過快取（`_list_filings()` 每次都打網路）。

**不涵蓋什麼**（這幾條完全不受這次改動影響，每次都照舊打網路）：
- **Non-GAAP**：`fetcher_nongaap.py` 有自己獨立的 `nongaap_cache.json`，
  跟這裡的 `filing_cache.py` 是兩套不同的快取
- **流通股數**：`company.get_facts()`（BVPS / FCF per Share / 流通股數 YoY
  用到的來源）不在這層快取範圍內
- **`_list_filings()` 的清單查詢**：每次抓取都重新問 SEC「這家公司有哪些
  filing」，不受這層快取影響（也不受 edgartools 自己的 HTTP 快取影響——
  `_get_cache_rules()` 只把 `/submissions` 這類清單端點快取 30 秒）。這是
  新一期財報一定抓得到的原因

### ⚠ 熱跑剩下的時間九成不是網路，是比對層

設計文件與本節初版都寫「熱跑仍然要花幾秒，主因是查清單等網路往返」。
**2026-09-03 實測推翻了這個說法**：對 ARLO 全命中的一趟抓取逐段計時，三次
量測結果一致：

| 區段 | 耗時 | 佔比 |
|---|---|---|
| 建 `Company` 物件（網路）| 0.04~0.72s | <5% |
| `_list_filings()` × 2（網路）| 0.40~1.04s | 3~6% |
| 查流通股數（網路）| 0.65~1.01s | 4~6% |
| 讀快取檔 JSON（33 次）＋ 還原 DataFrame（224 次）| 約 0.95s | 6% |
| **比對層（模板配對／Q4 合成／比率／segments，純本機 CPU）** | **約 14.2s** | **90%** |

也就是說快取已經把它能吃的那一段吃乾淨了，**剩下的瓶頸整個換了一個位置**：
從網路搬到本機 CPU。這件事有兩個後果，記在這裡免得下次找錯方向：

1. **再加任何一層快取都不會讓熱跑變快**——`_filing_obj()` 那條線的成本已經
   降到 0.1s 等級。要再快只能優化比對層本身（`IS/BS/CF_TEMPLATE` 的逐列
   配對、`_synthesize_q4()`、`segments`），那是完全不同的工程
2. **`payload_to_df()` 被呼叫 224 次而不是 33 次**——四個 builder 各自對同一
   份 filing 的同一張表重新呼叫 `to_dataframe()`，替身物件每次都重新從 JSON
   還原一次 DataFrame。真物件在 G9 記憶體快取下也是每次重算，所以行為一致、
   不影響正確性；但如果哪天要優化，在 `_CachedStatement` 上加一層 memo 是
   最便宜的一刀（約 0.85s，佔 6%）
- **修正案（10-Q/A、10-K/A）**：`_list_filings(amendments=False)` 現況本來
  就不抓，跟這個快取無關（獨立議題，見 `docs/TODO.md`）

## 跨公司比較（`comparison.py` + `comparison_writer.py`）

`build_comparison()` 對每個 ticker 呼叫一次 `fetch_gaap_statements()`，重組成
`{指標: {ticker: {日曆季: 值}}}`。單一公司失敗只記成 `CompanyFetchError`，
不中斷其他家。

**欄位鍵是對齊季（`2025Q2`），不是財季。** 財年結束月不同的公司，同一個財季標籤
指的不是同一段時間——NVDA 的 `FY2026Q2` 結束在 2025-07-27，AMD 的結束在
2026-06-27，硬擺同一欄會讓 NVDA 整條線在日曆時間上偏移約一年。

輸出五種 sheet：`Compare_Data`（唯一的原始資料表，每個指標一個區塊）、
`Notes`（說明，見下）、`Snapshot`（公式驅動的單一時間點切面）、
`Snapshot_Manual`（空白供人工貼值）、`Chart_<指標>`（每個指標一張折線圖）。

**`Compare_Data` 最上方是「日曆季 ↔ 財季」對應表**（G2）：一格
`FY2026Q2 (0727)` 講完「這一欄對這家公司是哪一財季、期末日幾號」，下面的指標
區塊只給日曆季。用對應表而不是「每家公司的財年開始月份」是因為每一格都逐期
從實際期末日算，公司改過財年那一欄自己就對，不需要例外處理。**插入這個區塊會
把所有列號往下推**，`write_snapshot_sheets()` 與 `write_chart_sheets()` 都吃
`block_ranges`——測試一律回頭查列號在 `Compare_Data` 上是誰，不比對列號常數。

**`Notes` 是資料驅動的**（G7）：`NOTE_ITEMS` 一條一行
`(標題鍵, 內文鍵, 判定函式)`，每條帶「這份檔案是否真的踩到」的勾選與實際情況。
新增一條只要加一行 + 四個 locale 各加兩條，文字不寫死在版面程式裡。其中「本檔案
缺少的公司」讀的是 `ComparisonResult.failures`——抓取失敗原本只寫進 GUI log，
檔案裡完全看不出來。

**哪些公司、哪些欄位會出現在表上由 `visible_layout()` 一處決定**，對應表、各指標
區塊與 `Notes` 三邊共用，不各算一次（會靜默不一致）。

**期末結算日列取同欄各公司最晚的那個日期**——同一個對齊季各家期末日不同，
Snapshot 用它做「不晚於 B1」的判斷，取早的會顯示還沒結算完的數字。

openpyxl 畫圖有一類反覆踩到的坑：**它不寫的元素，Excel 會當成未定義狀態而不是
預設值**。已知要明講的有 `<c:delete val="0"/>`（兩軸）、`<c:overlay val="0"/>`
（圖例 + 三個標題）、`<c:layout/>`（三個標題）、`axPos`／`tickLblPos`／`crosses`，
以及類別軸要手動換成 `strRef`（`set_categories()` 永遠寫 `numRef`）。

## companyfacts 平行路徑（`fetcher_facts.py`，G11 spike）

**尚未接上主流程。** 這是現行「逐份解 filing」的平行替代品，建好是為了先產出
逐格比對報告，讓 CTH 看數據決定要不要切換。完整報告見
`docs/superpowers/report-2026-08-22-g11-companyfacts.md`。

一句話：`data.sec.gov/api/xbrl/companyfacts/CIK##########.json`，**一家公司一個
request、0.34 秒**（現行每家 7.5 分鐘）。每筆 fact 自帶 `start`/`end`，所以不用
猜期間、不用猜哪欄是 YTD、Q4 常常直接有。

**它拿不到的**（結構限制，不是待辦）：沒有任何維度欄位 → `Data_Segments` 這條路
拿不到；沒有 presentation linkbase → 沒有公司自報的原文標籤。所以真要切換是
**混合架構**：模板列走 facts，segments 仍解 filing。

## Known Issues（已知限制，暫不修）

- **Investment Proceeds**：XBRL 沒有單一加總行，取 first match
- **金融股（GS/JPM/BAC/SCHW）與 REIT（PLD）**：現行模板 BS/IS 大量空白，待獨立模板。2026-08-22 有量化證據——`data_quality` 對這五家全數判定「模板不適用」（稀疏欄佔 90~100%），Index 上會直接這樣顯示。逐列覆蓋率見 `docs/template-coverage-baseline-*.md`
- **NG 分類誤判**：keyword-based 分類，label 含 "excluding" 的 GAAP 行可能誤進 NG sheet（可接受方向）
- **Data_EPS_Recon 從未產生**：edgartools `eps_reconciliation` 對 NVDA/AAPL/MSFT 均回傳 None，非 XBRL-tagged 公司無解；待 edgartools 改善或改用 AI 解析方案
- **模板列覆蓋率只有 40/97**（2026-08-22，52 家實測）：達到「≥45 家有值且填滿率 >90%」的只有 40 列。系統性問題見 `docs/TODO.md` H3——`Current Portion of LT Debt` 25/52 家被判矛盾、`Shares Outstanding` 43/52 家有洞。**部分列（`Accrued Compensation` 等）改 concept 名字救不了**——那些 filing 的報表表面根本沒有那一列，公司把它放在附註，只有 companyfacts 拿得到
- **多股別公司抓不到流通股數**：PLTR／GOOGL／META 有 Class A/B/C，封面頁的 `dei:EntityCommonStockSharesOutstanding` 按股別分開標，`company.get_facts()` 取不到。連帶 BVPS、FCF per Share、流通股數 YoY 空白

## 本地財報資料庫（`local_db.py`，2026-09-04，TODO J1-J4）

上一節的 `filing_cache.py` 管的是「一份 filing 怎麼存、怎麼讀」。
`local_db.py` 是疊在它上面的**狀態層**，補三塊「狀態與體驗」，
**不改儲存格式、不動 `fetcher_gaap` 的抓取迴圈一行**。
設計書：`docs/superpowers/specs/2026-09-04-local-filing-db-design.md`。

**三份清單，三種語意，刻意不合併**：

| | 快取內容 | 更新名單（`config["local_db_tickers"]`）| `watchlist` |
|---|---|---|---|
| 怎麼來的 | 掃目錄得出的**事實** | 使用者維護的**意圖** | 使用者維護 |
| 內容 | 所有抓過的（含 Tab 1 一次性查詢）| 要保持新鮮的（可含還沒抓過的）| 批次產 Excel 的對象 |
| 誰改 | 抓取／清除自動改 | 只有使用者 | 只有使用者 |

合併會壞在兩處：併進 `watchlist` → Tab 2 一按產 201 份 Excel；
改用「快取裡有的」→「未來把全部抓好」做不到，因為還沒抓的公司永遠不會被抓。
三者可重疊也可不重疊，**不強制包含關係**。

**`_meta.json`**：一家一份，放該公司資料夾（`filing_cache/AAPL/_meta.json`）。
不做全域單一檔——分開才不會在多視窗同時寫時互相蓋掉。分 form（10-Q／10-K）
記 `count`／`oldest`／`newest`／`reached_bottom`，因為 `max_filings` 與
`max_annual_filings` 是兩個獨立上限，一家公司可能 10-K 到底了、10-Q 還沒，
**合記會誤判成整家到底，然後永遠不再往下挖**。

**原則：掃目錄為準，meta 只是快照。** 讀 meta 時比對 `file_count` 與實際的
目錄列舉結果，不符（或檔案壞掉、schema 不符）就當場重建並寫回。
比對用的是**目錄列舉、不讀檔內容**，很便宜——201 家若每次都要讀 881 個 JSON
才畫得出 GUI 清單會卡住。meta 刪掉、寫壞、跟目錄不同步，功能都照樣正確，
只是慢一點，所以 `filing_cache.py` 那條「不維護額外索引檔」的原則沒有被破壞。

**「到底」怎麼判定（`derive_reached_bottom()`）**：抓取迴圈有三個停止條件
（撞到 `max_filings`、撞到 `_XBRL_CUTOFF`＝2008-01-01、清單用完）。
直覺做法是讓 builder 回報是哪一個停的，**但那要穿過 3 個 builder 與 8 個
`_current_q_col()` 呼叫點，就是 TODO G13 (a) 那個坑**。改成在迴圈外面推導，
零改動：

```
available = _list_filings(公司, form) 裡 filing_date >= 2008-01-01 的
cached    = 該資料夾裡的 accession 集合

reached_bottom = "no_more_filings"  cached ⊇ available，且原始清單沒有 2008 前的
               = "xbrl_cutoff"      cached ⊇ available，且原始清單有被 2008 擋掉的
               = null               否則（還沒抓完，下次要繼續挖）
```

`_list_filings()` 本來就回傳完整清單，所以這些資訊在「更新本地庫」跑的時候
順手就有，不必額外連網。日期解析不出來的一律**當成在窗內**（＝判 null）——
誤判「還沒到底」只是多查一次清單，誤判「到底」會讓那家公司永遠不再往下挖，
而且完全沒有症狀。

**「更新本地庫」（`update_local_db()`）**：對象是更新名單，深度一律拓到底，
**不產 Excel**（要 Excel 走既有的 Tab 2 批次）。每家四步：

1. `_list_filings()` 拿 10-Q／10-K 完整清單 ← 一次網路，很便宜
2. 跟快取現況比對：兩個 form 都到底、又沒有新 filing → **整家跳過**
3. `fetch_gaap_statements(max_filings=200, max_annual_filings=50)`，
   結果**直接丟棄**——要的是它的副作用（把快取填滿）。200/50 只是「大到不會是
   它先喊停」的餘裕值，實際由 `_XBRL_CUTOFF` 或清單用完停止
4. 用步驟 1 的完整清單重算 `_meta.json`

步驟 2 是「不要每次全部重抓」的具體實現。**實測（2026-09-04，AAPL／ARLO／META）**：
第一輪 49.4s（AAPL、ARLO 整家跳過，META 新增 27 份），**第二輪 1.0s、三家全跳過、
零下載**。

**抓取速率有兩個數字，不要混用**：對 `~/.edgar/_tcache`（edgartools 自己那層
持久化 HTTP 快取）已經熱的公司是 ≈1.8 s/份；**真的沒抓過的公司是 2.8 s/份**
（2026-09-04 連續抓 15 家新公司，取中段 900 秒的窗量到 321 份）。
估「重抓要幾小時」一律用冷跑那個。上一節記過同一個坑——只清本專案的快取
就量「冷跑」，量到的是混合狀態，第一次的快取效能量測整組作廢過一次。

單一公司失敗**不中斷整體**（比照 `comparison.py` 的 `CompanyFetchError` 原則：
公司層級跳過，跟同一家公司內部的科目缺漏是兩回事）。沿用 `collect_gaps()`，
跑完列出有抓取缺漏的公司——這對應 TODO D11：連續大量抓取時 SEC 會偶發失敗、
**靜默少格**，那幾家之後單獨重跑即可（第二輪從本地快取讀已成功的部分）。
中斷不會白費：`save_filing()` 是逐份即時落檔的。

**版本鎖（J4）**：`requirements.txt` 鎖成 `edgartools==5.29.0`。不鎖的話任何人
重跑一次 `pip install -r requirements.txt` 都可能在沒打算升級的情況下讓整個本地庫
失效。升級是在命令列發生的、不會在 GUI 裡發生，所以偵測點是**下次啟動**
（`main._warn_if_edgartools_changed()`），明示「N 家 M 份將失效、預估重抓 H 小時」
並附回退指令。**刻意不提供「照用舊快取」**——明知可能帶著舊 parser 的解析 bug
還拿來做投資判斷，換到的只是省一晚。

**⚠ `launcher.ps1` 每次啟動都會跑一次 `uv pip install -r requirements.txt`**
（`launcher.ps1:168`，venv 已存在的那條路徑也會跑）。加上版本鎖之後這變成一件
好事：走啟動器的使用者，edgartools 版本會被**每次啟動自動拉回 5.29.0**，
本地庫不會因為版本漂移而失效。所以 J4 那個提醒對話框實際上是給
「自己在命令列 `pip install` 過」的人看的——這兩塊是互補的，不是重複。

**刻意不做（YAGNI，設計書第十節，是硬性的）**：不存原始 XBRL（容量 1.4 GB → 42 GB）、
不做容量上限／自動淘汰（自動刪資料違反「抓過不用重抓」的核心承諾）、
不做「版本不符但照用」、不做比 filing 更細的增量（accession 已是最細粒度）、
不動 `fetcher_gaap` 的抓取迴圈。
