# 本地 filing 快取 設計

> 2026-09-03 · 對應 `docs/TODO.md`（待補條目）
> 狀態：設計已與 CTH 逐段確認（Q&A 形式），另一個 session 做過兩輪 spike 審查
> （第一輪：版本漂移、負向快取、manifest 定位、GUI 版面、寫入時機；第二輪：
> 快取命中時 `_filing_obj()` 的替身物件介面、dtype 存檔、edgartools 版本
> 取得方式、原子寫入、避免雙重編碼、log 記快取命中數），已全部併入本文件，
> 待寫實作計畫

## 目標

抓取一家公司的財務資料（Tab1 單一公司、批次抓取、跨公司比較共用同一個入口
`fetcher_gaap.fetch_gaap_statements()`）目前每次都要重新對 SEC 打 20 年份的
filing、逐份解析。CTH 的原話：

> 「我每次要抓新公司都要去重新做抓取 XBRL＋解析，我們能不能已經有解析好的
> 版本（如果以前有做）然後就存放在電腦內，就不用每次都解析 20 年的資料？
> 已經在本地有資料，看未來要運算、新增年度就會比較快。」

做一個**只在本機有效**的持久化快取：第一次抓某家公司會存一份「解析好的原始
資料」在本機，之後不管是 Tab1 重抓、批次抓、還是跨公司比較要同一家公司，
都優先讀本機這份，只在 SEC 有新 filing 時才補抓那幾份新的——不用每次都
重新掃過去 20 年的歷史。

## 不做的事

- **不做 SQLite 或任何資料庫依賴**。維持資料夾＋多個 JSON 檔（CTH 已選定），
  跟現有 `config.json`／`ticker_overrides.json` 同一套慣例，不多引入技術棧。
- **不做自動容量上限或到期清理**。CTH 決定手動清就好——GUI 給「清除某家公司
  快取」的按鈕即可，不用自動 LRU、不用天數上限。
- **不做「快取直接存最終比對完的表格」**。快取存在比對邏輯**之前**那一層
  （見下方「架構」），比對規則（Capex/Revenue 那種科目對應 hint）以後怎麼改
  都不會讓快取失效——這是這份設計最核心的取捨。
- **不改抓取的底層機制**（還是逐份 filing 打 SEC、逐份解析），不換成 SEC 的
  單一公司整批 XBRL 端點。那是更大範圍的改動，這次不做，留在「未來可能」。
- **不影響 Tab1/Tab2/跨公司比較目前的行為或輸出格式**。快取是 `fetch_gaap_statements()`
  內部的加速層，呼叫端（Tab1/批次/`comparison.py`）完全不用改。

---

## 一、架構：快取卡在哪一層

現有抓取分兩層：

1. **解析層**：`fetcher_gaap._filing_obj(filing)` 對單一份 filing 打 SEC、
   解析出這份 filing 裡有哪些 XBRL 科目、對應什麼數值——這是**慢**的那一步
   （網路 I/O ＋ 解析）
2. **比對層**：`_build_is_table()`／`_build_bs_table()`／`_build_cf_table()`／
   `_build_segment_tables()` 拿解析層的結果，套用 `IS_TEMPLATE`／`BS_TEMPLATE`／
   `CF_TEMPLATE` 的科目比對規則（含 H4/H6 那些 label hint、synonym），組出最終
   的 `StatementTable`——這是**快**的那一步（純本機運算，沒有網路）

**快取卡在解析層跟比對層之間**：快取的是「解析層」的輸出（每份 filing 有
哪些科目、對應什麼值），比對層永遠在快取資料**之上**即時重跑，從不快取。
好處：以後修 Capex 那種科目比對的 hint regex、加新的比率、調 Q4 合成邏輯，
**完全不用管快取**——快取內容沒變，只是重新比對一次，本機運算幾乎瞬間完成。

⚠ **這個論證有一個前提要顯式守住**：解析層的輸出是 **edgartools 這個第三方
套件**解出來的，不是我們自己的邏輯。`fetch_gaap_statements()` 自己的比對層
改版不影響快取沒錯，但 **edgartools 升版**（新的 standardization mapping、
XBRL parser 修 bug）可能讓同一份 filing 解出不一樣的結果——這條快取失效的
軸線跟「我們的比對規則」是兩件事，不能靠同一套 `schema_version` 保護，見
下方「儲存位置與檔案結構」的 `edgartools_version` 欄位。

掛勾點是 `_filing_obj()` 這一個函式：現有四個 builder 都已經共用它，改這裡
全部受益，呼叫端（IS/BS/CF/segment）不用個別修改。**解析層實際輸出的形狀
已經查證過**（見下方存檔格式）：四個 builder 對 `filing.obj()` 的用法只有
一種——`_financials_of(obj)` 取出 `financials`，再呼叫
`income_statement()`／`balance_sheet()`／`cashflow_statement()`，最後
`.to_dataframe()`，全部無參數。也就是說**要快取的不是 `filing.obj()` 這個
大型自訂物件本身**（會踩到 XBRL 物件圖的循環參照），而是這三次
`.to_dataframe()` 呼叫拿到的**三張 pandas DataFrame**。

```
現況：filing → _filing_obj() → (解析結果，只在這次執行的記憶體內有效)

改後：filing → _filing_obj() → 先查本機快取有沒有這份 accession
                                  ├─ 有 → 直接讀檔，不打 SEC
                                  └─ 沒有 → 打 SEC 解析 → 寫進本機快取 → 回傳
```

### ⚠ 快取命中時 `_filing_obj()` 要回傳什麼——替身物件（proxy）

這是動工前一定要先定義清楚的介面，不然實作第一天就會卡住：`_filing_obj()`
目前的回傳值是 edgartools 的 filing 物件，快取命中時我們手上只有存檔的
DataFrame，**沒辦法憑空生出一個真正的 edgartools 物件**，必須回傳一個
**長得像它、但只實作有人真的在用的那幾條路徑**的替身。

三個約束（都是查現有呼叫端查出來的，不是預先假想）：

1. **替身要有 `.financials` 屬性**，不能只支援 `_financials_of(tenq)`
   ——有兩處直接寫 `tenq.financials.xxx()` 繞過那層 helper
   （`fetcher_gaap.py:1123`、`fetcher_gaap.py:2584-2586`），這兩條路也要能吃
   替身，不然快取命中時會直接 `AttributeError`。
2. **替身遇到沒實作的屬性要照 Python 預設行為讓它拋 `AttributeError`，
   絕對不能吞下來回 `None`**。`_financials_of()` 本身是
   `getattr(tenq, "financials", None)`——如果替身對任何未知屬性都回
   `None`（例如用 `__getattr__` 兜底），以後有人在某個 builder 裡新用到
   filing 物件的其他屬性，快取命中的路徑會**安靜地把整份 filing 當成
   沒資料**，清快取重跑卻是好的——這種 bug 極難查。替身只該明確定義
   `.financials` 這一條鏈用到的方法，其餘什麼都不寫，讓存取直接照物件
   沒有這個屬性的正常方式失敗。
3. **三張表（IS/BS/CF）各自可能是 `None`**（`is_stmt is None` 這種判斷在
   `fetcher_gaap.py:1330`／`1499` 都有），快取格式要能分開表示「這張表
   不存在」跟「這張表存在但是空 DataFrame」兩種不同狀態，讀回來要準確
   還原成原本那一種——`_current_q_col(df)` 等下游函式對這兩種狀態的行為
   不一樣。

介面大致長這樣（實作時的確切類別設計留給實作計畫階段，這裡定調行為契約）：

```python
class _CachedFinancials:
    def income_statement(self):   # 回傳 _CachedStatement 或 None
        ...
    def balance_sheet(self):      # 同上
        ...
    def cashflow_statement(self): # 同上
        ...

class _CachedStatement:
    def to_dataframe(self):       # 回傳還原 dtype 之後的 DataFrame
        ...

class _CachedFiling:
    def __init__(self, cached_data):
        self.financials = _CachedFinancials(cached_data)
    # 不定義 __getattr__——任何其他屬性存取一律照 Python 預設拋 AttributeError
```

順帶一提：這也是為什麼現有 G9「單次執行內的記憶體解析快取」
（`_parse_cache_scope()`）完全不用改——它存的就是 `_filing_obj()` 的回傳值，
存真物件或存這個替身，對它來說沒有差別。

---

## 二、儲存位置與檔案結構

沿用 `config.py`／`override_engine.py` 已經在用的 `%APPDATA%` 慣例：

```
%APPDATA%\SEC Financial Tools\filing_cache\
├── NVDA\
│   ├── _manifest.json          ← 這家公司的快取索引
│   ├── 0001045810-25-000123.json   ← 一份 filing 一個檔，檔名是 accession number
│   ├── 0001045810-24-000456.json
│   └── ...
├── AMD\
│   ├── _manifest.json
│   └── ...
└── INTC\
    └── ...
```

- 一個 ticker 一個資料夾，資料夾名稱就是 ticker（大寫）——不用另外查表，
  肉眼在檔案總管就看得出「哪些公司有快取、大概多大」
- 一份 filing 一個 JSON 檔，檔名是 accession number（SEC 給的全域唯一 ID，
  格式固定 `\d{10}-\d{2}-\d{6}`，安全能當檔名）
- **不做全域索引檔**。快取了哪些公司，直接掃 `filing_cache/` 底下有哪些
  子資料夾就知道——公司數量最多幾十家，掃資料夾成本可以忽略，不需要另外
  維護一份容易跟實際檔案脫鉤的全域清單
- **`<accession>.json` 是否存在，才是「這份 filing 有沒有快取」的事實來源**，
  `_manifest.json` 只是**衍生出來、給 GUI 顯示用的索引**，不是查快取要問的
  第一個地方——掃 100 個檔名是微秒級成本，直接查檔案存不存在比維護一份
  「manifest 跟磁碟是否同步」的心智負擔更低（原本設計想省掉掃資料夾，但
  換來一整節錯誤處理在處理兩者對不上的情況，不划算）。manifest 壞掉、遺失
  或跟磁碟對不上，都直接從資料夾內容重建，不影響查快取本身

### `_manifest.json`（每家公司一份，衍生索引——壞了直接重建，不是事實來源）

```json
{
  "schema_version": 1,
  "ticker": "NVDA",
  "cik": 1045810,
  "last_checked_at": "2026-09-03T14:22:10+08:00",
  "filings": [
    {"accession_no": "0001045810-25-000123", "form": "10-Q",
     "filing_date": "2025-08-27", "cached_at": "2026-09-01T09:00:00+08:00",
     "edgartools_version": "5.29.0", "has_financials": true, "size_bytes": 61234},
    {"accession_no": "0001045810-25-000045", "form": "10-K",
     "filing_date": "2025-02-26", "cached_at": "2026-09-01T09:00:12+08:00",
     "edgartools_version": "5.29.0", "has_financials": true, "size_bytes": 88410}
  ]
}
```

- `cik`：company 的 CIK（跟 SEC 打交道真正的鍵，ticker 只是顯示用的別名）。
  ticker 會換手（例如公司更名、代號被回收給別家），載入快取時比對
  `company.cik` 跟這裡記的是否一致，**不一致就整包視同無快取重建**——這種
  錯不會報例外，只會安靜地把別家公司的數字餵給使用者，比其他任何失效情境
  都危險，資料夾本身仍用 ticker 命名（肉眼可讀這點是對的，不用改）
- `last_checked_at`：上次去 SEC 查「filing 清單」的時間（純粹給 GUI 顯示，
  不是快取有效期限，不用來判斷要不要重查——每次都查，見下方更新機制）
- `filings`：目前本機已經有的 filing 清單，含 `edgartools_version`（見下方
  「更新機制」的版本檢查）與 `has_financials`（見「負向快取」）
- `schema_version`：manifest 自己的格式版號，對不上就當整包不存在、從磁碟
  重建（不影響已經存在的 `<accession>.json`，那些檔案本身另有各自的
  `schema_version`）

### 單一 filing 快取檔（`<accession>.json`）

存三張 pandas DataFrame（income statement／balance sheet／cashflow
statement，即 `_financials_of()` 之後三個 `.to_dataframe()` 呼叫的結果）。
每張表可能是 `None`（見上方替身物件的約束 3），存的時候要能分辨「這張表
不存在」（值為 `null`）跟「存在但是空表」（有 `columns`/`index` 但
`data` 是空陣列）：

```json
{
  "schema_version": 1,
  "accession_no": "0001045810-25-000123",
  "form": "10-Q",
  "filing_date": "2025-08-27",
  "cached_at": "2026-09-01T09:00:00+08:00",
  "cik": 1045810,
  "edgartools_version": "5.29.0",
  "has_financials": true,
  "dataframes": {
    "income_statement": {
      "data": { "...": "json.loads(df.to_json(orient='split')) 的物件內容" },
      "dtypes": {"concept": "object", "level": "int64", "abstract": "bool", "...": "..."}
    },
    "balance_sheet": { "data": { "...": "..." }, "dtypes": { "...": "..." } },
    "cashflow_statement": null
  }
}
```

三個實作細節（都是第二輪 spike 才發現、原本這節有漏掉）：

1. **存 `json.loads(df.to_json(orient="split"))` 的物件，不要存
   `df.to_json(...)` 的字串。** 直接把字串塞進外層 JSON 等於整份內容被
   逃逸一次（每個 `"` 變 `\"`），檔案膨脹約 10~15%，而且用文字編輯器打開
   完全不能看，debug 不方便。讀取時反過來 `json.dumps(...)` 餵回
   `pandas.read_json()`，或直接用 `columns`/`index`/`data` 自己組
   `pd.DataFrame`。
2. **`dtypes` 是必要欄位，不是可省略的**。`to_json(orient="split")` 本身
   不含 dtype 資訊，`pandas.read_json()` 對「整欄都是 null」的欄位會推成
   `float64`，跟原本可能是 `str`／`object` 的型別對不上——讀回來之後要
   照存檔時記下的 `dtypes` map 明確 `astype()` 回去，不能依賴自動推斷。
3. **`edgartools_version` 的取得方式要指名**：用
   `importlib.metadata.version("edgartools")`（實測回 `"5.29.0"`）——
   **不要用 `edgar.__version__`，這個屬性不存在**（實測會
   `AttributeError`）。取不到版本號（例如未來套件改了發布方式）就直接
   視同無快取，不要寫一個 `"unknown"` 之類的預設值混進檔案裡，那會讓版本
   比對邏輯永遠比對成功或永遠失敗，看預設值怎麼寫，兩種都是錯的。
4. **`<accession>.json` 的寫入也要走 tmp + `os.replace()` 的原子寫法**
   ——理由跟 manifest 一樣：兩個實例（批次抓取＋跨公司比較）有機會同時
   解析到同一份 filing、同時要寫同一個檔名，非原子寫法會交錯出半截
   JSON，磁碟寫到一半空間不夠也一樣。tmp 檔名要帶 PID
   （例如 `<accession>.json.<pid>.tmp`），避免兩個實例互相蓋到對方的
   暫存檔。成本跟 manifest 那套一樣低，兩種檔案統一用同一套寫入 helper。

---

## 三、更新機制：怎麼知道有沒有新 filing

CTH 選的是「每次都自動查」。流程：

1. `fetch_gaap_statements()` 一開始，對每個要抓的表單類型（10-Q／10-K）
   呼叫現有的 `_list_filings()`——這一步本來就存在、本來就會做，**不是新增
   的網路成本**，只是原本抓完清單後每份都要再打一次 `filing.obj()`，現在
   多一個「查本機有沒有」的判斷
2. 拿到的 filing 清單（含 accession number）跟本機 `<accession>.json` 是否
   存在比對（不是查 manifest，見上一節）：
   - accession 已經在本機 → 讀本機 JSON。讀之前先比對 `cik` 與
     `edgartools_version` 是否跟目前環境一致，**任一個不一致就視同無快取**，
     照舊打 SEC 重解析、覆蓋掉舊檔——正確性優先於速度，寧可那次變慢也不要
     餵錯資料或吃到舊版 parser 的 bug
   - accession 不在本機（新 filing，或第一次抓這家公司） → 照舊打 SEC 解析。
     解析成功就**立刻寫入**這份 filing 的快取檔（見下方「寫入時機」），
     `financials` 是 `None`（pre-XBRL 舊申報，2009 年前常見）就寫一筆
     `has_financials: false` 的快取檔，記著「這份試過了、沒有財務資料」，
     下次不用再打一次 SEC 重試（負向快取；**跟網路失敗不同**——網路失敗
     不算進負向快取，那是暫時性的，繼續交給既有的 D11-B 缺漏帳本機制，
     每次都應該重試）
3. 全部處理完，更新 `_manifest.json` 的 `last_checked_at`（連帶重建整份
   `filings` 索引，反映這次跑完後磁碟上實際的內容）

**多出來的網路成本只有「查清單」這一步，本來就存在，不是新增的**——新增的
是省下「每一份都重新打 SEC 解析」這一大段。第一次抓某家公司完全沒有加速
（因為本機還沒有任何東西），第二次以後才吃到紅利，而且只需要抓「上次抓完
到現在新增的那幾份」。

**寫入時機：逐份即時落檔，不要累積到整趟抓完才一起寫。** 一趟抓取可能要
好幾分鐘，中途網路斷掉、或使用者關視窗，若快取是「整趟跑完才一次寫入」，
這中間已經抓到的進度會全部白費——這跟本專案 log 規則「當下就寫，不可累積
到結束才吐」是同一個道理，抓取失敗時才最需要保住已經抓到的部分。失敗的
那份 filing 不寫任何快取（不管是正向還是負向），下次照樣會重試。

**⚠ 修正案（10-Q/A、10-K/A）現況本來就不會被抓到，這不是這個快取功能造成
的，也不是這個功能要解決的。** `_list_filings()` 目前呼叫時
`amendments=False`（`fetcher_gaap.py:304`），所以「公司重編財報會開一份
新的 filing/accession」這句話雖然對，但那份修正案（amendment）本身**現在
就不在抓取清單裡**，不管有沒有快取都一樣抓不到。快取只保證「清單裡查得到
的 filing」會被正確、即時地補齊，不擴大也不縮小現有的抓取範圍。這件事跟
CTH 講清楚（後面「風險」那節會再提一次），要不要處理修正案是另一個獨立
議題，不在這次範圍內。

**不會有資料過舊的風險（在上面這個澄清的前提下）**：filing 一旦存在 SEC 上
內容不會變——本機快取的每一份 filing 永遠正確，差別只在「本機還沒有最新
那幾份」，而這一步每次都會自動補齊。

**log 要記快取命中數，不能只看耗時。** 耗時變快也可能只是那天 SEC 比較順、
跟有沒有吃到快取無關；沒有命中數字，使用者跟維護者都無法判斷「這次到底
有沒有吃到快取」。沿用現有「設定塞在起始 `===` 那行」的規則，加一個欄位：

```
=== 2026-09-03 14:22:10 Fetch NVDA | GAAP | 10-Q/10-K | max80 | cache 24/25 ===
```

`cache 24/25` 表示這趟要處理的 25 份 filing 裡有 24 份是直接讀快取的。這行
本身固定英文（2026-09-02 起 `logs/app.log` 一律英文的既有規則）。

---

## 四、GUI

CTH 要求「GUI 要有一塊地方看得到」。放在 **Tab3（進階設定）**——那裡本來就是
SEC identity、AI 設定、抓取上限、模板模式這類「維護者會去看」的區塊，快取
管理性質相同，不需要另開分頁。

新增一個小區塊「本地資料快取」：

```
本地資料快取                                    總容量：39.2 MB  [開啟資料夾]
┌─────────────────────────────────────────┐ ← 固定高度，約 4~5 列可視，
│ NVDA    102 份 filing    18.4 MB   [清除] │   超過就在這塊內部捲動
│ AMD      76 份 filing    12.1 MB   [清除] │
│ INTC     54 份 filing     8.7 MB   [清除] │
└─────────────────────────────────────────┘
                              [全部清除]
```

- **這塊本身要再包一層獨立的固定高度捲動容器**（沿用 `main.py` 已有的
  `_build_fixed_height_scrollable()`，高度落在 100~120px、約 4~5 列）。
  ⚠ Tab3 整頁本來就是靠 `_build_fixed_height_scrollable(tab, height=self._TAB3_HEIGHT)`
  撐住固定高度（`main.py:2017`，`_TAB3_HEIGHT = 355` 是實測貼齊值），快取
  常駐個位數到二三十家公司，清單直接攤開會把 Tab3 撐爆、擠掉上面 SEC
  identity／AI 設定的可視範圍——這塊必須自帶捲動，不是整頁一起長高
- **加完這個區塊要重新量測 `_TAB3_HEIGHT`**，確保 Notebook 整體高度與 log
  行數不變（`docs/ARCHITECTURE.md`「視窗擺放」那節的既有規則：改任何一頁
  版面都要重量），這條要寫進下方「驗收」
- 總容量那一行 + 「開啟資料夾」按鈕：CTH 選的是手動清、不做自動上限，那
  這個決定要有依據——「現在到底佔多少」比逐家分項更接近他真正想看的東西，
  「開啟資料夾」讓他可以自己用檔案總管進一步查看/處理，不用我們另外做更
  細的管理 UI
- 列表來源：掃 `filing_cache/` 底下的資料夾，每個資料夾讀 `_manifest.json`
  拿 filing 數，資料夾實際大小用 `os.path.getsize()` 加總（`except OSError`
  要接住——正在被清除或另一個實例正在寫入時檔案可能瞬間消失）
- 列表刷新時機：切到 Tab3 時、任一次「清除」之後、任一次抓取（Tab1／批次／
  跨公司比較）完成之後——不用即時輪詢
- 「清除」：整個刪掉那個 ticker 的資料夾。下次抓這家公司會當作全新開始
- **抓取進行中（Tab1／批次／跨公司比較任一個 worker thread 還在跑）時，
  兩顆清除按鈕都要 disable**——沿用專案既有「執行中鎖住相關按鈕」的慣例，
  不然會邊寫邊刪同一個 ticker 的資料夾
- **「全部清除」是唯一不可逆的破壞性操作，要有二次確認對話框**（雖然只是
  快取，重抓 20 年份是好幾分鐘的代價，值得防手滑）
- 沒有「立即更新」按鈕——因為每次抓取本來就會自動查新（見上一節），不需要
  另外提供手動觸發
- **所有新增字串都要進四個 locale 的 `gui.*` key**（`tests/test_i18n.py`
  第 3 條測試會擋掉 `src/` 裡任何寫死的中日文字面，這條是永久防線，不是
  這次才有的規則）；快取相關的 log 訊息比照 2026-09-02 起的既有規則，一律
  英文（`docs/ARCHITECTURE.md`「logs/app.log 的語言與格式」）

---

## 五、錯誤處理

- **快取檔案損毀（JSON 壞掉、schema_version／cik／edgartools_version 任一
  對不上）**：當作沒有這份快取，照舊打 SEC 重抓，抓到後覆蓋掉壞掉的那個
  檔——不讓損毀的快取拖垮整趟抓取，這跟現有專案「抓不到就跳過繼續、不中斷
  整體流程」的一貫原則一致
- **寫入快取失敗（磁碟滿、權限問題）**：記一筆 log、繼續執行——快取只是加速，
  寫不進去不影響這次抓取結果，跟現有的 `err.output_locked_excel` 那類
  「輸出失敗不中斷主流程」是同一個精神，但這裡連提示都不用跳出來吵使用者，
  純寫 log
- **manifest 遺失、損毀、或跟磁碟實際內容對不上**：manifest 是衍生索引不是
  事實來源（見「儲存位置與檔案結構」），直接**從資料夾實際內容重建**——
  掃有哪些 `<accession>.json`，逐份讀出 metadata 組回 `filings` 清單，不需要
  特別的「修正」邏輯，跟「manifest 本來就不存在」是同一條路徑
- **manifest 寫入要是原子操作**：先寫進暫存檔（例如 `_manifest.json.tmp`），
  完成後用 `os.replace()` 換成正式檔名——因為批次抓取跟跨公司比較有機會
  同時抓到同一家公司（專案全域規則本來就假設兩個實例可能同時跑），半寫
  一半被另一個行程讀到會壞。`<accession>.json` 本身不會有這個問題（不同
  filing 的檔名不會撞、寫入前用「檔案是否已存在」判斷要不要重寫，本來就
  是覆蓋語意單純的操作）

---

## 六、測試要釘什麼

- 磁碟已有 `<accession>.json` 且 `cik`／`edgartools_version` 都相符 → 不呼叫
  `filing.obj()`（mock 驗證呼叫次數為 0），三張 DataFrame 讀回來的內容與
  dtype 要跟原始 DataFrame 一致（含「整欄皆 null」那個 dtype 還原的細節）
- 新 filing（磁碟沒有對應檔案）才會真的呼叫 `filing.obj()`，成功後磁碟要
  多一個檔案、manifest 要多一筆
- `cik` 不符 / `edgartools_version` 不符 / schema_version 不符 / JSON 本身
  壞掉 → 四種情況都視同沒快取，正常重抓，不拋例外
- manifest 遺失或跟磁碟對不上 → 從磁碟內容正確重建，不需要人工介入
- **負向快取**：`financials` 解出來是 `None`（pre-XBRL）的 filing，第二次
  查詢不再呼叫 `filing.obj()`，但**網路失敗的 filing 絕對不能被當成負向
  快取**——這條要故意寫一個「filing.obj() 拋網路例外」的測試，確認下次
  還是會重試，沒有被誤記成「沒有 financials」
- **查詢邊界不能造成漏抓**：先用 `start_year=2020` 抓一次（只碰到 2020 年
  以後的 filing），再用完整期間（不限年）抓同一家公司，第二次一定要真的
  去補抓 2020 年以前那些從沒進過快取的 filing，不能因為「這家公司已經有
  快取」就少抓。同理要測 `max_filings` 調大、以及 quarterly/annual/Both
  三種頻率切換時都不會漏抓已存在快取之外的部分
- 寫入是逐份即時落檔：模擬「抓到一半丟例外」，已經成功解析的那幾份要留在
  磁碟上，不能因為整趟沒跑完就全部消失
- manifest 寫入要驗證是 tmp + `os.replace()` 的原子寫法（不會留下半寫的
  `_manifest.json`）
- 三個呼叫端（Tab1 單一抓取、批次抓取、跨公司比較）**行為不變**——這是加速層，
  不改變任何回傳資料的內容，用既有的 `scripts/excel_golden.py`（make base /
  make new / check base new）確認輸出沒有變化，這是本專案既有的驗證工具，
  G9 記憶體快取當初就是靠它驗過「5,678 格 0 格不同」
- GUI：Tab3 快取列表正確反映 `filing_cache/` 的實際內容、「清除」真的把資料夾
  刪乾淨、「清除」之後那家公司要重抓、抓取進行中兩顆清除鈕要 disable、
  「全部清除」要跳確認對話框、四個 locale 的新字串都要有（`test_i18n.py`
  既有測試會自動擋）

## 驗收

- **量化目標，不是「感覺比較快」**：參考 `docs/ARCHITECTURE.md` 記錄的 ARLO
  實測基準（25 份 filing 共 66 秒，其中 XBRL 解析 19.9s ＋ `to_dataframe()`
  28.4s ≈ 48s 是這次會被快取吃掉的部分，剩下約 18s 是查清單／其他網路
  往返，快取**不會**消除）。用同一家公司實測「清快取重抓」vs「全熱快取」，
  預期落在 **3~5 倍**（例如 ARLO 全熱 <15s），不是「秒開」——這件事要先
  跟 CTH 對齊預期，免得驗收時覺得「怎麼還要等」
- 手動在 SEC 上確認某家公司「最新一期 10-Q」已經存在快取後，公司真的發了
  新一期財報時，下次抓取要抓到新的那一期（不是繼續讀到舊資料）
- **用 `scripts/excel_golden.py` 做正式的數字比對**：同一份輸入，「清快取
  重抓」產出的 Excel 存一份 base，「讀快取」產出的存一份 new，跑
  `check base new`，逐格數值／格式要 **0 格不同**——快取不能改變任何輸出
  內容，只能改變耗時
- Tab3 加完「本地資料快取」區塊後，重新量測 `_TAB3_HEIGHT`，Notebook 整體
  高度與其他分頁的可視 log 行數要維持不變

## 風險

- **低（已實測，原本評估中～高）：DataFrame 序列化。** 對一份 ARLO 最新
  10-Q 做過 spike：income statement 52 列 × 20 欄，dtype 只有
  str／float64／int64／bool 四種，`to_json(orient="split")` 三張表合計約
  41KB，`read_json` 讀回來 shape 與 dtype 一致（除了前面提過的「整欄皆
  null → 被推成 float64」那個要顯式 `astype()` 處理的細節）。100 份 filing
  估計 5~10MB，符合 GUI 範例「102 份 = 18.4MB」的量級
- **中：edgartools 版本漂移**（原本這份設計沒有涵蓋的軸線）。快取的是
  第三方套件（edgartools，`docs/ARCHITECTURE.md` 記載目前是 5.29.0、
  個人維護、對照表薄）解析出來的結果，套件升版可能讓同一份 filing 解出
  不同內容——已經用 `edgartools_version` 欄位處理（見「儲存位置」／
  「更新機制」），版本不符就當沒快取重建，正確性優先於速度
- **低：新舊快取共存**。開發過程中 schema 若真的要改版，`schema_version`
  機制已經涵蓋——舊版快取被忽略、重新建立，不會讓程式壞掉，只是那次會變慢
  （等於沒快取）
- **低：跟現有 G9「單次執行內的記憶體解析快取」`_parse_cache_scope()` 的關係**——
  兩層快取不衝突，記憶體那層可以維持不動（同一次執行內，讀本機磁碟快取的
  結果一樣會進記憶體快取，避免同一次執行內重複讀檔）
- **低：兩個實例同時抓同一家公司**（批次抓取＋跨公司比較同時跑，或兩個
  GUI 視窗）。`<accession>.json` 的寫入是「檔案不存在才寫」，兩邊撞寫同一份
  結果應該相同，不算真的衝突；manifest 用原子寫法（tmp + `os.replace()`）
  避免半寫的檔案被讀到

## 效益範圍：這個快取涵蓋什麼、不涵蓋什麼

這件事要跟 CTH 講清楚，避免驗收時預期落空。快取只加速 `_filing_obj()` 這
一條線（IS/BS/CF 三張表的逐份 filing 解析），**不涵蓋**：

- Non-GAAP（8-K 新聞稿）那條路——`fetcher_nongaap.py` 已經有自己獨立的
  `nongaap_cache.json`，跟這次無關，不受影響也不會被取代
- `company.get_facts()`（流通股數 `_fetch_shares_outstanding`）與
  `_list_filings()` 的清單查詢——這兩個每次都照樣打網路，不受快取影響

## 未來可能，這次不做

- 改用 SEC 的「公司整批 XBRL 資料」單一端點取代逐份 filing 抓取（設計裡的
  方案 C，範圍大風險高，另外評估）
- 自動容量上限／到期清理（CTH 明確說手動清就好）
- GUI 加「立即更新」手動觸發按鈕（目前判斷不需要，因為每次抓取都會自動查新）
- 抓取修正案（10-Q/A、10-K/A）——現況 `_list_filings(amendments=False)`
  本來就不抓，這次不改；要不要處理是另一個獨立議題，可以另開 TODO
