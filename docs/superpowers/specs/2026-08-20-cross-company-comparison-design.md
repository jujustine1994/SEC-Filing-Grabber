# 跨公司財務比較功能 設計

> 2026-08-20 · 對應 `docs/TODO.md` F1
> 狀態：設計已與 CTH 逐段確認，待實作（D0-1 Q4 推算已驗證通過，卡點已解除）

## 目標

讓使用者一次選擇多家公司、多個財務指標與比較期間，輸出一份**獨立的新格式 Excel**——
含原始資料表、可互動的時間點快照、以及每個指標各自的歷史趨勢圖表。跟現有「一次抓一家
公司」的 Tab1/Tab2 抓取流程完全分開、互不影響、不共用輸出資料夾。

## 不做的事

- **不改動現有單一公司抓取流程**（Tab1/Tab2、`fetch_gaap_statements()` 本身、既有 Excel 格式）。
  新功能是外掛，只是呼叫既有函式取資料。
- **不做估值倍數**（P/E、EV/EBITDA、P/B）。需要股價/市值資料，工具目前完全不抓市場數據，
  這是比本功能更大的擴充，另記 TODO F2。
- **不做自己的 Q4 推算邏輯**。Q4 完全依賴 `fetcher_gaap.py::_synthesize_q4()`（D0-1 已完成
  並驗證），本功能只是呼叫 `fetch_gaap_statements()` 拿現成結果。
- **`Snapshot` sheet 的公式不保證能被其他程式讀到**——這是刻意的取捨（見下），需要給別的
  程式/skill 讀的資料另外存在 `Snapshot_Manual`。

---

## 一、架構

新增兩個獨立模組，不動現有單一公司抓取流程：

- **`src/comparison.py`**：對每個使用者選的 ticker 呼叫既有 `fetch_gaap_statements()`
  （季／年報表視使用者選擇的頻率），再用 `ratios.py::build_ratio_table()` 算出比率表。
  把使用者選的「原始科目」與「比率」從這些表裡抽出來，重組成
  `{指標: {公司: {期間標籤: 值}}}` 加上對應的「期末結算日」對照，供 writer 使用。
- **`src/comparison_writer.py`**：吃上面的資料結構，用 `openpyxl.chart`
  （目前專案完全沒用過，這次新增依賴）寫出新格式 Excel（結構見第三節）。

GUI 新增 **Tab4「跨公司比較」**，跟 Tab1/Tab2/Tab3 同級。

---

## 二、GUI

### Tab4 主畫面

顯示目前已選的公司／期間／指標／輸出資料夾摘要，一顆按鈕開啟「選擇比較內容」
`Toplevel` 視窗，下方是輸出資料夾設定、執行按鈕、進度條與 log（比照 Tab1 既有元件）。

### 選擇比較內容視窗（分兩段）

**① 選公司**

- Ticker 輸入框，打字時即時比對 `company_cache.json`（Tab1／Watchlist 已在用的同一份
  ticker→公司名快取）跳出自動完成建議清單
- 支援一次貼上 `nvda, amd, dell, avgo` 逗號分隔清單，逐一比對快取；查不到的 ticker
  標紅提示，不阻擋其餘公司送出
- 選中的公司以 `NVDA NVIDIA CORP ✕` 這種 chip 顯示在下方，可個別移除

**② 選指標**

- 起始年／結束年下拉 ＋ 頻率下拉（季度／年度，兩者都支援）
- 指標分類用**下拉選單**切換（損益表／資產負債表／現金流／比率，比率再細分成長性／
  獲利能力／槓桿償債／現金流／營運效率／結構規模／報酬率等子分類），下方勾選框只顯示
  當前分類的項目，勾選結果累加成「已選指標」chip 列表——**切換分類不會清掉之前選的**
- 快照時間點輸入框（如 `2025/12/31` 或 `2024Q4`），供輸出檔的 `Snapshot` sheet 使用

「取消」／「確定」關閉視窗，「確定」把選擇寫回 Tab4 主畫面的摘要顯示。

---

## 三、輸出 Excel 結構

檔名 `比較_A_B_C_YYYYMMDD.xlsx`，另存新輸出資料夾（不覆蓋 Tab1/2 輸出）。

以 5 家公司、8 個指標為例，共產出 `2 + N` 張 sheet（N＝選定指標數）：

### `Compare_Data`（唯一一張原始資料表）

每個選定指標各一個區塊，由上往下疊：

```
■ Revenue
  公司    2023Q1   2023Q2   2023Q3  ...
  期末結算日 2023/03/31 2023/06/30 2023/09/30 ...   ← 靜態文字，非公式
  NVDA    ...
  AMD     ...
  ...

■ Gross Margin(%)
  （同上格式）
```

「期末結算日」列是**靜態文字**（不是公式），沿用 D0-5 已知限制記錄過的 pattern——期間
標籤欄位若是公式，openpyxl 讀不到 Excel 沒開過的快取值；用靜態日期列當 `MATCH` 的 key
才穩。這一列同時是 `Snapshot` 公式查找的依據。

### `Snapshot`（活的，公式驅動）

頂端一格黃底輸入格（如 `2025/12/31`），下方列＝公司、欄＝所有選定指標。每一格用
`INDEX`/`MATCH` 公式對 `Compare_Data` 對應區塊的「期末結算日」列取值——**改輸入格
Excel 就自動重算**，不用重新產檔。

這份**只給使用者在 Excel 裡看**，用真公式，跟 `ratios.py`（`Data_Ratios`）刻意「不寫
公式、只寫算好的值」的慣例不同——因為 `Data_Ratios` 有下游 skill 用 openpyxl 直接讀，
公式讀不到快取值會出問題；`Snapshot` 目前沒有這個需求。

### `Snapshot_Manual`（空白，供人工凍結存檔）

跟 `Snapshot` 同樣的欄位結構（列＝公司、欄＝指標），**產出時是空白的**。使用者想保留某
個時間點的快照供其他程式/skill 讀取時，自己把 `Snapshot` 算出來的值複製貼上（貼值，不
貼公式）進這張表。這是目前唯一保證「開檔前用 openpyxl 讀得到值」的 sheet。

### `Chart_<指標>` × N（每個選定指標各一張，只放圖表）

資料來源是 `Compare_Data` 對應區塊。**一張圖，預設折線圖**（一條線一家公司，橫跨選定
期間）。使用者要看長條圖版本，在 Excel 裡對圖表右鍵「變更圖表類型」自己切——工具端
不用為同一指標同時產兩種圖表物件。

---

## 四、比率目錄擴充（`ratios.py`）

`RATIO_DEFS` 目前是扁平 list，只用註解分區塊。這次改成**每筆比率帶明確 category 欄位**，
讓選擇視窗的分類下拉照 category 分組，以後加新比率只要加一行帶對的 category，GUI 自動
長出新選項，不用改介面程式碼。

在現有 28 個比率之上，確認新增（前兩個是 CTH 原始需求缺的，其餘是完善建議並經 CTH
確認納入）：

| 分類 | 新增比率 |
|---|---|
| 成長性 | Gross Profit QoQ(%)、Operating Income QoQ(%)、EPS QoQ(%)、FCF YoY(%)、EBITDA YoY(%) |
| 槓桿償債 | **Debt Ratio(%)**＝Total Liabilities/Total Assets（CTH 要的「負債比」）、Debt-to-Equity(x)、Equity Multiplier(x)＝Total Assets/Equity、LT Debt to Capital(%) |
| 現金流 | Operating CF Margin(%)、OCF/Net Income(x) |
| 營運效率 | Asset Turnover(x)、Inventory Turnover(x)、Receivables Turnover(x) |
| 結構／規模 | **D&A/Revenue(%)**（CTH 要的「折舊佔營收比率」）、EBITDA($)、Total Debt($)、Net Debt($)、Working Capital($)、Cash Ratio(x)、COGS Ratio(%) |
| 報酬率（近似值） | ROIC(%) ≈ Operating Income×(1−有效稅率)/(Total Debt+Equity−Cash)——業界慣用簡化版，非嚴謹版（未拆一次性項目），欄位名稱要標註是近似值 |

---

## 五、錯誤處理

比照現有 `collect_gaps()` 原則：單一公司抓失敗（ticker 打錯、外國私人發行人如 D9 的
NBIS 那種抓不到 10-Q/10-K……）不中斷整體比較，跳過該公司、在 `Compare_Data` 與 log
標記「XXX 無法取得資料」，其餘公司照常輸出。

---

## 六、測試計畫

- `comparison.py`：單元測試涵蓋「多公司資料重組」「部分公司抓取失敗時的跳過邏輯」
  「季度／年度切換」「選定指標的抽取（原始科目 + 比率）」
- `comparison_writer.py`：單元測試驗證 `Compare_Data` 區塊排列、期末結算日靜態列、
  `Snapshot` 公式字串正確、`Snapshot_Manual` 表頭結構、圖表物件的資料來源 range 正確
- `ratios.py` 新增比率：比照現有 28 個比率的測試 pattern（見 `tests/test_ratios.py`），
  補齊新增的約 20 個公式的正確性與缺值時回 `None` 的行為
- 端對端：拿 2-3 家真實公司（如 NVDA/AMD/AVGO）跑一次完整流程，人工核對輸出 Excel
  的 `Snapshot` 公式在 Excel 裡改日期後確實正確重算
