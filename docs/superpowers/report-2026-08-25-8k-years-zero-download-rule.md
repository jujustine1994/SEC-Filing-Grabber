# 8-K `--years` 篩選：驗證結果與修法提案

> 驗證日期：2026-08-25　驗證者：另一個 session（純驗證，**未改任何專案檔案**）
> 資料：SEC EDGAR 實抓，26 家 × 最近 8 份 Item 2.02 8-K = 200 份，零 AI 呼叫
> 產出：`scratchpad/8k_period_audit.json`、`8k_period_audit_oos.json`、`final_check.json`

## 1. 原提案的診斷要修正兩點

原提案：「`fy_end_month` 查得太晚 → 把 `_fy_end_month()` 提前傳進
`_list_earnings_filings()`，內部改用 `fiscal_input.fiscal_quarter_of(period_end, ...)`」。

**現象描述正確**（`--years` 確實用發布日算的 label 在篩，邊界會錯），但兩點要修正：

1. **根因不是查詢時機。** 就算 `fy_end_month` 提前傳進去，`_list_earnings_filings()`
   裡**沒有 `period_end` 可用**。真實期末日是 `cli.py::_period_end_from_tables()`
   從新聞稿表格內文 regex 抓的（`cli.py:283`），前提是 `filing.obj()` 已把文件
   下載並解析完。提案第 3 點說「沿用同一個取值路徑即可」——那條路徑就是「下載整份 8-K」。
   `_list_earnings_filings()` 的 docstring 明寫 *"Filters entirely on listing
   metadata … no document is downloaded here"*，照原提案改等於把 20 年、上百份
   8-K 全下載完才知道哪幾份落在 `--years` 區間，而 `--years` 存在的目的正是避免下載。

2. **label 不只用於年份篩選。** 同一個 label 還是 `_dedupe_by_label()` 的 key 與
   `_find_missing_quarters()` 缺季偵測的依據。改 label 會連動這兩處（實測見 §4）。

## 2. 驗證了什麼

目標：找一條**零下載**、只吃 listing metadata + 一次 `Company.fiscal_year_end`
查詢就能算出正確財季的規則。試了三版：

| 規則 | 命中率（119 份可比對） | 敗因 |
|---|---|---|
| A. `fiscal_quarter_of(period_of_report) − 1 季` | 58.0% | `fiscal_quarter_of()` 內建「往前推 15 天」是為期末日設計的；實測發布延遲 **4–58 天**跨度太大，減 15 天後有時落回本季、有時落到下一季 |
| B. 名目季末＝財季結束月的**當月最後一天** | 79.8% | 52/53 週制真實季末落在名目月底前後最多 20 天（WDC 季末 1/2、COST Q3 季末 5/10），整家偏一季 |
| **C. 名目季末＝EDGAR `fiscal_year_end` 的完整 MMDD 往前推 3/6/9 個月** | **100%** | — |

### 規則 C 的定義

```
fye = Company(ticker).fiscal_year_end          # 例 WDC "0703"、COST "0830"
名目季末候選 = fye 日期往前推 0/3/6/9 個月（該月沒有那天就退到當月最後一天）
label = 「不晚於（發布日 + tol）」的最新一個名目季末，
        套 fiscal_input.fiscal_quarter_of(名目季末, fy_start_month(fye 月))
```

`tol` 允許名目季末落在發布日之後一點點：COST 的 Q3 真實在 5/10 結束、名目季末
算出來是 5/30，而它 5/28 就發布了。**tol 在 3～30 天之間結果完全相同**（高原很寬），
建議取 21。tol=0 掉到 95.0%。

### 結果

基準＝production 現在那條路：`fiscal_quarter_of(新聞稿抓到的期末日, fy_start_month(EDGAR FYE))`，
也就是 `cli.py::_fiscal_label()` 現在吐的 `fiscal_label`（`scripts/verify_8k_fiscal_labels.py`
已驗過 15 家 120/120）。刻意不用 audit script 的 `stated`（regex 從內文抓，已知
PANW/TGT/KR 會抓到財測或年度段落）。

| 樣本 | 基準可信份數 | 命中 | 殘差 |
|---|---|---|---|
| In-sample（調規則用的 16 家） | 113 | **113 = 100%** | 全 0 |
| Out-of-sample（沒調過的 10 家：WMT TGT KR DE FDX NKE ADBE CSCO HPQ JNPR） | 44 | **44 = 100%** | 全 0 |

「基準可信」＝新聞稿抽到的期末日通過 `0 < 發布日 − 期末日 ≤ 90 天`。剔除的
43 份全部是 **audit script 的期末日 regex 失敗**，不是規則失敗：TGT 有四份抽到
`2025-02-01`（資產負債表的年度表頭）、FDX 三份抽到 `2024-05-31`，發布日減期末日
超過 400 天。INTC 8 份全部抽不到（它的新聞稿內文從不寫 "ended <日期>"）——
**這正好是規則 C 的賣點：INTC 這種抽不到期末日的公司，零下載規則照樣給得出正確 label。**

26 家的 `Company.fiscal_year_end` 全部拿得到合法 MMDD，沒有一家是空的。

## 3. 已知的一個慣例落差（不是這次引入的）

1/2 月結算的零售股，本專案慣例（`fiscal_input._fiscal_year()`：起始月之後的月份
算下一個財年）會比公司自稱的財年**多一年**：TGT 結束在 2026-05-02 那季，本專案
標 `FY2027Q1`，Target 自己叫 fiscal 2025 Q1。這與 `fetcher_gaap._col_to_quarter_label()`
一致，NVDA/CRM 也是同樣情形，`docs/8k-period-off-by-one.md` 已明白接受
（「NVDA 4 月底那季，兩邊都是 FY2027Q1」）。規則 C 沿用同一套，不製造新的不一致。

## 4. 影響量化（200 份）

- **新舊 label 年份不同 → `--years` 會選錯的：63 份 = 31.5%**（不是只有邊界一兩份）
  - 各家：CRM 6、NVDA 6、WMT 6、TGT 5、KR 5、ORCL 4、QRVO 4、FDX 3、NKE 3，其餘 1–2
- 新舊 label 有任何差異（含只差季）：160 份 = 80.0%
- **dedupe 碰撞組數：舊 6 組 → 新 7 組**（沒有消失，`_dedupe_by_label()` 仍必要）
  - 新標籤下碰撞：QRVO FY2026Q2、WDC FY2025Q2、TGT FY2026Q4、KR FY2025Q4、
    FDX FY2026Q4、NKE FY2026Q4、HPQ FY2026Q1
  - 現行 dedupe 規則（有 Item 9.01 優先、其次取最新）是 2026-08-09 修好的，
    與 label 怎麼算獨立，**不需要跟著改**，但改 label 後要重跑一次確認碰撞的
    那 7 組挑對了。

## 5. 提案：接下來怎麼做

### 5-1　實作（建議走 TDD）

1. `_list_earnings_filings()` 新增參數 `fiscal_year_end: str | None = None`（傳
   `"0703"` 這種 MMDD 原字串，不是只傳月份——月份不夠，見 §2 規則 B 的敗因）。
2. 新增純函式（建議放 `fiscal_input.py`，與 `fiscal_quarter_of()` 同一支模組）：
   `quarter_label_from_announcement(announce_date, fiscal_year_end_mmdd, tol=21) -> str`。
   純日期運算、零 I/O，測試好寫。
3. `_list_earnings_filings()` 內：拿得到 `fiscal_year_end` 就用新規則算 label，
   **拿不到就退回現行 `_period_to_quarter_label()`**（26 家實測都拿得到，但不要
   讓 EDGAR 缺欄位變成整批失敗）。
4. `cli.py::_fy_end_month()` 改成回傳 MMDD 原字串（或另加一個
   `_fiscal_year_end()`），並把查詢**移到 `_earnings_filings()` 之前**——這步
   原提案是對的。多出來的成本是每個 ticker 一次 submissions 請求，本來就要查。
5. `cli.py` 的 `label_source` 由 `"period_of_report"` 改成
   `"announcement+fiscal_year_end"`，`_LABEL_WARNING` 整段可以拿掉或改寫——
   下游 skill 現在讀得到的那段警告會過期。**這是對外行為改變，要一起改
   `docs/CLI.md`。**
6. `fiscal_label`（下載後從真實期末日算的那個）**不動**。它仍然是最準的，
   新規則只是讓「還沒下載時」的 label 跟它一致。

### 5-2　驗收

- 新增單元測試：規則 C 的邊界案例，至少涵蓋 WDC（季末 1/2，跨年）、
  COST（Q3 季末 5/10 但 5/28 才發，需要 tol）、MU（FYE 0903，名目月底法會錯）、
  NVDA/TGT（1/2 月結算的財年編號慣例）。這四家就是三版規則的分水嶺。
- 端對端：`cli.py press-release --years 2023-2023` 對 16 家跑一次，確認季數與
  真實財期一致。
- 回歸：`scripts/verify_8k_fiscal_labels.py` 重跑，確認 `fiscal_label` 沒被動到。

### 5-3　文件

- `docs/8k-period-off-by-one.md`：加一節「零下載規則（2026-08-25 驗證）」，
  把 §2 的三版比較表與 100%/100% 結果寫進去。原本的「修法建議」三選項
  （A/B/C）需要更新——當初 A 的成本評估是「要下載文件才知道期間」，
  現在證明**不必下載**，A 的成本假設已經不成立。
- `docs/CHANGELOG.md`：待實作後才寫。
- `docs/ARCHITECTURE.md`：`_list_earnings_filings()` 從「純 listing metadata」
  變成「listing metadata + 一次 company 層級查詢」，這條界線要更新。
- `scripts/README.md`：若把驗證腳本收進 `scripts/`，Index 表要補一列。

### 5-4　尚未驗證、實作前要確認的

- **樣本只到最近 8 季（2024–2026）。** 更早期的申報沒測。規則本身與年份無關，
  但 2004-08 之前 Item 2.02 這個編號不存在，那段照樣抓不到。
- **`Company.fiscal_year_end` 會不會隨時間變？** 公司改財年時 EDGAR 只給
  「現在」的值，拿它去回推 10 年前的申報會整段偏掉。26 家沒遇到，但這是
  規則 C 最大的結構性風險，實作前值得先掃一批有改過財年的公司。
- 最大發布延遲實測 58 天。若有公司延遲超過約 70 天（財務重編、延遲申報），
  規則會把它標到下一季。可以在實作時加一道 sanity check。

## 6. 重跑指令

```
./venv/Scripts/python.exe scripts/audit_8k_period_labels.py <out.json> --cache-dir <cache>
# 新聞稿原文已快取在 scratchpad/8k_audit_html/（200 份），改規則重算不必重抓
```
