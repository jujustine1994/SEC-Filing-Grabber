# 8-K 掃描效率優化 — 設計文件

日期：2026-07-31

## 問題

`_get_earnings_filings()`（`fetcher_nongaap.py:447`）對公司的**每一份**歷史 8-K 呼叫 `filing.obj()`（下載 + 解析），之後才在 `fetch_nongaap_statements()` 裡切 `[:max_filings]` 與套年份過濾。

後果：

- 下載量與「公司發過幾份 8-K」成正比，而非與「使用者要幾季」成正比
- 年份區間縮小**不會**縮短掃描時間，只減少 AI 呼叫次數
- 首次抓一家公司需 5–10 分鐘，抽檢多家公司不可行

實測白費比例：

| 公司 | 全部 8-K | 含 Item 2.02 | 白下載 |
|---|---|---|---|
| AAPL | 235 | 94 | 60% |
| CRM | 290 | 88 | 70% |
| PANW | 143 | 57 | 60% |
| ARLO | 67 | 32 | 52% |

## 依據：申報清單已含判斷所需欄位

`company.get_filings(form="8-K").data` 為 pyarrow table，欄位包含：

```
accession_number, filing_date, reportDate, acceptanceDateTime, act, form,
fileNumber, items, size, isXBRL, isInlineXBRL, primaryDocument, primaryDocDescription
```

- `items` — 逗號分隔的 8-K 項目代號，如 `2.02,9.01`。財報發布為 **Item 2.02**
- `reportDate` — 期間結束日，可直接算季度標籤，取代下載後讀 `period_of_report`

兩者皆來自 SEC submissions JSON，**取得清單即有，零額外下載**。

## 設計

### 快速路徑（預設）

`_list_earnings_filings(company, start_year, end_year, max_filings)` 在清單階段完成全部篩選，**不下載任何檔案**：

1. `items` 含 `2.02` → 財報 8-K
2. 由 `reportDate` 算季度標籤（沿用 `_period_to_quarter_label()`）
3. 依標籤去重（同季多份取最舊，維持現行語意）
4. 套 `start_year` / `end_year`
5. 切 `max_filings`

`fetch_nongaap_statements()` 再扣掉 `nongaap_cache.json` 已有的季度，**只對剩下的呼叫 `filing.obj()`**。

效果：AAPL 抓 4 季由 235 次下載降為 4 次；首跑時間由 5–10 分鐘降至 30–60 秒（含 AI 呼叫）。

### 缺季自動補掃（取代「完整掃描」開關）

快速路徑挑出的季度標籤排序後檢查連續性。發現缺口（如有 FY2023Q1、FY2023Q3 而無 Q2）時，**僅對該缺口的日期區間**回退舊做法逐筆 `filing.obj()`，用 `has_earnings` 尋找漏標的財報。找到則補入，找不到則在 log 記錄該季確實不存在。

不做 GUI checkbox。理由：使用者無從判斷該不該勾（勾了慢、不勾怕漏），而缺季是可自動偵測的客觀條件；成本只在真的缺季時付出。

### 已知邊界：2004-08 之前抓不到

SEC 自 2004-08-23 起才啟用 `2.02` 編號，之前為舊制（財報為 Item 12 或 Item 5）。實測 AAPL 1996–2004 的 26 份 8-K，代號為 `2`、`5,7`、`12,7`、`7,9` 等，無法以 `2.02` 匹配。

不處理。`max_filings` 預設 80 筆約當 20 年，2004 至今 22 年已覆蓋預設範圍。缺季補掃只在區間內運作，不回溯至 2004 之前。

### 完整掃描的實測價值

以 ONTO 驗證：153 份未標 `2.02` 的 8-K 全數下載拆解，僅 1 份 `has_earnings` 為真（2021-04-29，`items=8.01,9.01`），且極可能為誤判（`8.01` 為其他事件，`has_earnings` 命中文字即亮）。

結論：SEC 標記可信度足夠，全量深掃成本 153 次下載、收穫 0–1 份可疑品，不值得作為常態選項。

## 介面相容性

- `_get_earnings_filings(company)` 現行簽章僅收 `company`，回傳 `list[tuple[label, filing, eight_k]]`
- 新函式改回傳 `list[tuple[label, filing]]`（不含 `eight_k`，因為尚未下載）
- `fetch_nongaap_statements()` 內部迴圈改為在需要時才 `filing.obj()`
- `fetch_nongaap_statements()` 對外簽章不變，GUI 與既有測試不受影響

## 測試

TDD，先寫測試再改實作。

**單元測試（不連網，以假清單 DataFrame 驅動）**

1. `items` 含 `2.02` 者入選，不含者排除
2. `items` 為空或 None 不拋例外
3. 同季多份 8-K 去重後保留最舊
4. `start_year` / `end_year` 正確過濾
5. `max_filings` 正確切割且取最新
6. 季度標籤由 `reportDate` 正確推算（含非 12 月財年）
7. 缺季偵測：連續季度回傳空缺口；中間缺一季回傳該季
8. 缺季偵測不把「最舊之前」與「最新之後」當缺口

**整合測試（連網，標記 `slow`）**

9. ARLO 快速路徑挑出的季度數與現行深掃結果一致
10. 抓指定 4 季時，`filing.obj()` 呼叫次數 ≤ 4

**驗收**

以 CRM（1 月財年）、PANW（7 月財年）、ARLO（小型股）各抓 4 季，確認首跑 60 秒內完成，季度標籤與 8-K 原文期間相符。

## 不做

- 不改 GAAP 路徑（`fetcher_gaap.py` 抓 10-Q/10-K，本來就有清單層 form 過濾，無此問題）
- 不改 GUI
- 不動 `nongaap_cache.json` 格式
