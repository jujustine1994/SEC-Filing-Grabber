# Scripts

開發與維護用的獨立腳本。不屬於主程式流程，可單獨執行。

## Index

| 腳本 | 說明 | 狀態 |
|------|------|------|
| `smoke_test_10.py` | 批次 live smoke test：10 間公司各抓 GAAP，自動檢查 Revenue/Gross Profit/Operating Income/Net Income/OCF/Capex/FCF 是否有值，輸出彙總表 | 啟用 |
| `audit_8k_period_labels.py` | 量化 Item 2.02 8-K 季度標籤的 off-by-one（TODO D4 前半）。用新聞稿內文判定實際財期，與 `_period_to_quarter_label()` 現行標籤比對；同時查 dedupe「保留最舊」丟掉哪些財報。**純文字比對，不呼叫 AI**。結論見 `docs/8k-period-off-by-one.md` | 啟用 |
| `verify_8k_fiscal_labels.py` | 驗證 `src/cli.py press-release` 的 `fiscal_label`（TODO D4 後半）。15 家 × 8 季，檢查期末日抓取率，並確認新舊標籤的偏移**同一家是常數**且等於 `docs/8k-period-off-by-one.md` 獨立推出的值。**不呼叫 AI**，約 3 分鐘。沒填過進階設定時第一個參數當 identity | 啟用 |
| `excel_golden.py` | **Excel 輸出的逐格回歸驗收。改 `excel_writer` / `excel_formatter` / `ratios` / `fiscal_input` 之前先 `make` 一份基準，改完再 `make` + `check`。** 把 `output/_final/*.xlsx` 讀回成 StatementTable 走真正的寫檔流程重產，比對值＋數字格式＋字型＋粗體＋底色。不打網路。2026-08-14 多語言遷移就是靠它確認 480 條字串搬完後繁中輸出逐格不變 | 啟用 |
| `gen_zh_cn.py` | 從 `src/locales/zh_tw.py` 重產 `zh_cn.py`（OpenCC tw2s + 自訂詞彙表）。**一次性工具，平常不跑**，且會覆蓋手改過的用詞。需要 `uv pip install opencc-python-reimplemented`（不在 requirements，執行期用不到） | 啟用 |
| `打包.bat` + `pack.ps1` | **打包散布用 zip，雙擊 `打包.bat` 即可，不必開 PowerShell、不必叫 AI。** 白名單複製 → 壓成 `dist\SEC-Financial-Fetcher-YYYYMMDD.zip` → 解到暫存目錄跑 **12 項自我驗證**（機敏檔、`.xlsx`、金鑰樣式、非預期 email、內部文件外流…）。**任一項沒過就刪掉 zip 並 exit 1**，不會產出一包不能傳的東西。是 `docs/PACKAGING.md` 的可執行版本，兩邊改動必須同步 | (停用，2026-08-18 CTH：GitHub 連得回來了，改走 clone/pull 發布，zip 打包暫不用；腳本保留待日後需要) |
| `gen_template_coverage_baseline.py` | **產出模板體檢基線** `docs/template-coverage-baseline-<日期>.md`：每家公司的缺漏判斷、最常出問題的列、逐列覆蓋率（現行路徑 vs companyfacts）。不打網路，吃 `output/_spike/` 的快取。改了模板 concept 對照或 `data_quality` 判斷規則之後重跑，看「40/97」這個數字有沒有往上走 | 啟用 |
| `spike_verify_mapping.py` | **TODO G11 第四步（驗收）**：用 `src/facts_mapping.py` 實跑，跟現行路徑的快取答案卷逐格比對。**完全不打網路**，幾秒跑完 52 家。分開統計「數字相同」與「只差正負號」——那是慣例對不齊不是抓錯，處理方式完全不同。輸出 `output/_spike/verify_mapping.xlsx`（每列×每家命中率，<80% 紅、<95% 黃） | 啟用 |
| `spike_validate_facts.py` | **TODO G11 第三步**：不依賴現行路徑的獨立驗證（反推有個先天限制——現行路徑錯的地方比對也會跟著錯）。四項檢查：會計恆等式、四季加總=年度、**SEC 官方 `frame` vs 我們的期中點判準**、重編頻率。只打 companyfacts 所以能一次驗二三十家。實測 24 家：frame 對齊 59,564/59,564 全數一致 | 啟用 |
| `spike_companyfacts_diff.py` | **TODO G11 決策依據**：SEC companyfacts API 路徑 vs 現行「逐份解 filing」路徑的逐格比對。不改任何現有程式。輸出耗時對照、逐格異同、模板列覆蓋率、Q4 由 10-K 直接 tag 的比例。實測 NVDA：companyfacts 0.51s vs 解 filing 109s（**215 倍**） | 啟用 |
| `spike_derive_mapping.py` | **TODO G11 第二步**：用現行路徑已知正確的數字當答案卷，反推 companyfacts 的 us-gaap concept 對照表。模板的 `std_concept` 欄是 edgartools 正規化過的名字（`NetIncome`），不是原始 element name（`NetIncomeLoss`），憑印象填 75 列一定會錯。這支對每個 concept 算「同期末日數字對得上的比例」，命中率最高的就是正確 mapping，順便偵測正負號相反。結果存 `output/_spike/mapping_candidates.json`，並快取 facts JSON 與現行路徑結果避免重跑 | 啟用 |
| `survey_nongaap_metrics.py` | 調查 32 家（大中小型跨產業）8-K 新聞稿實際使用的 Non-GAAP 指標，統計跨公司覆蓋率，決定 `Data_NonGAAP` 固定模板要收哪些行。**不呼叫 AI**（純文字比對，不吃配額）。原文會存到快取目錄，調整比對規則後可重跑分析不必重新下載 | 啟用 |

---

## 測試方案對比：smoke_test_10.py vs tests/test_live_snapshots.py

| | `scripts/smoke_test_10.py` | `tests/test_live_snapshots.py` |
|---|---|---|
| **執行方式** | `python scripts/smoke_test_10.py` | `python -m pytest -m slow` |
| **用途** | 人工快查：最新季數值有沒有抓到 | 自動化迴歸：程式行為是否符合預期 |
| **公司** | AAPL/MSFT/TSLA/AMD/NVDA/GOOGL/META/WMT/COHR/AMZN（10 間） | MSFT/AMZN/META/GOOGL/NVDA/JPM/GS/JNJ（8 間）+ CF overflow 組（COHR/LITE/AAPL/NVDA/GOOGL） |
| **抓取筆數** | `max_filings=80`（完整，偵測長期資料） | `max_filings=8`（只抓最新 8 季，省時） |
| **輸出** | 彩色 terminal 表格，直接顯示數值 | pytest PASS/FAIL，失敗才顯示原因 |
| **判斷標準** | 最新季 7 個 key rows（Revenue/Gross Profit/Operating Income/Net Income/OCF/Capex/FCF）全非 None | key rows 近 4 季有 ≥1 非 None + B1 overflow 結構完整 + CF YTD 減法正確 |
| **耗時** | 較長（完整抓取） | 約 12 分鐘（8 季/ticker） |
| **適合場景** | 開發後手動驗收、懷疑某公司資料有問題時 | 改動 fetcher 後確認沒有迴歸 |
