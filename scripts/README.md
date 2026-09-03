# Scripts

開發與維護用的獨立腳本。不屬於主程式流程，可單獨執行。

## Index

| 腳本 | 說明 | 狀態 |
|------|------|------|
| `smoke_test_10.py` | 批次 live smoke test：10 間公司各抓 GAAP，自動檢查 Revenue/Gross Profit/Operating Income/Net Income/OCF/Capex/FCF 是否有值，輸出彙總表 | 啟用 |
| `audit_8k_period_labels.py` | 量化 Item 2.02 8-K 季度標籤的 off-by-one（TODO D4 前半）。用新聞稿內文判定實際財期，與 `_period_to_quarter_label()` 現行標籤比對；同時查 dedupe「保留最舊」丟掉哪些財報。**純文字比對，不呼叫 AI**。結論見 `docs/8k-period-off-by-one.md` | 啟用 |
| `verify_8k_fiscal_labels.py` | 驗證 `src/cli.py press-release` 的 `fiscal_label`（TODO D4 後半）。15 家 × 8 季，檢查期末日抓取率，並確認列清單的 `label` 與 `fiscal_label` **偏移全部是 0**（B5 之後兩條路本來就該對齊；B5 前每家各自是常數偏移 -3~+1，舊值留在腳本的 `LEGACY_OFFSETS`）。同時是 B5 的端對端驗收與 `fiscal_label` 的回歸。**不呼叫 AI**，約 3 分鐘。沒填過進階設定時第一個參數當 identity | 啟用 |
| `excel_golden.py` | **Excel 輸出的逐格回歸驗收。改 `excel_writer` / `excel_formatter` / `ratios` / `fiscal_input` 之前先 `make` 一份基準，改完再 `make` + `check`。** 把 `output/_final/*.xlsx` 讀回成 StatementTable 走真正的寫檔流程重產，比對值＋數字格式＋字型＋粗體＋底色。不打網路。2026-08-14 多語言遷移就是靠它確認 480 條字串搬完後繁中輸出逐格不變 | 啟用 |
| `gen_zh_cn.py` | 從 `src/locales/zh_tw.py` 重產 `zh_cn.py`（OpenCC tw2s + 自訂詞彙表）。**一次性工具，平常不跑**，且會覆蓋手改過的用詞。需要 `uv pip install opencc-python-reimplemented`（不在 requirements，執行期用不到） | 啟用 |
| `打包.bat` + `pack.ps1` | **打包散布用 zip，雙擊 `打包.bat` 即可，不必開 PowerShell、不必叫 AI。** 白名單複製 → 壓成 `dist\SEC-Financial-Fetcher-YYYYMMDD.zip` → 解到暫存目錄跑 **12 項自我驗證**（機敏檔、`.xlsx`、金鑰樣式、非預期 email、內部文件外流…）。**任一項沒過就刪掉 zip 並 exit 1**，不會產出一包不能傳的東西。是 `docs/PACKAGING.md` 的可執行版本，兩邊改動必須同步 | (停用，2026-08-18 CTH：GitHub 連得回來了，改走 clone/pull 發布，zip 打包暫不用；腳本保留待日後需要) |
| `gen_template_coverage_baseline.py` | **產出模板體檢基線** `docs/template-coverage-baseline-<日期>.md`：每家公司的缺漏判斷、最常出問題的列、逐列覆蓋率（現行路徑 vs companyfacts）。不打網路，吃 `output/_spike/` 的快取。改了模板 concept 對照或 `data_quality` 判斷規則之後重跑。**看數字前先讀產出文件的第零節**——達標列數只是體溫計，真正的 KPI 是〔真缺口〕與〔假警報〕，而且永遠不該以 97/97 為目標 | 啟用 |
| `diag_probe.py` | **最常用的排查工具**：印出某家公司某張報表裡符合正則的列，含 `concept`／`standard_concept`／`label`／數值。`ARCHITECTURE` 那套「三步排查順序」的第 2、3 步就靠這支——先確認那一列在不在報表 dataframe 裡，再看 matcher 為什麼沒命中。2026-08-23 的 H3 修復全部從這裡開始 | 啟用 |
| `diag_rowprobe.py` | 某個**模板列**在多家公司的命中情況，同時列出 dataframe 裡所有長得像的候選。判斷一列的 concept 對照要不要改，看這支的輸出最快。**H6（2026-08-25）的抽查全靠它**：改 hint 前先用它回頭核對原始 10-Q，確認那一列真的是要的科目（UNP/IP/LIN/EXC），也是靠它推翻分類表「CS&APIC 那 7 家是 concept 層失守」的判讀 | 啟用 |
| `diag_hintsweep.py` | 掃出「`label_hint` 太窄把正確答案濾掉」的模板列——比較有 hint 與沒 hint 的命中差異。**2026-08-23（H3，22 家）與 2026-08-25（H6，201 家）兩輪修復都是從這支開始**。H6 用它量出前後：Capex 15→3、CS&APIC 14→2、Cash 20→15、Cost of Revenue 36→30，其餘 10 條有 hint 的列一筆沒變。201 家清單與原始輸出留在 `output/_hintsweep_201/`（**不要重跑，一輪 12 分鐘**）：`TICKERS=$(cat output/_hintsweep_201/tickers_joined.txt)`。⚠ 有些 hint 是必要的（擋現金流量表末尾的租賃補充揭露列、擋庫藏股列、擋銀行的人事費），不要看到被殺就拿掉——**改完一定要重跑一次比對 killed 清單有沒有長出新的** | 啟用 |
| `check_fye_drift.py` | **量「公司改過財年」對 8-K 零下載規則的風險（TODO B6）**。B5 的季度標籤是用 EDGAR 的 `fiscal_year_end` 現值回推的，公司改過財年的話，改制以前的申報會整段標錯。這支從 `output/_spike/facts_*.json` 取每份 10-K 的財年結束日，跟最新那年的月日比，超過門檻（預設 14 天，52/53 週制本來就會浮動 7 天）就判定改過財年。**零網路請求**，201 家跑幾秒。同時量「改制前那些季會標錯幾季」（發布日用「期末日 + 28 天」代入）。2026-08-25 首跑：201 家裡只有 **LHX**（2019 從 6 月底改成 12 月底，改制前 30 季**全錯**，一律差 2 季）與 **MSCI**（2010 從 11/30 改成 12/31，位移 31 天不跨季，0 季錯）兩家；其餘 199 家最大偏移都在 9 天內（52/53 週制的正常浮動） | 啟用 |
| `diag_celldiff2.py` | **改 concept 對照的回歸驗收**：兩份答案卷快取逐格比對。改之前先把 `output/_spike` 複製一份當基準，改完重建再比；驗收標準是「不能有任何一格從有值變成不同的值或空」。⚠ 鍵用 `(列名, 第幾次出現)`——`Net Income`／`SBC` 在 IS 和 CF 各有一列，用列名當鍵會拿 IS 那列比 CF 那列，2026-08-24 這樣憑空生出 3,659 個假異動 | 啟用 |
| `spike_verify_mapping.py` | **TODO G11 第四步（驗收）**：用 `src/facts_mapping.py` 實跑，跟現行路徑的快取答案卷逐格比對。**完全不打網路**，幾秒跑完 52 家。分開統計「數字相同」與「只差正負號」——那是慣例對不齊不是抓錯，處理方式完全不同。輸出 `output/_spike/verify_mapping.xlsx`（每列×每家命中率，<80% 紅、<95% 黃） | 啟用 |
| `spike_validate_facts.py` | **TODO G11 第三步**：不依賴現行路徑的獨立驗證（反推有個先天限制——現行路徑錯的地方比對也會跟著錯）。四項檢查：會計恆等式、四季加總=年度、**SEC 官方 `frame` vs 我們的期中點判準**、重編頻率。只打 companyfacts 所以能一次驗二三十家。實測 24 家：frame 對齊 59,564/59,564 全數一致 | 啟用 |
| `spike_companyfacts_diff.py` | **TODO G11 決策依據**：SEC companyfacts API 路徑 vs 現行「逐份解 filing」路徑的逐格比對。不改任何現有程式。輸出耗時對照、逐格異同、模板列覆蓋率、Q4 由 10-K 直接 tag 的比例。實測 NVDA：companyfacts 0.51s vs 解 filing 109s（**215 倍**） | 啟用 |
| `spike_derive_mapping.py` | **TODO G11 第二步**：用現行路徑已知正確的數字當答案卷，反推 companyfacts 的 us-gaap concept 對照表。模板的 `std_concept` 欄是 edgartools 正規化過的名字（`NetIncome`），不是原始 element name（`NetIncomeLoss`），憑印象填 75 列一定會錯。這支對每個 concept 算「同期末日數字對得上的比例」，命中率最高的就是正確 mapping，順便偵測正負號相反。結果存 `output/_spike/mapping_candidates.json`，並快取 facts JSON 與現行路徑結果避免重跑 | 啟用 |
| `survey_nongaap_metrics.py` | 調查 32 家（大中小型跨產業）8-K 新聞稿實際使用的 Non-GAAP 指標，統計跨公司覆蓋率，決定 `Data_NonGAAP` 固定模板要收哪些行。**不呼叫 AI**（純文字比對，不吃配額）。原文會存到快取目錄，調整比對規則後可重跑分析不必重新下載 | 啟用 |
| `check_excel_repair.ps1` | **驗證一份 `.xlsx` 會不會被 Excel 判定內容毀損**（TODO A/F8 修復用）。用 Excel COM 開檔，比對 `%TEMP%` 底下 `error*.xml` 修復日誌開檔前後的變化，並清點 `Chart_*` 分頁還剩幾張圖。**實測發現**：壞檔會讓 `Workbooks.Open()` 直接丟 COM 例外（不是卡對話框，Open() 呼叫本身快速失敗），這個訊號比等修復日誌更乾淨，用來跟已知正常的 `.xlsx` 對照最快。`powershell -File scripts/check_excel_repair.ps1 -Path <絕對路徑>`，回傳碼 0=OK、2=REPRODUCED（判定毀損） | 啟用 |

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
