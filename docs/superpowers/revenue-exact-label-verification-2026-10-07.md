# Revenue 完整名稱規則全庫驗證

**最終數值版本：`bdc930d`。** 215 家結果完整；本次只重建 JSON 形式的模板成果，沒有批次產生全部公司的 Excel，也沒有修改來源資料庫。精簡機器證據見 [verification summary](evidence/revenue-exact-label-verification-summary.json)。第一輪與各中間版本數字只作調查歷史，最新統計以下方「最終驗收」為準。

## 第一輪固定版本

金融程式版本 `2c81630`，來源為正式庫既有 JSON，沿用凍結的官方 accession/reportDate 清單與 80 份季報／20 份年報限制。AI diagnosis、companyfacts 股數、資料庫保存 hook 與 gap 網路探測均停用。結果寫入專案 `output/`，不更新來源庫。

- 215 家完整重建，22 個批次全部正常結束，缺公司／失敗 sidecar／比較例外為 0。
- 487 張財務表固定列檢查通過；7 個原 SEC 營收樣本及 15 個來源衍生拆季／CF 格通過。
- 正式庫非 `.git` 檔案 29,277 個、3,204,105,955 bytes，前後逐檔 SHA-256 完全相同，新增／刪除／變更皆 0。
- 與 `output/revenue-accepted-all` 比較，30 家有數值差異。固定列只有 Revenue（季度 914、年度 231）及連動 Gross Profit（季度 233、年度 62）；其他固定指標沒有數值差異。期間標籤及表頭差異為 0。
- `AmbiguousRevenueTotal` 從 29 增至 153，不能全算成新錯誤：含成本誤判與拒絕舊分項選值兩類；MissingStandalonePeriod 230、MissingCurrentPeriod 86 不變。
- 包含前後全檔案雜湊的第一輪耗時 1,462.51 秒（24.38 分鐘）。此值不是纯模板運算時間，也不含全部 Excel 匯出／重算。

大型證據：`output/revenue-exact-label-all`、`revenue-exact-comparison`、`revenue-exact-analysis.json`、`revenue-exact-run-summary.json` 與前後 database manifests。前一輪大量样本驗收與這輪結果分開保存。

## 樣本發現的修正

AMD 的原始 `SalesRevenueGoodsNet` label 為完整 `Revenue`，同表 `Cost of Revenue` 被寬鬆競爭候選檢查當成另一個營收。這是本次新增誤判，不是原 SEC 或解析 JSON 的問題。

新增 RED 測試 commit `4eb35a5` 重現，再排除完整成本 label `Cost(s) of Revenue(s)`／`Cost(s) of Sales`；不更動已知總計優先序、不接受部分 Revenue 字串、不新增金融業口徑推論。首次修正後非 slow suite：**1,764 passed / 65 deselected / 5 warnings**，85.77 秒。後續成本註腳與處分 concept 及其歷史名稱另以精確原始 GAAP concept 排除，最終結果如下。

APH 等歷史 `Net sales`、不同產業的特殊總額名稱與銀行分項仍須另外核對精確名稱規則；不能把恢復舊版數字當成來源驗證。COP／OXY／TGT 原有 29 筆口徑衝突仍未解。

## 最終驗收

- 先完整重建 215 家，再掃描全庫 14,417 份原有 filing JSON；成本／處分排除相關 22,005 個來源期間比對，僅 9 家選值器結果改變，已用 `8f2f32e` 重建這 9 家。後續兩個歷史處分 concept 的精確 alias 另掃全庫，相關 722 個期間中只有 CMG／HCA 受影響，再用 `bdc930d` 重建兩家。每一步僅在確認選值不變後保留其餘公司輸出；不是宣稱全部公司都由最後一次程序重新執行。
- 最終輸出 `output/revenue-exact-final-all`，穩定身份比對 `output/revenue-exact-final-comparison`。215 對 215，缺公司、失敗 sidecar、比较例外皆 0；官方 metadata 完整，使用正式深度。
- 487 張財務表固定列通過；7 份原始 SEC XML 獨立重讀通過，7 個營收輸出與 15 格拆季／CF 來源衍生樣本通過。抽樣來源認證不代表全庫各格已核對原申報。
- 非 slow suite **1,770 passed / 65 deselected / 5 warnings**，77.86 秒；`git diff --check` 通過。未執行 65 個 slow 測試。
- 與前一輪 `revenue-accepted-all` 比較，26 家有數值變動。Revenue 季度 845／年度 215，Gross Profit 季度 179／年度 48；BS／CF 與其他固定指標沒有新數值差異，期間標籤及表頭差異 0。
- Revenue 的 1,060 筆比較記錄分為 880 筆留空、176 筆改值、4 筆補值；這是固定列／期末日的比較記錄，不是獨立 facts 數量或來源正確率。未辨識歷史名稱、拒絕舊分項與不同口徑都可能造成留空，不能把 880 筆全部稱作改善。
- 最終 `AmbiguousRevenueTotal` **85**：COP 24、OXY 4、TGT 1 是原有 29 筆；新增待核對 BK 9、CMG 7、HCA 26、MS 12、NEM 2 共 56 筆。這是 ledger 衝突記錄，**不是全部 Revenue 空值數**；缺乏精確名稱時的未匹配另有空值，下一輪需補規則覆蓋。
- `MissingStandalonePeriod` 230、`MissingCurrentPeriod` 86 不變。新衝突尚未逐筆原 SEC 複核，不宣稱每筆都是真正口徑衝突。
- 所有重建、來源掃描與比較後再逐檔 SHA-256：正式庫 **29,277 檔、3,204,105,955 bytes，新增／刪除／變更 0**。不清理 edgartools HTTP 快取，不更新來源 JSON。
- 先前七本驗證 Excel 的完整輸入表（AFL／AMZN／C／COST／HLT／MAR／MSFT）與最終結果完全相同；本次沒有新增 workbook 或重新 Excel 重算，不把舊工作簿驗收推廣成全 215 家 Excel 已完成。
- 執行期公司處理及比較無 LLM 呼叫。第一輪重建加前後雜湊約 24.38 分鐘；新增漏洞調查、回歸與影響掃描是額外工時，不能混成純模板跑完的時間。

## 下一步

1. 先補 `Net sales` 等精確完整名稱與產業總額 concept 的來源核對；來源規則應同時防止分項重新冒充總額。
2. 分類新增 56 筆衝突，優先 CMG 殘留候選與 HCA 扣除呆帳前後口徑，再處理 BK／MS／NEM；核對原 accession/context，不從後年比較欄任選申報作證。
3. 延續 COP／OXY／TGT 原有 29 筆、KR 當期 facts、KHC 併購口徑與 GUI 一致性。營收覆蓋仍未完整，不把當前全公司成果當成已認證的最終 Excel 批次。

## Checkpoints 與重現

`4eb35a5` 成本候選 RED；`a36d871` 完整成本名稱排除；`8f2f32e` 精確 GAAP expense／處分排除；`bdc930d` 處分歷史 concept 名稱。每一步測試與證據分開保存，中間輸出不代替最終輸出。

在具有相同凍結官方 listing 的 `output/` 中，可用現有唯讀腳本從最終版本重新驗證全部公司；新目錄勿覆蓋本輪證據：

```powershell
$revenuePython = Join-Path $env:USERPROFILE 'venvs/SEC Financial Tools/Scripts/python.exe'
$env:PYTHONIOENCODING = 'utf-8'
$env:PYTHONDONTWRITEBYTECODE = '1'
$env:EDGAR_LOCAL_DATA_DIR = Join-Path (Get-Location) 'output/database-verification-edgar'
& $revenuePython scripts/audit_fiscal_pipeline.py --all --require-official-metadata --output output/revenue-exact-reproduced-all
& $revenuePython scripts/compare_revenue_audits.py output/revenue-accepted-all output/revenue-exact-reproduced-all --output output/revenue-exact-reproduced-comparison
& $revenuePython scripts/verify_fixed_financial_rows.py output/revenue-exact-reproduced-all
& $revenuePython scripts/verify_revenue_source_examples.py output/revenue-exact-reproduced-all
& $revenuePython scripts/verify_fiscal_source_examples.py output/revenue-exact-reproduced-all
& $revenuePython -m pytest -m 'not slow' -q -p no:cacheprovider
```
