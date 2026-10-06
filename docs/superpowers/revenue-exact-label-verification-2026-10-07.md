# Revenue 完整名稱規則全庫驗證

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

新增 RED 測試 commit `4eb35a5` 重現，再排除完整成本 label `Cost(s) of Revenue(s)`／`Cost(s) of Sales`；不更動已知總計優先序、不接受部分 Revenue 字串、不新增金融業口徑推論。修正後非 slow suite：**1,764 passed / 65 deselected / 5 warnings**，85.77 秒。全庫影響掃描及受影響公司再驗證尚在進行。

APH 等歷史 `Net sales`、不同產業的特殊總額名稱與銀行分項仍須另外核對精確名稱規則；不能把恢復舊版數字當成來源驗證。COP／OXY／TGT 原有 29 筆口徑衝突仍未解。
