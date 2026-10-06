# 營收選值調查（2026-10-06，暫停點）

本次只有唯讀調查，沒有修改正式財報資料庫或產品模板。

> **後續狀態：** 本文件保存改版前的調查，不是目前模板狀態。總計選值及 override 防護已在後續實作；完成範圍、驗收與剩餘問題見 [Revenue 驗證紀錄](revenue-selection-results.md)。下方「下次第一件事」也已完成，來源／解析保存／模板邊界已寫入 ARCHITECTURE.md。

原因：edgartools 5.29.0 將營收構成項和總計都標準化為 Revenue；本專案 `_match_is_row()` 優先取第一個 standard_concept=Revenue 的無維度實體列，成功即不再查原始 Revenues 等總計概念。無維度不代表總計。

MAR accession 0001193125-10-030603：模板選第 1 列 ManagementFeesBaseRevenue，530M；第 7 列 Revenues, Total 為 10.908B。原始概念名稱仍在資料庫。

HLT accession 0001585689-17-000131、2017Q1：目前實際輸出 Revenue=83M；獨立下載原 SEC instance 的相同無維度 2017-01-01 至 2017-03-31 context，基本管理費83M、Revenues總計2.161B，確認同類錯誤。來源：https://www.sec.gov/Archives/edgar/data/1585689/000158568917000131/hlt-20170331.xml

全庫14,417 canonical filings均通過正式讀取閘口；13,143份有IS表。管理費被模板選中：MAR36份、HLT6份。使用實際 `_match_is_row()` 與 `_current_q_col()`，88家公司、2,736份保存IS出現所選列與已辨識標準營收總計數值不同；這是結構風險候選，不是逐份原SEC確認錯誤或最新輸出錯誤率。金融／保險業口徑另需判斷，不能一律換成最大值或任何Revenues列。

本機證據：output/revenue-selection-audit.json、revenue-selection-summary.json、revenue-hlt-original.json；調查腳本在output，未改金融計算。

下次第一件事：釐清改動邊界與資料血緣，確認修正只在模板／選值及衍生輸出，還是需要重新解析／重建資料庫。此營收案例原始concept與數值仍保存，預期可只修選值並重新輸出；KR缺當期facts則可能需要另行恢復解析結果，必須分開討論。變更前後要核对原始檔hash，不直接覆寫歷史來源。
