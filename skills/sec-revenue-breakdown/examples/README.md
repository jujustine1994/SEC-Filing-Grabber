# B formula-v1 驗證案例

`NVDA_Revenue_Breakdown_formula_v1.xlsx` 是新公版使用 NVDA 既有研究資料產生的範例，不是正式公司 B，也不是所有公司的預设分類。固定四頁，展示數字與比重／QoQ／YoY／利益率使用公式；來源與歷史分類保留。

使用方式見 [共用建置入口](../references/build-b-workbook.md)，模板見 [catalog](../assets/templates/catalog.json)。其他公司須建立自己的分類與顯示選擇，不直接改此範例的公司名稱。C 仍維持 ver.1。

## 驗證與重現

- `NVDA_Revenue_Breakdown_formula_v1.verification.json`：原生 Excel 計算與開存後的嚴格對照、成品雜湊。
- `verification-summary.json`：來源雜湊、三模板／三版型合成案例、公用 CLI、篡改拒絕、C 隔離匯入證據。
- 本地同名 `.plan.json`：本次產生的儲存格／來源／公式清單，可交給 validate_workbook_format.py 與 verify_b_excel.py。它可由輸入重建，未納入 Git，以免重複保存大型研究資料。
- 原始研究輸入以既有 `NVDA_segment_history_v2` 的 breakdown／views 和 `NVDA_history_sources/manifest.json` 為準；實際解出的路徑、SHA-256 記在摘要。來源不完整欄位照原狀留空，沒有補猜。

19 個單元測試通過；三份模板及三版型合成案例原生計算通過。NVDA 6,671 個公式、零公式錯誤；C 隔離副本 14 個映射列與 255 個原有歷史值相符。正式 A／B／C 與新來源 B 雜湊未變，沒有外部連結。第二家真實公司尚待驗證，合成案例不冒充真實公司研究。

## 已確認問題與處理

### 2026-10-13 狀態補充

上方 19 測試與第二家公司待驗證是先前紀錄。AMD 已完成真實公司 B 驗證（2,579 公式，原生計算零錯誤）；NVDA 既有結果仍保留。本輪新增 CIK 防護後，24 測試及 13 subtests 通過；NVDA（CIK 1045810，7,771 records）與 AMD（CIK 2488，6,336 records）的全部 plan 儲存格與既有結果一致，未重做 Excel 開存、未改正式工作簿。

NVDA 現用研究輸入已移至公司資料夾「資料／營收拆分／證據」，breakdown.json、views.json、sec_sources/manifest.json 雜湊沿用原證據；舊摘要中的路徑是歷史位置。入口指令已精簡，SEC 下載器與 agent 補查公司 IR 的界線明列，未新增 IR 爬蟲。

1. 作者工具直接填大量公式會逐格求值，耗時／記憶體上升；改為受控公式占位、輸出後只安裝 planner 的公式，交原生 Excel 計算。未通過原生計算的檔案只有結構驗證狀態，不算可發布。
2. 模板示意公式可能殘留於清空後的頁面；輸出修補移除所有不在 plan 的公式，再由嚴格對照攔截任何額外內容。新版 CLI 的合成產檔與計算已通過。
3. 作者工具的字串空白可能在 Excel 開存時被修剪；用 inline string 與 xml:space 保留原始縮排與文字。3 份模板、NVDA 開存後逐格對照通過。
4. 現有 C 以空白列結束父子範圍，且以 Operating income／OP% 定位利益列；移除分類區塊內空白分隔並保留這些列名。區塊間仍保留間距。隔離 C 回歸通過，正式 C 不修改。
5. Windows 作者工具的預覽程序曾在輸出後異常退出；正式入口不再預設執行該預覽。版面改以原生 Excel 的隔離展示副本檢查，讀來源實際計算結果，避免複製公式後相對位置變動使預覽失真。
