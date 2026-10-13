# B 公式報表：入口、模板與起始 prompt

## 如何使用

1. 讀 SKILL、data-contract、classification-lifecycle；確認公司、期間與官方來源。views 是顯示選擇，不是 AI 任意公式；每個數字選到 breakdown 的 record ID。新公司需要按官方披露建立 views，不呼叫 NVDA prepare 工具來猜分類。
2. 選 `assets/templates/catalog.json` 的版型：flat（平面，拒絕有父子關係的分類）、hierarchy（階層／多維度）、segment-op（部門及披露利益）。沒有利益資料仍留空，不按版型捏造；同一公司可有多個獨立維度區塊。
3. 執行下列 CLI，使用可用 Python（需 openpyxl 唯讀驗證、lxml 文字保留）與已存在的 artifact-tool runtime；不安裝依賴。out 必須為新檔，拒絕覆寫。

```powershell
python <skill>/scripts/build_b_workbook.py --breakdown <breakdown.json> --views <views.json> --manifest <來源包/manifest.json> --template-id hierarchy --verify-sources --out <新輸出/[Ticker]_Revenue_Segment_History.xlsx>
python <skill>/scripts/validate_workbook_format.py <新輸出.xlsx> --plan <新輸出.plan.json>
python <skill>/scripts/verify_b_excel.py <關閉且隔離的新輸出.xlsx> <新輸出.plan.json>
```

`--verify-sources` 逐檔核對本地原文 SHA-256；省略時不能宣稱原文完整性已核對。`--node-modules <runtime/node_modules>` 可指定已存在的 runtime，不改 repo 套件。生成 .plan.json、.build.json 與 Excel；plan 是本次確定性儲存格／來源／公式清單，並非另一套研究資料庫。

4. 用原生 Excel 在自有副本計算、保存、重新驗證；公式錯誤、來源差額或分類文字被修剪一律不發布。Sources 排序時資料列須整列移動，不只排序單欄；公式以 record ID 找值。主表增列／新期間應重跑 builder，不能只拖移樣板公式便宣稱完成更新。
5. 保留原報表、研究 JSON、來源包及 checks。新 B 驗證後再登記到 C 的 SourceBPath，C 匯入只存計算結果與格式。不要自動覆寫任何已開啟的原始資料。

空白模板的示例公式僅說明輸入／比重，沒有真實期號或分類版本，不是一份完成的研究報表；真正來源與成長公式由 builder 根據明確 period／record／version 產生。AI 不任意增加頁、移動固定標頭或寫死計算值。

## 可複製的起始 prompt

> 請使用 sec-revenue-breakdown skill，為指定公司建立／更新 B 營收拆分報表。先讀公司 README 與現有成果，核對 ticker／CIK、研究期間及來源包；按官方披露整理維度、節點、父子關係與分類版本，不套 NVDA 類別。沿用 breakdown、checks、分類 registry 及 views，逐個顯示值指定 record ID；分類或來源不明標 provisional／unresolved，保留原報與官方重編。依實際結構選 catalog 中的 Excel 版型，使用共用 build_b_workbook.py 產生四頁報表。Sources 保存原始值；營收展示、營收比重、QoQ、YoY、披露利益與利益率使用程式產生的 Excel 公式，不寫死計算結果。核對原文雜湊、分類與來源對照、--plan 嚴格驗證及 Excel 實際計算／存檔；另存新版，通過後才供 C 匯入。沒有新來源時先核對並重用，不重新猜分類。回報範圍、缺口、驗證與檔案位置。

公司及範圍使用當次指示與 README，不必每次重填相同背景；缺少必要資料才詢問。prompt 負責任務入口，精確規範引用本 skill 文件，不重複抄進各公司 README。

## 目前界線

- 分類規範有來源／版本核對，程式能拒絕 identity／parent／期間／來源引用錯誤，不能證明原文分類理解必然正確。
- NVDA 與 AMD 均已有真實公司 B 驗證；歷史原文與期間缺口仍保留。2026-10-13 CIK 防護回歸以兩家公司目前證據重建 plan，所有儲存格與既有 plan 一致；本輪未重做原生 Excel 開存。
- 固定母版若被手動修改，先更新版本與 catalog hash，另跑驗證。builder 不盲目接受其他格式的任意 XLSX。
