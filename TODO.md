# TODO

1. 確認 Excel 排版現況：實跑一次輸出，逐 sheet 檢視 `Data_Financials(Q/Y)`、`Data_Financials_NG(Q/Y)`、`Data_Seg_*`、`Index` 的欄寬、凍結窗格、數字格式（÷1M 與 EPS 例外）、section 分隔行、subtotal 粗體是否都正確，記錄實際問題再決定要不要調整 `excel_formatter.py`。
2. ~~**`Data_NonGAAP` 資料品質修復**~~ ✅ **已完成（2026-08-01 下午，方案 c）**，詳見 CHANGELOG。規則表集中在 `metric_rules.py`，改表後重跑即生效、不必重抓。ARLO 實跑對 8-K 原文逐項核對全中。**殘留待決定事項（都有現成開關，不急）**：
   - ~~年度值補洞~~ ✅ 已改為 `FY_ONLY_HANDLING = "label"`：年度值另成一列加 ` (FY)`，不再填進季欄位。
   - ~~百分比存法~~ ✅ 已改為 Excel 原生比例（`PERCENT_AS_EXCEL_RATIO = True`，0.375 + `0.0%`）。
   - ~~同義名合併~~ ✅ 已拆開：查原文確認 ARLO 2025 年把 `service gross margin` 改名為 `subscriptions and services gross margin`，營收基礎同時變動，不可視為同一條線。
   - **對照表覆蓋率**：目前只依 ARLO / CRM / PANW 三家的實際輸出建表。跑到新公司時若出現沒收錄的中文名，會**原樣顯示中文**（不會丟資料，但那一列不會跟英文季合併）。第 3 項抽檢時順便擴充。
   - **CRM／PANW 尚未用新 prompt 重抓**：2026-08-01 驗證時 Gemini 配額（`HTTP 429`）用完，這兩家目前顯示的是舊中文快取經對照層救回的結果，已可讀但比 ARLO 稀疏，且 CRM 的 `Free Cash Flow` 在 `FY2026Q1` 是年度值 14.4B（舊 prompt 沒有「只取當期」約束）。配額恢復後各跑一次 `--fresh` 即可。
   - ~~AI 呼叫沒有重試機制~~ ✅ 已加：`AI_MAX_ATTEMPTS = 3`、退避 5s／15s，跑完列出未取得季度並推給 `progress_cb`。次數刻意壓低——Gemini 每日配額型 429 重試必敗，只有每分鐘限流型救得回來。
   - **多把 API key 輪替尚未實作**：配額用盡時目前要手動去進階設定換 key。若常撞到，可考慮讓 config 收多把 key 自動輪替。
3. **期末流通股數取不到（2026-08-02 實查，`Data_Ratios` 有 3 列因此永遠空白）**：BS 模板已加 `Shares Outstanding`，對映 `CommonStockSharesOutstanding`，但實測 ARLO / AAPL / NVDA / MSFT / COHR **五家全部沒有在資產負債表裡 tag 這個概念**——股數只寫在 `CommonStockValue` 的 label 文字裡（「shares issued and outstanding: 108,745,373 at March 29, 2026」）。正確來源是封面頁的 `dei:EntityCommonStockSharesOutstanding`，在 edgartools 走 `Company.get_facts().dei_facts`，與現行 `balance_sheet` 取數路徑不同，需另接。未接通前 `BVPS ($)`、`FCF per Share ($)`、`流通股數 YoY (%)` 三列恆為空。替代方案是改用 `Diluted Shares`（加權平均）當分母，但那與「期末股數」定義不同，需先確認可接受。

4. 8-K 抽取正確性抽檢：AI 隨機挑一批公司（跨產業、跨市值，含非 12 月財年）跑 Non-GAAP 流程，比對 `Data_NonGAAP` 抽到的指標與 8-K press release 原文，確認項目沒漏抓、沒誤抓、期間標籤（FY/Q）對得上，把失敗案例整理成清單。**另需特別檢查**：`_period_to_quarter_label()` 是用 Item 2.02 8-K 的 `period_of_report`（EDGAR 上這欄放的是**發布/事件日**，不是財報所屬的財期結束日）去分季。實測驗證：INTC 的 `20260723` 被標成 `FY2026Q3`，但那份新聞稿報的其實是 FY2026 Q2（~6/27 結束）；COST 的 `20260528` 被標成 `FY2026Q2`，實際報的是 FY2026 Q3（5 月中結束）；WDC 的 `20260430` 被標成 `FY2026Q2`。也就是說 `Data_NonGAAP` sheet 上**每一欄的標籤幾乎都比數字實際所屬的財季晚了一季**，是系統性的 off-by-one，不只是「日曆季 vs 公司財季」的差異。這是長期存在的行為（非本次分支引入），但範圍不小，抽檢時要一併確認影響程度並評估修法。**另外，同一根因還會造成漏抓**：WDC 曾在同一個日曆季（2025 Q1）內發布兩份 Item 2.02 8-K（`20250110` 與 `20250129`），兩者都被標成 `FY2025Q1`，dedupe 邏輯「保留最舊那份」會直接把 1/29 那份財報悄悄丟掉，這點也要一併排查。
5. `max_filings` 不再是硬上限：`_list_earnings_filings()` 先套用切片，缺季回補（`_recover_missing_quarters()`）在補齊後不會重新裁切，故要求 8 季、若保留區間有 2 個缺口，實際可能下載到 10 份。需評估是否要在回補後補一次裁切，或至少在文件中說明此行為。
6. CLI 工具層（`cli.py`）：讓外部 skill 用指令調用現有 fetcher，不經 GUI。例如 `python cli.py nongaap NVDA --years 2020-2026 --json`、`python cli.py gaap AAPL --xlsx out.xlsx`；輸出支援 JSON（給 skill 讀）與 Excel（給人看）。GUI 與核心函式不動，只加一層薄封裝。等「Excel 排版確認」與「8-K 抽取正確性抽檢」驗收完再做。
7. 套件汰換：`google-generativeai` 官方已終止支援（不再更新與修 bug），需改用 `google-genai`。影響 `fetcher_nongaap.py:_call_ai()` 與 `override_engine.py:_llm_call()` 的 google 分支，以及 `requirements.txt`。現行版本仍可運作，非緊急。
8. 金融股（GS/JPM 等）獨立模板：現行 IS/BS 模板對金融股部分欄位空白，需另建模板。低優先。
9. 重複抓同一 ticker 的檔案行為需補防護：`excel_writer.write_statements()` 開啟既有檔後刪除所有 `Data_*` sheet 重寫（`My_*` 等自訂 sheet 保留），無備份無版本號。兩個風險：(a) 第二次抓的年份範圍較窄時，`Data_*` 是整批替換而非合併，舊季度直接消失（GAAP 無快取需重抓）；(b) 該 xlsx 正被 Excel 開啟時 Windows 鎖檔，`wb.save()` 拋 `PermissionError`，但這發生在全部抓取與 AI 呼叫都跑完的最後一步，白等一分鐘且無友善提示（僅落到 `main.py:1502` 的泛用 except）。至少要在寫檔前先偵測可寫入並提早提示，或改為寫暫存檔再置換。
