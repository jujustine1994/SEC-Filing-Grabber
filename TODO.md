# TODO

1. 確認 Excel 排版現況：實跑一次輸出，逐 sheet 檢視 `Data_Financials(Q/Y)`、`Data_Financials_NG(Q/Y)`、`Data_Seg_*`、`Index` 的欄寬、凍結窗格、數字格式（÷1M 與 EPS 例外）、section 分隔行、subtotal 粗體是否都正確，記錄實際問題再決定要不要調整 `excel_formatter.py`。
2. 8-K 抽取正確性抽檢：AI 隨機挑一批公司（跨產業、跨市值，含非 12 月財年）跑 Non-GAAP 流程，比對 `Data_NonGAAP` 抽到的指標與 8-K press release 原文，確認項目沒漏抓、沒誤抓、期間標籤（FY/Q）對得上，把失敗案例整理成清單。**另需特別檢查**：`_period_to_quarter_label()` 是用 Item 2.02 8-K 的 `period_of_report`（EDGAR 上這欄放的是**發布/事件日**，不是財報所屬的財期結束日）去分季。實測驗證：INTC 的 `20260723` 被標成 `FY2026Q3`，但那份新聞稿報的其實是 FY2026 Q2（~6/27 結束）；COST 的 `20260528` 被標成 `FY2026Q2`，實際報的是 FY2026 Q3（5 月中結束）；WDC 的 `20260430` 被標成 `FY2026Q2`。也就是說 `Data_NonGAAP` sheet 上**每一欄的標籤幾乎都比數字實際所屬的財季晚了一季**，是系統性的 off-by-one，不只是「日曆季 vs 公司財季」的差異。這是長期存在的行為（非本次分支引入），但範圍不小，抽檢時要一併確認影響程度並評估修法。**另外，同一根因還會造成漏抓**：WDC 曾在同一個日曆季（2025 Q1）內發布兩份 Item 2.02 8-K（`20250110` 與 `20250129`），兩者都被標成 `FY2025Q1`，dedupe 邏輯「保留最舊那份」會直接把 1/29 那份財報悄悄丟掉，這點也要一併排查。
3. `max_filings` 不再是硬上限：`_list_earnings_filings()` 先套用切片，缺季回補（`_recover_missing_quarters()`）在補齊後不會重新裁切，故要求 8 季、若保留區間有 2 個缺口，實際可能下載到 10 份。需評估是否要在回補後補一次裁切，或至少在文件中說明此行為。
4. CLI 工具層（`cli.py`）：讓外部 skill 用指令調用現有 fetcher，不經 GUI。例如 `python cli.py nongaap NVDA --years 2020-2026 --json`、`python cli.py gaap AAPL --xlsx out.xlsx`；輸出支援 JSON（給 skill 讀）與 Excel（給人看）。GUI 與核心函式不動，只加一層薄封裝。等「Excel 排版確認」與「8-K 抽取正確性抽檢」驗收完再做。
5. 套件汰換：`google-generativeai` 官方已終止支援（不再更新與修 bug），需改用 `google-genai`。影響 `fetcher_nongaap.py:_call_ai()` 與 `override_engine.py:_llm_call()` 的 google 分支，以及 `requirements.txt`。現行版本仍可運作，非緊急。
6. 金融股（GS/JPM 等）獨立模板：現行 IS/BS 模板對金融股部分欄位空白，需另建模板。低優先。
