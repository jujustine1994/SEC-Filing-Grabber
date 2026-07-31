# TODO

1. 確認 Excel 排版現況：實跑一次輸出，逐 sheet 檢視 `Data_Financials(Q/Y)`、`Data_Financials_NG(Q/Y)`、`Data_Seg_*`、`Index` 的欄寬、凍結窗格、數字格式（÷1M 與 EPS 例外）、section 分隔行、subtotal 粗體是否都正確，記錄實際問題再決定要不要調整 `excel_formatter.py`。
2. 8-K 掃描效率優化（先做，TODO 3 才跑得動）：`_get_earnings_filings()`（`fetcher_nongaap.py:447`）目前把全部歷史 8-K 逐筆 `filing.obj()` 下載解析後才篩年份與 `max_filings`，AAPL 235 份只有 94 份是財報。改成先在申報清單上篩：清單已含 `items` 欄（`2.02,9.01`）與 `reportDate`，可零下載完成「挑財報 → 算季度標籤 → 去重 → 套年份 → 切 max_filings → 扣掉 cache 已有」，剩下的才下載。AAPL 抓 4 季由 235 次下載降到 4 次。邊界：SEC 2004-08 才啟用 `2.02` 編號（更早用 Item 12/5），故加「完整掃描」開關保留現行逐筆解析，供需要 2004 年前資料或懷疑漏抓時使用。
3. 8-K 抽取正確性抽檢：AI 隨機挑一批公司（跨產業、跨市值，含非 12 月財年）跑 Non-GAAP 流程，比對 `Data_NonGAAP` 抽到的指標與 8-K press release 原文，確認項目沒漏抓、沒誤抓、期間標籤（FY/Q）對得上，把失敗案例整理成清單。
4. CLI 工具層（`cli.py`）：讓外部 skill 用指令調用現有 fetcher，不經 GUI。例如 `python cli.py nongaap NVDA --years 2020-2026 --json`、`python cli.py gaap AAPL --xlsx out.xlsx`；輸出支援 JSON（給 skill 讀）與 Excel（給人看）。GUI 與核心函式不動，只加一層薄封裝。等 TODO 1、2 驗收完再做。
5. 套件汰換：`google-generativeai` 官方已終止支援（不再更新與修 bug），需改用 `google-genai`。影響 `fetcher_nongaap.py:_call_ai()` 與 `override_engine.py:_llm_call()` 的 google 分支，以及 `requirements.txt`。現行版本仍可運作，非緊急。
6. 金融股（GS/JPM 等）獨立模板：現行 IS/BS 模板對金融股部分欄位空白，需另建模板。低優先。
