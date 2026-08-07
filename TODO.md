# TODO

> **專案定位（2026-08-03 確立）**：這個程式只做一件事——**把 SEC EDGAR 的原始財務資料抓好**。
> 後續的判讀、分析、報告一律交給外部 skill。任何「幫使用者判斷」的功能都不屬於本專案範圍。

## A. 現行重構（進行中）

A1. **輸出精簡為 5 張 sheet**（原 11 張）：
   - `Data_Q` 季度 GAAP、`Data_Y` 年度 GAAP：表頭 4 列（財季 / 日曆季 / 期末結算日 / 申報日）+ IS 22 + BS 42 + CF 26 固定列（B 欄機器鍵）+ **overflow 移到最底部**
   - `Data_Ratios` 比率（Python 計算，零 AI）
   - `Data_Segments` 長格式
   - `Data_Meta` 資訊 + 品質檢查（併入原 `Index`）
   - 砍掉：`Data_Std`（併入 Q/Y）、`Data_Seg_*` 寬格式 5 張、`Data_Financials_NG`、`Data_EPS_Recon`（從未產生）、`Index`
   - **關鍵**：overflow 原本插在每個 section 之間，害 BS/CF 整段位移（`Cash` 在 28~56 列之間跑）。移到底部後模板列號跨公司固定，`Data_Std` 就不需要獨立存在

A2. ~~**GAAP 抓取移除 AI 依賴**~~ ✅ 已完成：`override_engine.E2_LLM_ENABLED = False`。關在 override_engine 而非呼叫端——即使 GUI 照舊把 `ai_config` 傳進來也不會真的打 API。E1 模糊比對照常運作，找不到就警告（不叫 AI 猜）。實測傳真實 ai_config 抓 COHR，AI 呼叫次數 0。

A3. ~~**Non-GAAP 暫停產出**~~ ✅ 已完成：`main.NONGAAP_ENABLED = False`。兩個 GUI checkbox 停用並改標「暫停中，改由 skill 處理」，抓取路徑也加了守衛。
   - ⚠ 差點犯的錯：一開始只在輸出端過濾掉 `Data_NonGAAP`，但 checkbox 還能勾——會照常呼叫 AI 抓完 6 季**才被丟掉**，等於白燒額度。停用要停在源頭。
   - 相關程式碼（`nongaap_layout` / `metric_rules` / 快取）全部保留，改 `True` 就回來。

## B. Non-GAAP 改走 skill（下一階段，另開對話處理）

B1. **CLI 工具層**（原 TODO 6）：`python cli.py press-release ARLO --years 2025-2026 --tables --json`
   - 輸出每季新聞稿**已解析的表格**（只留含 non-gaap / reconciliation 的），每季約 2~3K 字元
   - 對照：丟原文是 320K 字元，丟已解析表格約 20K —— CLI 負責結構化、skill 負責判讀
   - 另需 `python cli.py build-excel ARLO`：讀快取產 Excel，讓抽取與產表分離

B2. **skill 端由 Claude Code 抽取**，寫回 `nongaap_cache.json`（沿用現有格式）。下游的固定模板、殘差檢查、`Data_Ratios` 一行都不用改，只是把「誰來抽取」從 Gemini 換成 Claude
   - 優點：無 API 額度、無金鑰、模型強很多（gemini-flash 漏抓 ARLO 的無形攤銷與稅務影響，殘差 −3.65M 就是那塊）
   - 限制：不可重現（靠殘差檢查 + 快取緩解）、批次規模要分家跑、GUI 使用者拿不到

B3. **確定性表格解析（中期，可完全免 AI）**：已驗證 `pandas.read_html` 能從 ARLO 新聞稿拆出 15 個表，調節表數字完整。工程量在清理 Workiva 的空白間隔欄與重複欄位。多數美股新聞稿由 Workiva 產生，版面相近，值得做。

## C. 與 financial-assistant 體系的銜接（另開對話處理）

C1. `finance-analysis.md` 第二步的作業類型表加一列「SEC 財報抓取（美股）」指向本專案。

C2. `maintain-company-us.md` 的 Phase 1 目前是「開啟公版 Excel → 刷新 Bloomberg → 人工確認歷史數字」，可改為**先用本工具抓 SEC 數字與 Bloomberg 對帳**——兩個獨立來源不一致即為警訊，比單一來源可靠。

C3. `financial-assistant/scripts/` 的 `read_excel.py` / `query_excel.py` 可直接吃 `Data_Q` 的固定列位與機器鍵（原本每家公司 `Cash` 在 28~56 列之間跑，任何固定參照都會錯）。

C4. ⚠ **衝突要處理**：`finance-analysis.md` 規定「每次更新前先給 CTH 看草稿，確認後才寫入檔案」。本工具是直接寫 Excel 的，接進 financial-assistant 流程時不可直接覆蓋公司資料夾的檔案。

## D. 既有待辦

1. 確認 Excel 排版現況：實跑一次輸出，逐 sheet 檢視 `Data_Financials(Q/Y)`、`Data_Financials_NG(Q/Y)`、`Data_Seg_*`、`Index` 的欄寬、凍結窗格、數字格式（÷1M 與 EPS 例外）、section 分隔行、subtotal 粗體是否都正確，記錄實際問題再決定要不要調整 `excel_formatter.py`。
2. ~~**`Data_NonGAAP` 資料品質修復**~~ ✅ **已完成（2026-08-01 下午，方案 c）**，詳見 CHANGELOG。規則表集中在 `metric_rules.py`，改表後重跑即生效、不必重抓。ARLO 實跑對 8-K 原文逐項核對全中。**殘留待決定事項（都有現成開關，不急）**：
   - ~~年度值補洞~~ ✅ 已改為 `FY_ONLY_HANDLING = "label"`：年度值另成一列加 ` (FY)`，不再填進季欄位。
   - ~~百分比存法~~ ✅ 已改為 Excel 原生比例（`PERCENT_AS_EXCEL_RATIO = True`，0.375 + `0.0%`）。
   - ~~同義名合併~~ ✅ 已拆開：查原文確認 ARLO 2025 年把 `service gross margin` 改名為 `subscriptions and services gross margin`，營收基礎同時變動，不可視為同一條線。
   - **對照表覆蓋率**：目前只依 ARLO / CRM / PANW 三家的實際輸出建表。跑到新公司時若出現沒收錄的中文名，會**原樣顯示中文**（不會丟資料，但那一列不會跟英文季合併）。第 3 項抽檢時順便擴充。
   - **CRM 5 季／PANW 6 季待補抓**（2026-08-02 更新）：ARLO 已用修正後的完整文字重抓完成。CRM 只成功 1 季（`FY2026Q2`，殘差恰為 0，調節橋完整對平）、PANW 0 季，皆因 Gemini `HTTP 429` 每日配額用盡。**失敗的季沒有寫入快取，換一把 key 或隔天直接重跑即會補抓**。
   - ~~考慮只送「前 12K + 後 40K」省 token~~ ❌ **實測後否決（2026-08-03）**。拿 `scripts/survey_nongaap_metrics.py` 快取的 26 份新聞稿統計指標位置：
     - 前 12K 就夠的只有 **2 家**（LRCX、WDAY）——回頭證實原本的截斷 bug 影響 24/26 家，不是 ARLO 特例
     - 「前 12K + 後 40K」會漏掉 **5 家**（AMZN、CRM、NOW、ORCL、QRVO），漏的位置全在 12K–25K 中段：CRM 漏 15,664/24,697、ORCL 漏 14,489/19,987/23,739、QRVO 漏 12,297/14,624/15,474。這幾家的調節表就放在中段
     - 照此裁切等於 19% 的公司會靜默掉資料，與這次修掉的 bug 是同一種錯誤
     - 另：若配額瓶頸是**每日請求數**而非 token 數，砍內文完全無助。要省應該省「重複重抓同一家」
     **結論：維持送完整文字（`PROMPT_TEXT_LIMIT = 200_000`）。**
   - ~~AI 呼叫沒有重試機制~~ ✅ 已加：`AI_MAX_ATTEMPTS = 3`、退避 5s／15s，跑完列出未取得季度並推給 `progress_cb`。次數刻意壓低——Gemini 每日配額型 429 重試必敗，只有每分鐘限流型救得回來。
   - **多把 API key 輪替尚未實作**：配額用盡時目前要手動去進階設定換 key。若常撞到，可考慮讓 config 收多把 key 自動輪替。
3. ~~**期末流通股數取不到**~~ ✅ 已解決（2026-08-02）：改走封面頁 `dei:EntityCommonStockSharesOutstanding`（`Company.get_facts()`），歷史序列完整。**殘留注意**：該 fact 的日期是封面頁「最近可行日期」，比財季結束晚幾週（ARLO FY2025Q1 財季結束 2025-03-30，股數是 2025-05-02 的數字），因此 `BVPS` 是「期末權益 ÷ 幾週後的股數」，量級無虞但不是同一天。若要更精確需另找 BS parenthetical 的 tag，多數公司沒有。

4. 8-K 抽取正確性抽檢：AI 隨機挑一批公司（跨產業、跨市值，含非 12 月財年）跑 Non-GAAP 流程，比對 `Data_NonGAAP` 抽到的指標與 8-K press release 原文，確認項目沒漏抓、沒誤抓、期間標籤（FY/Q）對得上，把失敗案例整理成清單。**另需特別檢查**：`_period_to_quarter_label()` 是用 Item 2.02 8-K 的 `period_of_report`（EDGAR 上這欄放的是**發布/事件日**，不是財報所屬的財期結束日）去分季。實測驗證：INTC 的 `20260723` 被標成 `FY2026Q3`，但那份新聞稿報的其實是 FY2026 Q2（~6/27 結束）；COST 的 `20260528` 被標成 `FY2026Q2`，實際報的是 FY2026 Q3（5 月中結束）；WDC 的 `20260430` 被標成 `FY2026Q2`。也就是說 `Data_NonGAAP` sheet 上**每一欄的標籤幾乎都比數字實際所屬的財季晚了一季**，是系統性的 off-by-one，不只是「日曆季 vs 公司財季」的差異。這是長期存在的行為（非本次分支引入），但範圍不小，抽檢時要一併確認影響程度並評估修法。**另外，同一根因還會造成漏抓**：WDC 曾在同一個日曆季（2025 Q1）內發布兩份 Item 2.02 8-K（`20250110` 與 `20250129`），兩者都被標成 `FY2025Q1`，dedupe 邏輯「保留最舊那份」會直接把 1/29 那份財報悄悄丟掉，這點也要一併排查。
5. `max_filings` 不再是硬上限：`_list_earnings_filings()` 先套用切片，缺季回補（`_recover_missing_quarters()`）在補齊後不會重新裁切，故要求 8 季、若保留區間有 2 個缺口，實際可能下載到 10 份。需評估是否要在回補後補一次裁切，或至少在文件中說明此行為。
6. CLI 工具層（`cli.py`）：讓外部 skill 用指令調用現有 fetcher，不經 GUI。例如 `python cli.py nongaap NVDA --years 2020-2026 --json`、`python cli.py gaap AAPL --xlsx out.xlsx`；輸出支援 JSON（給 skill 讀）與 Excel（給人看）。GUI 與核心函式不動，只加一層薄封裝。等「Excel 排版確認」與「8-K 抽取正確性抽檢」驗收完再做。
7. 套件汰換：`google-generativeai` 官方已終止支援（不再更新與修 bug），需改用 `google-genai`。影響 `fetcher_nongaap.py:_call_ai()` 與 `override_engine.py:_llm_call()` 的 google 分支，以及 `requirements.txt`。現行版本仍可運作，非緊急。
8. 金融股（GS/JPM 等）獨立模板：現行 IS/BS 模板對金融股部分欄位空白，需另建模板。低優先。
9. 重複抓同一 ticker 的檔案行為需補防護：`excel_writer.write_statements()` 開啟既有檔後刪除所有 `Data_*` sheet 重寫（`My_*` 等自訂 sheet 保留），無備份無版本號。兩個風險：(a) 第二次抓的年份範圍較窄時，`Data_*` 是整批替換而非合併，舊季度直接消失（GAAP 無快取需重抓）；(b) 該 xlsx 正被 Excel 開啟時 Windows 鎖檔，`wb.save()` 拋 `PermissionError`，但這發生在全部抓取與 AI 呼叫都跑完的最後一步，白等一分鐘且無友善提示（僅落到 `main.py:1502` 的泛用 except）。至少要在寫檔前先偵測可寫入並提早提示，或改為寫暫存檔再置換。
