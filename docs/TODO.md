# TODO

> **專案定位（2026-08-03 確立）**：這個程式只做一件事——**把 SEC EDGAR 的原始財務資料抓好**。
> 後續的判讀、分析、報告一律交給外部 skill。任何「幫使用者判斷」的功能都不屬於本專案範圍。

## B. Non-GAAP 改走 skill（下一階段，另開對話處理）

B2. **skill 端由 Claude Code 抽取**，寫回 `nongaap_cache.json`（沿用現有格式）。下游的固定模板、殘差檢查、`Data_Ratios` 一行都不用改，只是把「誰來抽取」從 Gemini 換成 Claude
   - 優點：無 API 額度、無金鑰、模型強很多（gemini-flash 漏抓 ARLO 的無形攤銷與稅務影響，殘差 −3.65M 就是那塊）
   - 限制：不可重現（靠殘差檢查 + 快取緩解）、批次規模要分家跑、GUI 使用者拿不到

B4. **最新法說會資料能不能自動更新**（2026-08-12 CTH 提出）
   - 現況：EDGAR 的 10-Q/10-K 財報資料本來就有時間落差（申報日晚於法說會/財報發布日），沒有即時性
   - 預期方向：自動抓最新 8-K（尤其 Item 2.02 財報發布 8-K）來補即時數字，而不是等 10-Q/10-K 才更新——跟 B1 的 `cli.py press-release` 子指令、B3 的 `press_release_tables.py` 應該是同一條路線，需要研究怎麼接
   - **待研究，未動手**

## C. 與 financial-assistant 體系的銜接（另開對話處理）

C1. `finance-analysis.md` 第二步的作業類型表加一列「SEC 財報抓取（美股）」指向本專案。

C2. `maintain-company-us.md` 的 Phase 1 目前是「開啟公版 Excel → 刷新 Bloomberg → 人工確認歷史數字」，可改為**先用本工具抓 SEC 數字與 Bloomberg 對帳**——兩個獨立來源不一致即為警訊，比單一來源可靠。

C3. `financial-assistant/scripts/` 的 `read_excel.py` / `query_excel.py` 可直接吃 `Data_Financials(Q)` 的固定列位與機器鍵。

C4. ⚠ **衝突要處理**：`finance-analysis.md` 規定「每次更新前先給 CTH 看草稿，確認後才寫入檔案」。本工具是直接寫 Excel 的，接進 financial-assistant 流程時不可直接覆蓋公司資料夾的檔案。

## E. GUI 功能需求（CTH 提出）

E1. ~~**GUI 語言可切換**~~ ✅ **已完成（2026-08-14）**
   - 繁體中文／简体中文／English／日本語四語，範圍含 GUI 與 Excel 輸出的 B 欄
   - 實作見 `docs/superpowers/specs/2026-08-14-i18n-design.md` 與
     `docs/ARCHITECTURE.md`「多語言」章節
   - **未涵蓋（刻意）**：`logs/app.log` 內容、`cli.py` 自己的主控台訊息與
     argparse help、`nongaap_layout.py` / `metric_rules.py`（Data_NonGAAP 的
     版面，該功能停用中）。理由與豁免清單在 `tests/test_i18n.py`

E2. **Data_NonGAAP 版面的 i18n**（承 E1，等 Non-GAAP 功能恢復再做）
   - `nongaap_layout.py` 11 條、`metric_rules.py` 60 條中文字串仍是寫死的
   - 現況：`main.NONGAAP_ENABLED = False`，這張 sheet 根本不產出，沒有
     golden 覆蓋，遷移風險驗不掉
   - 要做的話得連 A 欄機器鍵一起改英文（跟 Data_Ratios 同一套處理），
     不是單純把字串搬進 locale
   - **綁定 TODO B**：Non-GAAP 改走 skill 之後這張表的形態可能整個變，
     現在做很可能白做

## 執行順序建議

| 順位 | 項目 | 需要 API？ | 需要人？ | 說明 |
|---|---|---|---|---|
| 1 | B2 skill 端抽取 | 否 | 是 | 介面是 `cli.py press-release --json` |
| 2 | D8 金融股模板 | 否 | 是 | 51 家調查已有資料；但要不要另開模板是判斷題 |
| 3 | D7 `google-genai` 汰換 | 是 | 否 | 要真呼叫才驗得了，等 B 段定案後再說 |

## D. 待 CTH 決定的已知限制

D0-1. **`Data_Financials(Q)` 永遠沒有 Q4** → TTM 類比率算不出來。Q4 沒有 10-Q，數字在 10-K。
   - 實測：NVDA/AVGO/PLTR 的 ROE／ROA／FCF per Share／淨負債EBITDA **整列全空**
   - 修法：Q4 = 年報 − Q1 − Q2 − Q3（流量項）＋ 資產負債表直接取 10-K。**要 CTH 決定做不做**
   - **研究記錄（2026-08-12）**：`fetcher_gaap.py` 目前只抓 `form="10-Q"`，Q4 標籤結構上就不存在；`ratios.py::_ttm()` 缺任一季就回 `None`（設計如此，不當 0 加總），窗口一跨到 Q4 就整列空。修法可仿照既有 Q2/Q3 YTD 拆算的 pattern，在 `_merge_financials` 後補一欄合成 Q4。**風險：中**——會碰核心合併流程；要處理「只抓季報沒抓年報」時優雅跳過、Non-GAAP/segment 不適用、override 套用順序。附帶好處：算出的 Q4 跟真實 10-K 的差額可以順便當資料品質檢查

D0-2. **多股別公司抓不到期末流通股數**：PLTR／GOOGL／META `company.get_facts()` 裡 `dei:EntityCommonStockSharesOutstanding` **0 筆**（只有 `EntityPublicFloat`），因為 Class A/B/C 是分開標的。TSLA 61 筆、COHR 62 筆、BRK.B 7 筆正常。連帶 BVPS／FCF per Share／流通股數 YoY 空白。`output/_final/META.xlsx` 現在就有這個洞。
   - **研究記錄（2026-08-12）**：`fetcher_gaap.py` 只查單一 XBRL concept，多股別公司這欄位本來就是空的（各股別分開報）。修法要改成按股別分別抓再加總，但程式碼裡已有註解記錄過「連單一股別公司這欄位都不乾淨」的踩坑史（抓到的是財報日後幾週的股數，非期末當天）——多股別可能根本沒有乾淨來源。**風險：中偏大**，搞不好要接受「這幾家就是空」當已知限制，或退而求其次抓 `EntityPublicFloat` 換算（精確度打折）

D0-5. **期間標籤是公式、沒有快取值 → 來源檔關著時跨檔案 `MATCH` 抓不到**（2026-08-08 確認）
   - CTH 已確定用**跨檔案讀取**當日常工作方式（不在 `Data_*` 裡加欄、另開工作檔），所以這條從邊角問題變成主要使用路徑上的坑
   - 成因：第 1、3、4 列是 Excel 公式，openpyxl 不算公式也不寫 `<v>` 快取值。來源檔**開著**時 Excel 會重算（`fullCalcOnLoad = True`）沒問題；**關著**時外部參照只讀得到檔案裡的值，那裡是空的
   - 第 5 列（期末結算日）是靜態文字，不受影響——目前的建議解法就是叫使用者拿第 5 列當 `MATCH` 的 key
   - 可能修法：寫檔時同時寫入 Python 算好的快取值（`fiscal_input` 已有 `fiscal_quarter_of()` 參考實作，本來就是公式的規格）。openpyxl 不支援 formula + value 並存，要另外處理
   - **要 CTH 決定做不做**：不做就是文件講清楚用第 5 列；做了才能直接用 `FY2026Q1` 當 key
   - **研究記錄（2026-08-12）**：兩條路——(a) 存檔後直接改 xlsx 內部 XML 把快取值塞進公式旁邊，**風險中**，動底層 XML 較脆弱；(b) 乾脆改成純值不用公式，**風險小**，但會犧牲「改 Index B4 財年起始月即時連動 1/3/4 列」這個功能（而且 B4 本來就只能在 Excel 裡改，程式端本來就抓不到那個編輯動作去重算，這個「即時連動」的價值本身也可以重新評估）。這項是取捨題不是純技術題

D7. 套件汰換：`google-generativeai` 官方已終止支援（不再更新與修 bug），需改用 `google-genai`。影響 `fetcher_nongaap.py:_call_ai()` 與 `override_engine.py:_llm_call()` 的 google 分支，以及 `requirements.txt`。現行版本仍可運作，非緊急。
   - **研究記錄（2026-08-12）**：實際有**三處**要換，不是兩處——`main.py::_test_ai_worker()`（設定面板的「測試連線」按鈕）也用同一套舊 SDK，漏掉會出現「測試連線過了但實際抓取失敗」的詭異情況。三處都是同樣 3 行 pattern，機械式替換。**風險：小**，但一定要真的呼叫 API 才驗得出來（金鑰/模型名稱格式可能有差），沒法只靠看程式碼確認

D8. 金融股（GS/JPM 等）獨立模板：現行 IS/BS 模板對金融股部分欄位空白，需另建模板。低優先。
   - **研究記錄（2026-08-12）**：程式碼裡目前除了印一行警告訊息，**完全沒有任何金融股特殊處理**——現有科目對照表直接套用，對不上的欄位就是空，不是 bug 是模板天生沒設計給銀行股。真要做要重新研究銀行/券商專屬 US-GAAP 科目（存款、放款、備抵呆帳...），等於另建一整套模板。**風險：大**。這其實是「要不要用這工具抓銀行股」的產品決定，不是技術問題，建議先想清楚要不要再談技術
