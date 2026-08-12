# TODO

> **專案定位（2026-08-03 確立）**：這個程式只做一件事——**把 SEC EDGAR 的原始財務資料抓好**。
> 後續的判讀、分析、報告一律交給外部 skill。任何「幫使用者判斷」的功能都不屬於本專案範圍。

## E. 專案目錄結構整理（2026-08-12 CTH 提出，目前忙沒空處理）

**目前狀況（跟 `windows-tool.md`「專案目錄結構」規則對不上的地方）：**

規則要求根目錄只放 `啟動器.bat` / `launcher.ps1` / `README.md` / `.gitignore` /
`requirements.txt`，Python 原始碼進 `src/`，所有 MD 文件（除 README）進 `docs/`。
現況：

1. **17 個 `.py` 原始檔全部散在根目錄**，完全沒有 `src/`：
   `main.py`、`cli.py`、`config.py`、`conftest.py`、`errsafe.py`、
   `excel_formatter.py`、`excel_writer.py`、`fetcher_gaap.py`、
   `fetcher_nongaap.py`、`fiscal_input.py`、`metric_rules.py`、
   `nongaap_layout.py`、`output_tables.py`、`override_engine.py`、
   `press_release_tables.py`、`ratios.py`、`segments.py`、`zh_labels.py`
2. **`ARCHITECTURE.md`、`CHANGELOG.md`、`PITFALLS.md`、`TODO.md`、
   `docs_statement_template_proposal.md` 都在根目錄**，但專案其實已經有
   `docs/` 資料夾（目前只放 `8k-period-off-by-one.md`），文件分裂成兩處
3. 根目錄有一個 **`20260812 sec工具.rar`**（466KB）——看起來像手動備份/匯出檔，
   不屬於規則列的任何一類，內容還沒確認過（CTH 忙，還沒回覆要留還是搬走）
4. `company_cache.json`、`config.example.json` 在根目錄——規則沒明講這類
   config/資料檔歸屬，概念上比較接近 `data/`，優先度低

`scripts/`、`tests/`、`logs/`、`docs/`、`venv/` 資料夾本身位置沒問題。

**我的判斷：**

- 最高風險的是搬 `.py` 進 `src/`——17 個檔案互相 import（`main.py` 底下
  `from fetcher_gaap import ...` 這類），加上 `cli.py`／`main.py`／
  `launcher.ps1` 的呼叫路徑、`tests/` 底下每個測試檔的 import、`conftest.py`
  的路徑假設，全部要跟著改，一次性搬完再測風險太高，**必須照
  `templates\doc-init-protocol.md` 的規定分批搬、每搬一批就跑一次
  `pytest -q` 確認全綠再搬下一批**
- MD 文件搬 `docs/` 風險低很多（只有 `README.md` 開頭的相對連結、
  `project-rules` 讀取路徑需要檢查有沒有寫死根目錄），可以先做，
  跟 `.py` 搬遷分開兩次處理
- `.rar` 最單純、跟程式碼結構完全無關，不影響任何 import，隨時可以先問
  CTH 要留要丟要搬，不用等 `.py`/MD 那兩塊

**建議的詳細步驟（照風險排序，每步做完才做下一步）：**

1. **確認 `.rar`**：問 CTH 這份手動備份是要保留（搬到專案外或 `.gitignore` 掉）
   還是可以刪除；這步不影響任何程式碼，最先做
2. **搬 MD 文件進 `docs/`**：
   a. `git mv ARCHITECTURE.md CHANGELOG.md PITFALLS.md TODO.md docs/`
   b. `git mv docs_statement_template_proposal.md docs/`（先確認這份是否還有用，
      過時的話直接問 CTH 要不要一併丟棄而非搬移）
   c. 檢查 `README.md` 開頭「規則檔」欄位、`doc-init-protocol.md` 提到的
      「若 `docs/` 下找不到就改讀根目錄」相容邏輯，確認 AI 之後讀得到新位置
   d. 檢查有沒有程式碼或文件內文寫死 `"./ARCHITECTURE.md"` 這類根目錄相對路徑
      （`grep -rn "ARCHITECTURE.md\|CHANGELOG.md\|PITFALLS.md\|TODO.md"` 排除
      `docs/` 本身跟 `.git/`）
   e. 跑一次 `pytest -q` 確認測試沒有讀這些檔案路徑
3. **搬 `.py` 進 `src/`（風險最高，務必分批）**：
   a. 先列出完整的檔案間 import 關係圖（`grep -n "^from \|^import "` 逐檔掃）
   b. 從**沒有被其他模組 import、只被 `main.py`/`cli.py` 當入口呼叫**的檔案開始搬
      （風險最低），每搬一個就更新所有引用它的 `import` 路徑，跑一次
      `pytest -q` 全綠才搬下一個
   c. 核心、被最多檔案 import 的（`fetcher_gaap.py`、`config.py`）留到最後搬，
      牽動面最大
   d. 全部搬完後，更新 `launcher.ps1` 呼叫 `main.py` 的路徑（若 `main.py` 也搬
      進 `src/`）、`啟動器.bat` 若有寫死路徑也要一併檢查
   e. 最後跑一次 `啟動器.bat` 實際雙擊驗證（不能只看 pytest 綠燈就宣稱完成，
      這是啟動器類專案的驗收慣例）
4. **（低優先，可選）** `company_cache.json`／`config.example.json` 要不要收進
   `data/`——這塊沒有明確規則規定，可以最後再問 CTH 要不要做

**要 CTH 有空時決定**：要不要做、要不要一次做完三步還是分次處理。

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

D0-2. **多股別公司抓不到期末流通股數**：PLTR／GOOGL／META `company.get_facts()` 裡 `dei:EntityCommonStockSharesOutstanding` **0 筆**（只有 `EntityPublicFloat`），因為 Class A/B/C 是分開標的。TSLA 61 筆、COHR 62 筆、BRK.B 7 筆正常。連帶 BVPS／FCF per Share／流通股數 YoY 空白。`output/_final/META.xlsx` 現在就有這個洞。

D0-5. **期間標籤是公式、沒有快取值 → 來源檔關著時跨檔案 `MATCH` 抓不到**（2026-08-08 確認）
   - CTH 已確定用**跨檔案讀取**當日常工作方式（不在 `Data_*` 裡加欄、另開工作檔），所以這條從邊角問題變成主要使用路徑上的坑
   - 成因：第 1、3、4 列是 Excel 公式，openpyxl 不算公式也不寫 `<v>` 快取值。來源檔**開著**時 Excel 會重算（`fullCalcOnLoad = True`）沒問題；**關著**時外部參照只讀得到檔案裡的值，那裡是空的
   - 第 5 列（期末結算日）是靜態文字，不受影響——目前的建議解法就是叫使用者拿第 5 列當 `MATCH` 的 key
   - 可能修法：寫檔時同時寫入 Python 算好的快取值（`fiscal_input` 已有 `fiscal_quarter_of()` 參考實作，本來就是公式的規格）。openpyxl 不支援 formula + value 並存，要另外處理
   - **要 CTH 決定做不做**：不做就是文件講清楚用第 5 列；做了才能直接用 `FY2026Q1` 當 key

D7. 套件汰換：`google-generativeai` 官方已終止支援（不再更新與修 bug），需改用 `google-genai`。影響 `fetcher_nongaap.py:_call_ai()` 與 `override_engine.py:_llm_call()` 的 google 分支，以及 `requirements.txt`。現行版本仍可運作，非緊急。

D8. 金融股（GS/JPM 等）獨立模板：現行 IS/BS 模板對金融股部分欄位空白，需另建模板。低優先。
