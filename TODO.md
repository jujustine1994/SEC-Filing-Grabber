# TODO

> 🛑 **擋路中：D1 Excel 排版驗收（CTH 親自做，做完之前不要動別的項目）**
>
> **AI 注意**：新對話開始時若這個區塊還在，**先提醒 CTH 這件事還沒做**，不要直接開始做其他 TODO。
> 驗收完成後把這個區塊整個刪掉，並在下面的執行順序表把 D1 劃掉。
>
> **CTH 要做的事**：開 `output/_final/` 的七份檔案（AAPL / NVDA / META / AVGO / MSFT / COHR / PLTR，
> 均為 2026-08-08 新字型重產），用眼睛逐項看：
>
> - [x] 欄寬 — **CTH 2026-08-08 驗收通過**
> - [x] 凍結窗格 — **維持現狀 D3**（CTH 決定，D0-4 可結案不做）
> - [x] 三表底色與 5 列間隔 — 通過
> - [x] 數字格式（÷1M、`0.0%`、每股兩位、`(x)`/`(days)`）— 通過
> - [x] 新字型與 Index 新字級 — 通過
> - [x] ~~Index 提醒列（A5）會切到文字~~ **已修（2026-08-08）**：列高改成依文字長度與欄寬推算
>       （`fiscal_input._wrapped_row_height()`），AAPL 實測 94.5。**待 CTH 複驗**
> - [ ] `Index!B4` 改一個數字，看第 1、3、4 列有沒有跟著變 — **待 CTH 試**
>       （做法：開 `Data_Financials(Q)` 分頁看第 1 列。AAPL 的 B4 從 `10` 改成 `1`，
>       D1 應該從 `FY2023Q1` 變成 `FY2022Q4`。**試完記得改回 10**）
>
> **還沒重產的四份**：AVGO / MSFT / COHR / PLTR 在 2026-08-08 修列高後重跑時被鎖檔
> （CTH 驗收中開著）。AAPL / NVDA / META 已是最新。下次要重跑這四家。
>
> 為什麼卡在這裡：排版沒驗收就繼續疊功能，等於在未確認的基礎上加東西；而且這是
> 執行順序表的第 1 順位。**此項不可由 AI 代勞。**

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

B1. ~~**CLI 工具層**（原 TODO 6）~~ ✅ **已完成（2026-08-07）**：`cli.py` 兩個子指令
   - `cli.py gaap AAPL --years 2023-2026 --xlsx out.xlsx`：實測 23.7 秒，與 GUI 產的 `output/_final/AAPL.xlsx` 逐格比對 4,089 格、差異 13 格全是「抓取日期」
   - `cli.py press-release ARLO --years 2025-2026 --tables --json`：實測 5.9 秒抓 7 季，每季 2.8~6.6K 字元（原文 450K）
   - 薄封裝，核心函式沒動。唯一重構是把 `main._append_ratio_table` 抽到 `output_tables.py`（`import main` 會拉進 tkinter），GUI 保留舊名別名
   - 網路呼叫集中在 3 個函式，24 個測試全離線
   - **未做**：`cli.py build-excel ARLO`（讀快取產 Excel）。等 B2 的 skill 真的開始寫快取再做，現在做等於猜介面

B2. **skill 端由 Claude Code 抽取**，寫回 `nongaap_cache.json`（沿用現有格式）。下游的固定模板、殘差檢查、`Data_Ratios` 一行都不用改，只是把「誰來抽取」從 Gemini 換成 Claude
   - 優點：無 API 額度、無金鑰、模型強很多（gemini-flash 漏抓 ARLO 的無形攤銷與稅務影響，殘差 −3.65M 就是那塊）
   - 限制：不可重現（靠殘差檢查 + 快取緩解）、批次規模要分家跑、GUI 使用者拿不到

B3. ~~**確定性表格解析（中期，可完全免 AI）**~~ ✅ **已完成（2026-08-07）**：`press_release_tables.py`
   - 實測 12 家最新 Item 2.02 8-K：全部零 ragged row、零異常寬表；ARLO 450K→4.4K、NVDA 275K→2.2K、MSFT 901K→2.1K
   - AAPL / COST 篩出 0 張是**正確**的——這兩家本來就不報 Non-GAAP
   - 關鍵地雷：間隔欄要用「所有**資料列**都空」判斷，不能用「整欄都空」。Workiva 的表頭是 colspan 展開的，`Three Months Ended` 會把 15 欄（含間隔欄）全部填滿
   - **殘留**：NVDA 的重點摘要表沒有間隔欄，多期數字會併成一格（`"$81,615 $68,127 $44,062"`）。刻意保留而不取其一——資料異常要看得見。真正的調節表不受影響

## C. 與 financial-assistant 體系的銜接（另開對話處理）

C1. `finance-analysis.md` 第二步的作業類型表加一列「SEC 財報抓取（美股）」指向本專案。

C2. `maintain-company-us.md` 的 Phase 1 目前是「開啟公版 Excel → 刷新 Bloomberg → 人工確認歷史數字」，可改為**先用本工具抓 SEC 數字與 Bloomberg 對帳**——兩個獨立來源不一致即為警訊，比單一來源可靠。

C3. `financial-assistant/scripts/` 的 `read_excel.py` / `query_excel.py` 可直接吃 `Data_Q` 的固定列位與機器鍵（原本每家公司 `Cash` 在 28~56 列之間跑，任何固定參照都會錯）。

C4. ⚠ **衝突要處理**：`finance-analysis.md` 規定「每次更新前先給 CTH 看草稿，確認後才寫入檔案」。本工具是直接寫 Excel 的，接進 financial-assistant 流程時不可直接覆蓋公司資料夾的檔案。

## 執行順序建議（2026-08-03 排定）

| 順位 | 項目 | 需要 API？ | 需要人？ | 說明 |
|---|---|---|---|---|
| ~~1~~ | ~~**B1 CLI 工具層**~~ | 否 | 否 | ✅ 2026-08-07 完成 |
| ~~2~~ | ~~**D4 前半：查清 8-K off-by-one**~~ | 否 | 否 | ✅ 2026-08-07 完成，見 `docs/8k-period-off-by-one.md` |
| ~~3~~ | ~~**B3 確定性表格解析**~~ | 否 | 否 | ✅ 2026-08-07 完成 |
| **1** | **D1 Excel 排版驗收** | 否 | **是** | 要開 Excel 用眼睛看，AI 代勞不了。**現在是唯一擋路的** |
| 2 | **D4 後半：實際修 off-by-one** | 否 | **是** | 影響已量化（只有 13% 標對），三個修法選項在報告末尾。會動到快取 key 需重抓 |
| 3 | **dedupe 「保留最舊」要改** | 否 | 否 | 與 D4 同根因但可獨立修：實測 2 次撞標籤**兩次都保留錯的那份**（WDC 整季消失、QRVO 拿到 preliminary） |
| 4 | B2 skill 端抽取 | 否 | 是 | B1/B3 已就緒，介面是 `cli.py press-release --json` |
| 5 | D8 金融股模板 | 否 | 是 | 51 家調查已有資料；但要不要另開模板是判斷題 |
| 6 | D7 `google-genai` 汰換 | 是 | 否 | 要真呼叫才驗得了，等 B 段定案後再說 |

## D0. 2026-08-08 四家實測抓到的（NVDA/AAPL/PLTR/AVGO）

D0-1. **`Data_Financials(Q)` 永遠沒有 Q4** → TTM 類比率算不出來。Q4 沒有 10-Q，數字在 10-K。
   - 實測：NVDA/AVGO/PLTR 的 ROE／ROA／FCF per Share／淨負債EBITDA **整列全空**
   - AAPL 原本有兩欄有值，但那是**錯的**——標籤 off-by-one 造成四欄「看似連續」，TTM 實際上跳過 Sep 季又重複 Dec 季。標籤修好後正確地變成空白
   - 修法：Q4 = 年報 − Q1 − Q2 − Q3（流量項）＋ 資產負債表直接取 10-K。**要 CTH 決定做不做**

D0-2. **多股別公司抓不到期末流通股數**：PLTR／GOOGL／META `company.get_facts()` 裡 `dei:EntityCommonStockSharesOutstanding` **0 筆**（只有 `EntityPublicFloat`），因為 Class A/B/C 是分開標的。TSLA 61 筆、COHR 62 筆、BRK.B 7 筆正常。連帶 BVPS／FCF per Share／流通股數 YoY 空白。`output/_final/META.xlsx` 現在就有這個洞。

D0-3. ~~**AAPL 財季標籤 off-by-one**~~ ✅ 已於 2026-08-08 隨「財年起始月公式化」修掉。原本 `2023-04-01` 被標成 `FY2023Q3`（實為 FY2023 Q2，營收 94,836 = Apple 自報的 Q2 FY23）。389 格比對只有 AAPL 這 2 欄有差，改後全對。

D0-5. **期間標籤是公式、沒有快取值 → 來源檔關著時跨檔案 `MATCH` 抓不到**（2026-08-08 確認）
   - CTH 已確定用**跨檔案讀取**當日常工作方式（不在 `Data_*` 裡加欄、另開工作檔），所以這條從邊角問題變成主要使用路徑上的坑
   - 成因：第 1、3、4 列是 Excel 公式，openpyxl 不算公式也不寫 `<v>` 快取值。來源檔**開著**時 Excel 會重算（`fullCalcOnLoad = True`）沒問題；**關著**時外部參照只讀得到檔案裡的值，那裡是空的
   - 第 5 列（期末結算日）是靜態文字，不受影響——目前的建議解法就是叫使用者拿第 5 列當 `MATCH` 的 key
   - 可能修法：寫檔時同時寫入 Python 算好的快取值（`fiscal_input` 已有 `fiscal_quarter_of()` 參考實作，本來就是公式的規格）。openpyxl 不支援 formula + value 並存，要另外處理
   - **要 CTH 決定做不做**：不做就是文件講清楚用第 5 列；做了才能直接用 `FY2026Q1` 當 key

D0-4. **凍結窗格在 D3，但表頭有 5 列**：往下捲會看不到第 3~5 列（財季／日曆季／期末結算日）。建議改 D6，但屬使用習慣，等 CTH 決定。

## D0b. 排版調整（2026-08-08 CTH 指定）

D0b-1. ~~**`Index` 字體大小要調**~~ ✅ **已完成（2026-08-08）**：CTH 選「整體放大一級」。
   A1 抬頭 14→16、A2 資訊列 9→10、表格（標題＋內容）10→11、B4 輸入格 11→12、財年提醒 9→10。
   字級常數集中在 `excel_formatter`（`INDEX_TITLE_SIZE` / `INDEX_META_SIZE` /
   `INDEX_TABLE_SIZE` / `INDEX_INPUT_SIZE` / `INDEX_NOTE_SIZE`），`fiscal_input` 直接 import。
   提醒那列的列高一併 28→32——字級變大而列高不動，wrap 的尾巴會被切掉。

D0b-2. ~~**整份文件內文字體設為「微軟正黑體」**~~ ✅ **已完成（2026-08-08）**：
   `excel_formatter.FONT_NAME = "微軟正黑體"` + `_font(**kw)` 工廠，兩個模組所有 `Font()` 統一走它。
   - **數字欄用 `NUMBER_FONT_NAME = "Consolas"`（等寬）**：CTH 一開始選「全部含數字欄」，
     看過實際輸出後改成數字等寬（2026-08-08 同日）。判斷依據是**儲存格的值是不是數字**
     （`_is_number()`），不是欄號——`Data_Meta` / `Data_Segments` 的 D 欄起是文字，用欄號會誤傷
   - 順手修掉的洞：Index 表格 B~D 欄（含中文說明）原本**根本沒設 font**，會掉回新細明體——正是這項要修的症狀
   - 測試用掃描式而非逐格斷言（`test_every_font_on_*_uses_the_configured_family`），漏掉任何一處就會紅
   - 合併範圍的填充格（`A1:E1` 的 B1~E1）留在 openpyxl 預設 Calibri，不影響顯示（Excel 只看左上角那格）

## D. 既有待辦

1. **人工驗收 Excel 排版**（只剩這件需要人眼）：`output/_final/` 有 AAPL / NVDA / META / AVGO / MSFT / COHR 六份。逐項看欄寬、凍結窗格、三表底色與 5 列間隔、數字格式（÷1M、百分比 0.0%、每股兩位小數）、中文說明欄。**此項不可由 AI 代勞**，要開 Excel 用眼睛看。
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

4. **8-K 季度標籤 off-by-one**：~~前半（查清影響範圍）~~ ✅ **2026-08-07 完成，報告見 `docs/8k-period-off-by-one.md`**。
   - **結論比原本記的嚴重**：不是「晚一季」，是**偏 −3 到 +1 季，偏多少由財年結束月決定**。16 家 128 份實測，119 份成功比對，**只有 16 份（13%）標對**（AAPL、AVGO 兩家，是兩層誤差剛好抵消而碰巧對上）
   - 根因兩層：`period_of_report` 是發布日不是期末日；算的是日曆季但 `Data_Q` 用公司財季
   - **dedupe「保留最舊」實測 2 次撞標籤、兩次都保留錯的那份**：WDC `FY2025Q1` 保留了沒有新聞稿附件的那份、丟掉真正的 Q2 財報（整季消失）；QRVO `FY2025Q4` 保留了「Preliminary」那份、丟掉正式財報。這條規則要跟標籤分開檢討
   - **後半（實際修）未做**，三個選項（A 改用期末日 / B 只加期末日欄 / C 交給 skill 判定）與各自代價寫在報告末尾，要 CTH 決定
   - 目前是**潛伏**的：`NONGAAP_ENABLED = False`，`Data_NonGAAP` 沒在產。唯一對外吐季度標籤的是 `cli.py press-release`，該路徑每一季都帶 `label_warning`
   - 仍待做：跑 Non-GAAP 流程比對「指標項目」有沒有漏抓誤抓（這部分要等 B2 skill 方案定案）
5. `max_filings` 不再是硬上限：`_list_earnings_filings()` 先套用切片，缺季回補（`_recover_missing_quarters()`）在補齊後不會重新裁切，故要求 8 季、若保留區間有 2 個缺口，實際可能下載到 10 份。需評估是否要在回補後補一次裁切，或至少在文件中說明此行為。
6. ~~CLI 工具層（`cli.py`）~~ ✅ 已完成（2026-08-07），見 B1。`nongaap` 子指令沒做——Non-GAAP 停用中，改由 `press-release` 出結構化表格給 skill。
7. 套件汰換：`google-generativeai` 官方已終止支援（不再更新與修 bug），需改用 `google-genai`。影響 `fetcher_nongaap.py:_call_ai()` 與 `override_engine.py:_llm_call()` 的 google 分支，以及 `requirements.txt`。現行版本仍可運作，非緊急。
8. 金融股（GS/JPM 等）獨立模板：現行 IS/BS 模板對金融股部分欄位空白，需另建模板。低優先。
9. ~~重複抓同一 ticker 的檔案防護~~ ✅ 已完成（commit `d9c684c`）：抓取前 `check_output_writable()` 偵測鎖檔、寫暫存檔再 `os.replace()`、覆蓋前留一份滾動 `.bak.xlsx`。
