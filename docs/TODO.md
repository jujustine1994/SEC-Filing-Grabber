# TODO

> **專案定位（2026-08-03 確立）**：這個程式只做一件事——**把 SEC EDGAR 的原始財務資料抓好**。
> 後續的判讀、分析、報告一律交給外部 skill。任何「幫使用者判斷」的功能都不屬於本專案範圍。

> **本檔維護規則（2026-08-22 CTH 指示）**：**做完的條目直接刪掉**，內容搬進
> `docs/CHANGELOG.md`。這份只留「還沒做」與「還沒決定」的事，不累積已完成清單。
> 想查歷史看 CHANGELOG 或 git log。

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

> 已完成項目一律搬去 `docs/CHANGELOG.md`，這裡不留殘骸。

E2. **Data_NonGAAP 版面的 i18n**（承 E1，等 Non-GAAP 功能恢復再做）
   - `nongaap_layout.py` 11 條、`metric_rules.py` 60 條中文字串仍是寫死的
   - 現況：`main.NONGAAP_ENABLED = False`，這張 sheet 根本不產出，沒有
     golden 覆蓋，遷移風險驗不掉
   - 要做的話得連 A 欄機器鍵一起改英文（跟 Data_Ratios 同一套處理），
     不是單純把字串搬進 locale
   - **綁定 TODO B**：Non-GAAP 改走 skill 之後這張表的形態可能整個變，
     現在做很可能白做

E3. **「查可用期間」按鈕要分兩階段改名**（2026-08-17 CTH 提出，2026-08-18 CTH 確認實際行為，先不做）
   - 現況：不管第幾次按，按鈕都叫 `🔍 查可用期間`（2026-08-17 才從 `▶` 改成
     這個名字，見 `docs/CHANGELOG.md`）
   - **2026-08-18 CTH 實測回報的實際行為**：第一次按下去只顯示股票代號（公司
     名稱），要按**第二次**才真的查到可用期間。跟按鈕名稱字面上的「查可用期間」
     對不上，容易讓人以為按一次就查完、其實還要再按一次
   - CTH 要的：名稱要照實際分兩階段的行為命名，**第一次按 = `查詢公司代號`**
     （確認打的 ticker 存不存在、是哪家公司），**第二次按 = `查詢最新資料上傳
     時間`**（查 SEC 上最新一筆申報是什麼時候）。兩次做的事情不同，名字就該不同
   - **要先查證再動手**：目前程式碼讀起來 `_run_preview_scan`（`main.py`）一次
     點擊裡會同時發動名稱查詢（`_confirm_company`）跟期間掃描
     （`_preview_scan_worker`）兩個背景執行緒，不是刻意設計成分兩次點擊——CTH
     觀察到的「第一次只顯示名稱、第二次才查到期間」有可能是真的分兩階段，也
     可能是掃描結果回填有時間差造成的視覺錯覺（例如名稱查詢走本機快取幾乎
     瞬間回來，期間掃描要打 EDGAR 5~15 秒，兩者本來就不同步完成）。動手前要
     先重現確認是哪一種，兩種原因的修法不同（一種是改名對應真實分階段行為，
     一種是把「還在查期間」的等待狀態講清楚，不用真的分兩次點）
   - **2026-08-18 CTH 已重現**：只有**第一次開啟程式**才會出現「第一次按鈕只
     顯示代號、要按第二次才查到期間」，**第二次以後（同一次執行期間內再查
     其他 ticker，或關掉重開？CTH 目前只確認「第二次以後不會」，還沒細分是
     哪一種「第二次」)** 就不會再發生。原因待查，先記錄現象。這個「只有首次
     發生」的線索指向啟動時的某種延遲初始化（例如 `edgar` 套件、`identity`
     設定、或本機快取第一次建立時的額外開銷），不是每次點擊都有的穩定行為
     ——跟原本猜測的「兩個背景執行緒本來就不同步完成」方向不同，要重新查
   - 要釐清的：如果真的要分兩階段（而不是只有首次啟動的延遲問題），「第二次」
     的定義是什麼？同一個 ticker 內按第二次，還是換了 ticker 就重新算第一次？
     換 ticker 後按鈕文字要不要退回「查詢公司代號」？動手前先問 CTH
   - 注意：按鈕文字是介面文字，四個 locale（`zh_tw`／`zh_cn`／`en`／`ja`）都要加，
     不可只改繁中（見 `docs/ARCHITECTURE.md` 多語言章節）
   - **2026-08-18 查證結果：無法在這個環境重現，改成先補 log 診斷**。三個角度
     都測過，都沒重現「第一次只顯示名稱、掃不到期間」：
     1. 直接量 `fetcher_gaap.preview_sheets()` 的耗時——連 `~/.edgar`／
        `~/.edgar_cache` 都清空模擬全新安裝，耗時還是 ~1.5 秒，edgartools
        沒有明顯的「第一次特別慢」現象（這個環境測的，不是 CTH 那台機器）
     2. 直接建 `SECFetcherApp` 呼叫 `_run_preview_scan()` 模擬點擊（含真的打
        EDGAR），**第一次點擊 sheet 面板就正常顯示、標題列也正確帶出最新期間**，
        程式邏輯本身沒有「兩次點擊才生效」的 bug
     3. 讀過 `_run_preview_scan`／`_preview_scan_worker`／`_poll_queue` 的
        `preview_scan_done` 處理，沒有找到任何「第一次執行會被跳過或吃掉結果」
        的路徑
     - **結論**：根因很可能是 CTH 那台 Windows 機器上第一次對外連線特別慢
       （防火牆彈窗、Proxy 自動偵測、防毒軟體攔截新程式的第一次連線之類），
       這個環境的網路路徑跟 Windows 桌機不同，重現不了，也沒辦法排除或確認
     - **已做的補強**：`_preview_scan_worker`（`main.py`）原本完全沒寫 log，
       出事只能靠使用者口述秒數。現在起訖都記一行到 `logs/app.log`
       （`查可用期間 {ticker}` 開始 + `{ticker} 查可用期間完成/失敗，耗時 X 秒`
       結果，失敗時完整例外內容也會寫進去，不像 GUI 上只顯示例外類型名）
     - **下次重現時**：麻煩 CTH 開新視窗後**第一次**點「查可用期間」就馬上關掉
       程式，把 `logs/app.log` 最後幾行貼出來——如果耗時真的異常長（例如
       10+ 秒），就證實是連線慢，可以考慮把 hint 文字從「約需 10 秒」改成
       更保守的說法；如果耗時正常但 GUI 沒反應，那就是還沒抓到的另一種 bug，
       要再回頭查

E5. **查可用期間偶發網路錯誤，疑似與中途改輸入框有關**（2026-08-17 CTH 回報，先不做）
   - 症狀：查 A 公司的可用期間時，**在查詢還沒跑完就把輸入框改成 B 公司**，會跳
     網路錯誤。CTH 不確定是不是這個原因造成的，**需要查證，不要當成已確認的因果**
   - 若屬實，可能的方向（都未驗證）：查詢在背景執行緒跑，完成時回讀輸入框當下的
     值而不是發動當下的值；或前一次查詢沒取消就發第二次，兩邊搶同一個
     `net_retry` 狀態
   - **重現優先**：先想辦法穩定重現（記下打了哪兩家、間隔多久、`logs/app.log`
     當下那幾行），再談修法。錯誤訊息長什麼樣、是不是真的網路層例外，是第一個
     要確認的事——`net_retry.is_network_error()` 會沿 `__cause__` 走訪，
     非網路問題也可能被包成網路錯誤顯示

E11. **關閉視窗前沒有未儲存提示**（2026-08-17 CTH 提出，先不做；存檔按鈕
   位置那半已於 2026-08-18 完成，見 `docs/CHANGELOG.md`）
   - 沒存檔就切走 Tab3 或關主視窗，目前沒有任何提示，可能遺失編輯
   - 這條需要先問 CTH 範圍（只比對 Tab3 欄位、還是含 Tab1/2 的執行參數？後者
     通常不算「設定」，多半不用比），也要查主視窗 `WM_DELETE_WINDOW` 現在有沒有
     處理常式，動手前先問清楚再做

## F. 跨公司比較功能的延伸需求

F2. **估值倍數（P/E、EV/EBITDA、P/B 等）**（2026-08-20 CTH 提出，記錄用，未確認方向）
   - 前提：需要先有股價/市值資料來源，工具目前完全沒有市場數據，是比 F1 更大的擴充（要接股價 API/資料源）
   - **待研究，未動手**，等 F1 做完再說
## G. 2026-08-22 期間對齊／缺值／效能系列（G1、G3 已完成並移入 CHANGELOG）

> **⚠ 動手前必讀：`docs/superpowers/design-2026-08-22-period-alignment-and-gaps.md`**
> 那份是 CTH 逐項確認過的**完整規格**（每項含：為什麼、規格、動哪些檔案、測試要釘什麼、
> 怎麼驗收、風險、最容易踩的坑）。下面 G1~G8 只是索引與決策結論，細節不重複寫，
> 兩邊有出入時**以設計書為準**。

G0. **執行順序與相依性**（2026-08-22 定案）
   - 順序：**G1 → G3 → G2+G7 → G6 → G8 → G4/G5**
   - 相依：G6（缺季留空白欄）會走到「第 5 列不是完整 ISO 日期」那條路徑，而那正是 G1 要收掉的殘留錯值來源，所以 **G1 必須先做**
   - G2 與 G7 都改跨公司輸出版面，一起做，不要分兩次動 `comparison_writer.py`
   - **不要開平行 subagent**：G2/G6/G7 都改 `comparison_writer.py`（會衝突），而且每項都要跑全套測試（含真連 SEC 的 live 測試，實測併發會互相搶 SEC 頻寬，兩份 20 分鐘的測試併著跑變成 40 分鐘）。單一 worker 循序做最快

G2. **跨公司 `Compare_Data` 最上方加「日曆季 ↔ 財季」對應表** ← **CTH 已決策，可直接做**（規格見設計書 G2，與 G7 綁一起做）
   - **決策**：用對應表而不是「只寫財年開始月份」——對應表逐期從實際期末日算，公司改過財年自己就對，不需要例外處理
   - **決策**：下面的財務指標區塊只給日曆季，不重複財季
   - ⚠ 最容易壞的地方：插入區塊會把列號往下推，`write_snapshot_sheets()` 也吃 `block_ranges`，要有測試釘住 Snapshot 公式仍指到正確的列
   - 現在 `Compare_Data` 只有日曆季欄位標題（`2025Q2`）與期末結算日列。CTH 要能在原始 sheet 上看到各公司自己的財季（如 NVDA 的 `FY2026Q2`），**只在資料表上呈現，不用進圖表**
   - 設計要點：同一個日曆季欄位下，各公司的財季不同（NVDA 是 FY2026Q2、AMD 是 FY2025Q2），所以不能只加一列——要嘛每家公司一列財季備註，要嘛跟公司資料列並排成兩欄。動手前先跟 CTH 確認要哪一種版面
   - `comparison.ComparisonResult` 目前只留 `period_ends`，沒有留原本的財季標籤（`_aligned_labels()` 轉換後就丟了），要先多帶一個 `fiscal_labels: dict[ticker, dict[日曆季, 財季]]`

G4. **`Other (as reported)` overflow 區天生就會斷斷續續**（2026-08-22 查 G3 時一併確認，**可能不是 bug，是設計取捨，要 CTH 決定要不要處理**）
   - overflow 列的 key 是 XBRL concept name。公司隔幾年換一個 concept 或改標籤，同一個經濟意義的科目就會長出**第二列**，舊列從此整片空白。實測 NVDA 檔案裡 `Deferred income taxes` 出現兩列（一列缺 48 期、一列缺 18 期，兩列互補）、`Tax benefits from stock-based compensation` 也重複兩列
   - 另外 `_synthesize_q4()` 明講**只補模板列、overflow 列留空**（D0-1 就記錄過），所以每個合成 Q4 欄位在 overflow 區整片是空的
   - 要處理的話有兩條路：① 用 synonym 把同義 concept 合併成一列（風險：合錯就是把兩個不同科目加在一起）；② 維持現狀但在 Excel 上把 overflow 區標示成「原樣呈現、不保證跨期連續」。**建議先做 ②**

G6. **抓不到的季度不要跳過，留一整欄空白** ← **CTH 2026-08-23 決定要做，規格封閉可直接開工**（設計書 G6）
   - **CTH 原話**：「要做，我寧願他是空白欄」
   - ⚠ **缺漏判斷改版後的變化**（G6 是在那之前定的規格，接手時要知道）：插了空白欄之後，那一季不再是「欄位不存在」而是「欄位在但整排空」，所以會從判定 A（季度斷層）改由判定 D（整欄稀疏）報出來。**兩邊都抓得到不會漏**，而且使用者在 `Data_Financials` 上直接看得到那一欄是空的，比只在 Index 顯示更好
   - **判定公式**：`missing = round((下一期末日 − 這期末日).days / 91) - 1`。91 = 13 週 × 7 天，要寫成具名常數並註明來源
   - 1,482 對相鄰期間驗證：95.3% 落在 70~110 天（正常一季）；111~150 天那 16 筆**全部是 COSTCO**（16 週的第四季），`round(112/91)=1` 正確判為沒缺——**這就是不能用固定門檻的原因**；151~210 天那 52 筆都是 182~189，`round=2` 補 1 欄；**沒有任何 >210 天的案例**
   - 配套：單一缺口最多補 4 欄（真出現就是資料異常，不該無限生欄）
   - 附帶發現：SNOW 有 `2022-01-31 → 2022-01-31`（0 天）**同一期末日出現兩次**，是重複列不是缺季，另開 G13 查
   - **決策**：補到「**最早抓到的那一季**」為止，不往更早補
   - ⚠ 跟現有的「全空期間整欄拿掉」邏輯直接衝突，要靠 label 形態分辨「殘骸欄（`FY2009Q4`）」與「該有但沒資料的空白欄（`2025Q3`）」
   - ⚠ 尾端插空白欄會拉低 Index 完成度（`check_key_rows()` 只看最後 4 欄）。**先確認會不會誤傷，會的話回報 CTH，不要自己改評分標準**
   - ⚠ 會讓更多 TTM 比率變 `None`。這是正確的（本來就缺，以前看不出來），但要寫進 CHANGELOG 免得被當成回歸
   - 現況：欄位清單是「成功抓到什麼就放什麼」，某一季掛掉就整欄消失，畫面上 25Q3 直接跳到 26Q1，使用者與 AI 都看不出中間漏了一季
   - CTH 要的：缺的那一季**保留欄位、內容全空**，讓「有漏」這件事看得見
   - 要動的地方：① `fetcher_gaap._merge_financials()` 建 `all_qs` 時改成產生完整季序列 ② 跨公司 `comparison_writer.write_compare_data_sheet()` 同理 ③ 第 5 列期末日——缺的那格沒有真實日期，只能用 `_fiscal_period_end()` 反推年月（`2025-10`），但那不是完整 ISO，`fiscal_input._apply_to_sheet()` 會跳過該欄的公式，落回靜態值
   - **相依 G1**：上一點正是 G1 要收掉的殘留錯值路徑，G1 沒先做的話這裡會直接生出一堆錯的日曆季
   - **動手前要問 CTH**：往前補到哪裡為止？補到「最早抓到的那一季」還是使用者指定的起始年？不設界線會一路補到 2000 年生出幾十欄空白
   - 順帶檢查：`override_engine.check_key_rows()` 只看最後 4 欄，尾端插入空白欄會影響 Index 完成度分數（跟 G5 有關聯）

G7. **跨公司比較加一張「說明」sheet** ← **規格已定案（含 checkbox），可直接做**（設計書 G7，與 G2 綁一起做）
   - **CTH 2026-08-22 追加**：要有「本檔是否適用」的勾選欄，大多數條目可由實際資料算出來，使用者不必讀完九條猜哪條跟他有關
   - **必須新增的三條**：① **抓取失敗的公司**（最重要——現在失敗只寫 GUI log，檔案裡完全看不到，`INTC_NVDA_AMD_TSM_v3.xlsx` 檔名有 TSM 但裡面沒有 TSM）；② 符號照公司原始申報未做正規化；③ 數字是當初申報值還是含重編
   - **已查出原草稿有一條寫錯**：「Q4 是推算的，`Other (as reported)` 區不補」的後半句對這份檔案是錯的——跨公司可選指標只來自三張模板 + `RATIO_DEFS`，overflow 列不在選單裡（`main.py:1749`）。要刪掉後半句
   - **決策**：CTH 明講這張表未來會擴充（開發中發現新的定義問題就往裡加），所以**要做成資料驅動**（一個 list of (標題鍵, 內文鍵)），新增一條只要加一行 + 四個 locale 各加兩條，不可把文字寫死在版面程式裡
   - **CTH 的總結**：「重點就是一些定義問題要標在 sheet 讓使用者了解即可」——這張表是這一輪的核心交付物之一，不是附屬品
   - 目的：把這份檔案用到的定義一次講清楚，不要讓使用者自己猜
   - 要寫的條目（草稿，動手前跟 CTH 過一次）：
     1. 時間軸＝日曆季，判準「該季天數多數落在哪一季」
     2. 為什麼不用財季當共同欄位（各公司財年結束月不同）
     3. 為什麼不用期末日判準（13 週季末日會漂到下一季）
     4. 期末結算日列取同欄**最晚**的日期，以及它給 Snapshot 用的意義
     5. Q4 是「年報 − Q1 − Q2 − Q3」推算的，`Other (as reported)` 區不補
     6. 空白＝該期沒抓到，圖上顯示斷點、不假造連線
     7. 資料來源 SEC EDGAR XBRL；2009-2011 分階段上路，更早期間沒有結構化數字
     8. 單位（金額 $mm、比率 %）
     9. **這份的日曆季定義只適用跨公司比較**，單一公司輸出第 4 列用的是另一套（期末日判準），刻意不同
   - 跟 G2 一起做（同一次改 `comparison_writer.py`）

G8. **用「比較欄」當 fallback 補洞** ← **CTH 已決策，放最後做**（規格見設計書 G8）
   - **決策**：做成**通用的補洞機制**，不要寫死成「只補 pre-XBRL」（CTH：「看能不能套用到其他狀況」）
   - **決策（鐵則）**：**當期申報優先，比較欄只補洞，永遠不覆蓋已有的值**。重編造成兩版不同時以當期那版為準
   - ⚠ 驗收時除了看缺漏數下降，**還要比對「原本有值的期間數字有沒有變」——一格都不該變**
   - ⏸ 要不要在 Excel 上標示「這一格是從比較欄補的」——動手前問 CTH
   - 2009Q4 全空的直接原因：合成 Q4 需要 Q1/Q2/Q3，而 2009Q1 那份 10-Q 是純 HTML 沒有 XBRL（SEC 強制申報 2009-06 才對大型申報人上路）
   - **但那一期的數字其實拿得到**：10-Q 的 XBRL 除了當期還帶去年同期的比較欄。實測 NVDA 2010-05-02 那份 10-Q 的欄位是 `['2010-05-02 (Q2)', '2009-04-26 (Q2)']`——2009 那一季就在第二欄。我們的 `_current_q_col()` 只取第一欄，**比較欄整個丟掉**
   - 撿回來大約可以往前多補一年，也能讓 2009 的 Q4 合成成立
   - 要處理的問題：① 同一期間會在多份 filing 出現（當期一次、隔年比較欄一次），取哪一版？當期優先、比較欄只補洞 ② 公司重編（restatement）時兩版數字不同，要不要標示 ③ 比較欄的欄名同樣不能採信 `(Qn)`，要走跟 D0-6 一樣的日期反推
   - 另一條路（成本較高，列為備案）：SEC Company Facts API（`data.sec.gov/api/xbrl/companyfacts/CIK##########.json`），一次拿回該公司歷來所有 XBRL fact。edgartools 的 `company.get_facts()` 已經在用來抓流通股數。交叉補洞能力更強，但 restatement 取版本的問題更嚴重

G10. **`D&A` 與 `Capex` 的 concept 對照對某些公司完全失效**（2026-08-22 產六家比較檔時發現）
   - 實測 `output/_compare/Semis_6co_2020_2025.xlsx`（AMD/NVDA/AVGO/INTC/MRVL/LITE，2020-2025 共 24 期）：
     ```
     Revenue / Gross Margin / Operating Margin   六家都 24/24   ✓
     D&A     AMD 2/24、MRVL 0/24（其餘四家 24/24）  ✗
     Capex   NVDA 13/24（其餘五家 24/24）           ✗
     ```
   - **跟 G3 是不同問題**：G3 是 IS 區的 `D&A (CF memo)` 走錯路徑；這裡 CF 區的 `D&A` 對 NVDA 是滿的，對 AMD/MRVL 幾乎全空 → 是那兩家的現金流量表用了 `CF_TEMPLATE` 沒收錄的 XBRL concept
   - **2026-08-22 已排查完，根因跟原本推測的不一樣**：不是「公司用了我們沒收錄的 concept」，是 **edgartools 把 `standard_concept` 標錯了**。實測 AMD 2026-06-27 與 MRVL 2026-05-02 的 10-Q 現金流量表：
     ```
     AMD   standard_concept = NonoperatingIncomeExpense   label = "Depreciation and amortization"
     MRVL  standard_concept = NonoperatingIncomeExpense   label = "Depreciation and amortization"
     兩家的 Capex   standard_concept = CapitalExpenses    label = "Purchases of property and equipment"
     ```
     `NonoperatingIncomeExpense` 跟折舊攤銷毫無關係，但 `label` 寫得清清楚楚。`_match_is_row()` 優先比對 `std_concept`，比不中就漏掉
   - 另外看到 AMD 有兩列 `standard_concept` 是 `nan`（`Purchases of property and equipment`、`Stock repurchases for tax withholding...`），以及 `CapitalExpenses` 在 AMD 出現兩次（第二次是 "accrued but not paid"，**不可以加總，會重複計算**）
   - **修法選項**：① 幫 `D&A` 在 `CF_TEMPLATE` 補 `label_hint`，讓 std_concept 比不中時退回 label 比對（成本最低）；② 走 `SYNONYM_MAP`。**注意 ①/② 都要處理「同一個 concept 出現兩次」的去重**
   - **這一類問題在 G11（改用 companyfacts）之後會整個消失**——companyfacts 直接給原始 us-gaap concept 名稱，沒有 edgartools 的 `standard_concept` 轉譯層。所以動手前先確認 G11 的決策，不然大概率白做

G11. **改用 SEC companyfacts API 取數** ← **CTH 2026-08-23 決定：先不切換，主力維持 edgartools**
   - **CTH 原話**：「g11 先不要切，我們主力維持從 edgartool 抓取」
   - 平行路徑（`src/fetcher_facts.py` + `src/facts_mapping.py`）保留不刪，40 個測試維持綠燈。它同時是「第二個獨立資料來源」，做交叉驗證用得到
   - **✅ 重評已完成（2026-08-23 晚），結論：不換，這題可以收了。** 兩個獨立的理由，任一個單獨成立就足夠：
     1. **速度優勢在混合架構下幾乎消失。** 「快 215 倍」的前提是**完全不解 filing**。實測 ARLO 16 份 10-Q 的時間拆分：下載＋解 XBRL **佔 54%**（這步 `Data_Segments` 非做不可，而且它跟三表用同一個 `max_filings`），三表各自的 `to_dataframe()` 合計才 46%。**CTH 2026-08-23 確認 segments 要 20 年份、只有 8 季不可接受** → 那 54% 一分都省不掉 → 混合架構只快 1.9 倍，不值得付下面那些代價
     2. **H3 做完後重跑 `spike_verify_mapping.py`：83.96% 精確／95.17% 符號對齊，比 H3 之前的 92.82%／95.35% 還低。** 不是 facts 變差，是**現行路徑今天變好了**，差距反而拉開。最明顯的是 `Debt Repayments` 只剩 20.67%——H3-2 把那一列改成加總所有借款線，facts 那邊還是單一 concept。目標 99% 不但沒收斂，方向是反的
   - **要付的代價（不換就不用付）**：C 欄的公司原文標籤會消失（facts 沒有 presentation linkbase）、`Data_Segments` 結構上拿不到（fact 沒有維度欄位）、`Other (as reported)` 語意從「報表印出來但模板沒收的列」變成「tag 過但模板沒收的 concept」，會混進附註層的東西
   - **重評前要補的**：那條路**從來沒產出過一份完整 Excel**——只驗過取數層，沒跑過 `_merge_financials` → `ratios` → `excel_writer` 整條下游，比率與版面對不對都還不知道
   - **完整報告：`docs/superpowers/report-2026-08-22-g11-companyfacts.md`**（52 家逐格比對、所有決定與理由）
   - 現況：`src/fetcher_facts.py` + `src/facts_mapping.py` 已完成，40 個測試。**現有程式一行都沒動**
   - 實測：每家 0.34 秒 vs 現行 7.5 分鐘（**215 倍**）；最早涵蓋 2008-07（現行 2009-07）
   - 逐格比對 **92.8% 相同，符號對齊後 95.4%**。剩下的差異分三類，全部有解釋（符號慣例、現行路徑會加總而 facts 是單一 concept、現行路徑本身算錯）
   - **切換前要先做完的兩件事**：
     1. **符號慣例定案**——現行輸出自己就不一致（`Income Tax` 有 15% 的格子符號相反）。要定義「每列一個明確慣例」並在兩條路強制執行。**這是行為改變，要 CTH 拍板**
     2. **加總型的列**（`Investment Proceeds`／`Debt Proceeds`／`Debt Repayments`／`Total Non-op`）要在 facts 這邊補上跟 `_sum_matching_rows()` 一樣的加總邏輯
   - 做完再跑 `scripts/spike_verify_mapping.py`，目標 99% 再談切換
   - **限制（不是待辦）**：companyfacts 沒有維度資料 → `Data_Segments` 非走解 filing 不可；沒有 presentation linkbase → 公司自報標籤與 `Other (as reported)` 的語意會改變。建議混合架構
   - **CTH 2026-08-22 決策（符號）**：**一律照公司原始申報，不做正規化**（「尊重公司原始資料，使用者要查找時會自己處理」）。`facts_mapping` 已全面移除 `negate`。附帶說明：反推出來的符號旗標仍有診斷價值——它揭露現行路徑的符號本身就不一致（AAPL 的 Capex 早年正、近年負），而 `ratios.py` 早就用 `abs()` 包住 Capex／Interest Expense，所以比率不受影響
   - **CTH 2026-08-22 決策（模板列）**：`Other Operating Expense` **不刪，改對照名字**。原本判定「建議刪除」是錯的——52 家都抓不到是因為模板猜的 `OtherOperatingExpenses`／`OtherOperatingExpense` 根本沒有公司在用，實際 tag 的是 `OtherCostAndExpenseOperating`／`OtherOperatingIncomeExpenseNet`（各 11 家）。跟 G10 同一類問題。facts 側已修，**現行路徑側的 concept 名稱也要跟著修**
   - `Free Cash Flow` 本來就是 DERIVED（OCF − Capex），不動。其餘列都有對應

G13. **同一個期末日出現兩次（重複列）**（2026-08-22 做 G6 判定規則分析時發現）
   - 實測 SNOW 的 `Data_Financials(Q)` 有兩欄期末日都是 `2022-01-31`
   - 52 家、1,482 對相鄰期間裡只有這 1 筆，屬於低頻但真實的資料問題
   - 要查：是同一期被兩份 filing 用不同財季標籤收進來（label 沒撞號但期末日撞了），還是別的原因
   - G6 的補欄規則不會被它影響（`round(0/91)=0`），但重複欄本身會讓使用者看到兩個一樣的日期

## H. 單一公司抓取的核心體檢（2026-08-22 CTH 指定為最重要的任務）

> CTH 的五個要求：① 數字都要對且沒有缺漏 ② 有缺漏要能發現 ③ 呼叫要快
> ④ 既有公版項目要正確 ⑤ 確認公版完整性

H0. **已有的體檢數字（2026-08-23 H3 修完後重跑，52 家實測，可直接當基線）**
   - 現行路徑：97 個模板列中，**47 列達到「≥45 家有值且填滿率 >90%」**（H3 之前是 40 列）
   - **「現行抓不到但 companyfacts 抓得到」原本有 7 列，H3 修掉其中 3 列**：`Other Operating Expense`（0 → **13**）、`Deferred Revenue, current`（3 → **28**）、`Current Portion of LT Debt`（23 → 21，但矛盾判定從 25 家降到 3 家，見 CHANGELOG H3）
   - **剩下 4 列改 concept 名字救不了**（公司沒在報表表面單獨列，只有附註有）：`Accrued Compensation`、`Op. Lease Liabilities, current`、`Operating Lease ROU Assets`、`Amortization of Intangibles`。詳見 H3-1
   - 資料在 `output/_spike/`（52 家的 facts JSON 與答案卷快取），重跑體檢不用再打網路。**注意答案卷的抓取窗不一致**：AAPL/ADBE/AMD/AVGO/COST/GOOGL/INTC/META/MSFT/NVDA/TSLA/WMT 十二家是全部 filing（44~69 期），其餘 40 家是 `max_filings=16`（約 21 期）。重建時要沿用同樣的參數，不然逐列覆蓋率沒得比

H1. **⚠ 新發現：companyfacts 對「現金流量表的流量項」只有約 25% 覆蓋**（做 H0 體檢時發現，**修正 G11 的評估**）
   - 症狀：`Capex`／`Dividends Paid`／`Change in Receivables`／`FX Effect on Cash` 等 CF 流量列，facts 的填滿率中位數只有 **25%**（＝一年四季只拿得到一季）
   - **根因**：公司在 XBRL 裡把現金流量表的項目 tag 成 **YTD 累計**，不是單季。只有 Q1 的 YTD 剛好等於單季，所以 `classify_period()` 篩「80~100 天」只撈得到 Q1
   - **這不是 facts 的缺陷，是我的實作只收單季 duration**。facts 其實有那些 YTD fact，而且**自帶精確起訖日**，做 YTD 拆算比現行路徑更可靠（現行是靠欄名猜哪欄是 YTD）
   - **修法**：`fetcher_facts` 要加一層「同一財年內用本期 YTD − 上期 YTD 還原單季」。判定用 `start` 相同、`end` 遞增這個結構性特徵，不用猜
   - **影響 G11 的結論**：facts 不是 CF 的 drop-in replacement，要補這一層才算數。IS/BS 不受影響（那些本來就 tag 單季／時點值）

H2. **公版（模板）內容改成使用者可選** ← **CTH 2026-08-23：確認就是「讓使用者選公版格式」，不是目前最急迫的問題，保留方向未來再處理**
   - **關鍵取捨已定調（A 案）**：**維持固定列位，只切換顯示／隱藏**。關掉的列還在、只是隱藏或留白，**下游完全不受影響**（`financial-assistant` 的 `read_excel.py`、使用者自己的 Excel 公式、跨檔案 `MATCH` 都靠固定列位取值）
   - 否決的 B 案：接受列位浮動、關掉的列真的消失。好處是檔案更乾淨，但代價是打斷所有既有 Excel 公式，換不到
   - 定了 A 之後才好談的其餘五個問題（選擇粒度、預設清單、存哪裡、跟 `Data_Ratios` 的相依、overflow 能不能升級成正式列）見下方原始記錄
   - 需求：讓使用者決定公版應該有哪些內容
   - 要想清楚的問題（動手前先跟 CTH 過一輪）：
     1. **選擇的粒度**：整列開關？還是分「一定要有／有就抓／不要」三段？
     2. **跟機器鍵的關係**：`Data_Financials` 的列位目前是固定的，下游（`financial-assistant` 的 `read_excel.py`／使用者自己的 Excel 公式）靠固定列位取值。**列位一變下游全壞**——是要維持固定列位只切換顯示，還是接受列位浮動？這是最關鍵的設計取捨
     3. **預設值**：新使用者看到的預設清單是什麼？用 H0 的體檢結果（≥45 家有值那 40 列）當預設？
     4. **存在哪**：`config.json` 還是每個 ticker 各自一份？換公司要不要換清單（金融股跟製造業要的列本來就不同，見 D8）
     5. **跟 `Data_Ratios` 的相依**：比率是從模板列算出來的，關掉某列會讓哪些比率變空？要不要在 UI 上提示
     6. **overflow 區**：使用者能不能把 `Other (as reported)` 裡的某列「升級」成正式列
   - **不要在 G11 決策之前動手**——公版列的來源（concept 對照）如果要換，先確定換不換

H4. **⚠ 公司自訂延伸 tag（`nvda_` / `tsla_` / `goog_` 這種）完全抓不到——模板沒有 label 比對層**（2026-08-23 端到端實測 NVDA 發現，**優先度高**）
   - **實測證據**：NVDA 的 `Capex` 在 57 期裡只有 36 期有值，**年報 17 年裡有 13 年整年抓不到**。根因是 NVDA 從 FY2019 到 FY2023 用自己的延伸 tag `nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets`，FY2024 起才改用標準的 `us-gaap_PaymentsToAcquireProductiveAssets`
   - **連鎖影響**：年報那一格空白 → `_synthesize_q4()` 算不出 Q4（年報 − Q1 − Q2 − Q3）→ **2014~2023 每一年的 Q4 Capex 都空**，`Free Cash Flow` 跟著一起空
   - **為什麼現行架構救不了**：`_match_is_row()` 有三層（std_concept → concept 正則 → label），但**模板的 6-tuple 只餵得進前兩層**。延伸 tag 的 concept 名字每家公司都不一樣，只有 label 對得上——NVDA 那幾年的 label 一直是「Purchases related to property and equipment and intangible assets」，穩定得很
   - **修法**：模板 tuple 加第 7 欄 `label_fallback`，把 `_match_is_row` 已經支援的第三層接起來。動到全部 97 列的結構，**要獨立一輪做，並且逐列前後對照**（`_match_is_row` 的第三層很寬，亂加會誤抓）
   - **範圍未量化**：目前只確認 NVDA 的 Capex。動手前先掃 102 家的 overflow 區，統計「有多少非 `us-gaap_` 開頭的 concept，其 label 對得上某個模板列」——那就是這個問題的實際規模

H3-1. **剩下的模板列缺漏：多數是「公司沒在報表表面單獨列」，不是 concept 對照錯**（2026-08-23，H3 主體已完成搬進 CHANGELOG，覆蓋率 40/97 → 47/97）
   - **`Op. Lease Liabilities, current` 14/52 家**（最大的一批）：實測 AMD／AMZN／ARLO／COST／GOOGL／MSFT／MU／NVDA／ORCL／PANW／SWKS／TSLA **十二家全部**只在資產負債表表面列「非流動」那條（`OperatingLeaseLiabilityNoncurrent`），流動部分併在「其他流動負債」裡、只有附註才拆開。**改 concept 名字救不了**，現行逐份解 filing 的路徑結構上拿不到
   - `Change in Inventories` 10 家（ADI／CVX／KO／MCD／NEE／PFE／XOM 等）同一類：現金流量表沒有單獨的存貨變動列，整包在「changes in operating assets and liabilities, net」
   - **CTH 2026-08-23 決定：先當「暫時性已知限制」，不動手。** 這是**擱置、不是結案**——資料是拿得到的（`fetcher_facts` 走 companyfacts 讀得到附註層的 fact：`Op. Lease Liabilities, current` 44/52、`Deferred Revenue, current` 31/52、`Accrued Compensation` 35/52），只是現行逐份解 filing 的路徑結構上碰不到。想解的時候有現成的第二資料源可以接
   - **重啟這條的時機**：(a) 使用者實際回報這幾列的空白造成困擾；(b) 之後有其他理由要動 `fetcher_facts` 的接線時順手一起做
   - **受影響的列（Index 會一直標紅，這是預期行為不是 bug）**：`Op. Lease Liabilities, current` 14 家、`Change in Inventories` 10 家、`Op. Lease Liabilities, LT` 4 家；另外 `Accrued Compensation`、`Operating Lease ROU Assets`、`Amortization of Intangibles`、`Finance Lease Liabilities, LT` 的覆蓋率偏低也是同一個成因

H3-2. **「中間有洞」對流量列會過度報警**（2026-08-23 發現，需 CTH 決定要不要調整判準）
   - `Acquisitions` 25/52、`Debt Proceeds` 24 家、`Short-term Debt` 15 家仍在榜首，但這些是**episodic 流量列**——公司沒併購、沒借款的那一季，通常是整條列不寫，不是寫 0
   - 拿 companyfacts 交叉驗證 `Acquisitions` 的 290 個洞：**只有 42% 在 companyfacts 裡找得到那個期末日的數字**，其餘 58% 是真的那季沒有併購
   - `data_quality` 的 B 判斷（首末有值之間不能有空格）對存量列（資產負債表）誤判率確實是 0，但對流量列不是。**選項**：(a) 維持現狀、接受這幾列一直標紅；(b) 給流量列另一套判準；(c) 只在洞的比例超過某個門檻才報
   - 不建議直接把這些列從 B 拿掉——那會連真的抓漏一起藏掉

## 執行順序建議（2026-08-22 更新）

> **維護規則**：做完的條目**直接從本檔刪除**，內容搬進 `docs/CHANGELOG.md`。
> TODO 只留「還沒做的」與「還沒決定的」，不累積已完成清單，否則會無限變長、
> 接手的人分不出哪些還有效。歷史紀錄查 CHANGELOG 或 git log。

| 順位 | 項目 | 需要 API？ | 需要人？ | 說明 |
|---|---|---|---|---|
| 1 | **G11 切換決策** | 否（spike 已做完） | 是 | 平行路徑已驗證完成（92.8%／符號對齊後 95.4%）。要 CTH 決定符號慣例與是否切換。做完會讓 G8／G9／G10 全部不需要 |
| 3 | G2 + G7（對應表 + 說明 sheet） | 否 | 是（說明條目的措辭要 CTH 過） | 跨公司輸出改版，兩項一起做 |
| 3 | G6（缺季留空白欄） | 否 | 是（補到哪為止已決定，但要確認不誤傷 Index 完成度） | G1 已完成，相依已解除 |
| 4 | B2 skill 端抽取 | 否 | 是（skill 設計） | 介面是 `cli.py press-release --json`。Non-GAAP 現在整個關閉，E2 等後續 GUI 工作卡在這條後面 |
| 5 | D8／D0-2／D0-5／D9 一次決策 | 部分否 | 是 | 都是「已知限制要不要修」的產品判斷題，建議合併成一次決策對話 |
| 7 | E 系列 GUI 細節（E2/E3/E5/E11） | 否 | 部分要 | 多半已標「先不做」或「待重現」，不影響資料正確性 |
| — | G4 overflow 標示 | 否 | 是 | 建議只做標示，不做 synonym 合併 |
| — | G8／G10 | — | — | **等 G11 決策，做了大概率白做**（G9 已完成） |
| — | F2 估值倍數 | 是（要股價來源） | 是 | 待研究，未確認方向 |

## D. 待 CTH 決定的已知限制

D10. **⏳ 暫時性限制：只寫在附註、沒印在報表表面的科目抓不到**（CTH 2026-08-23 決定先擱置，詳見 H3-1）
   - **這條標「暫時性」是有意義的**——資料拿得到，只是現行路徑碰不到。`fetcher_facts`（companyfacts）讀得到附註層的 fact，隨時可以接。不是「這個資料不存在」那種永久限制
   - 症狀：`Op. Lease Liabilities, current`、`Change in Inventories`、`Accrued Compensation`、`Amortization of Intangibles` 等列，Index 會標紅或留白
   - 判斷方式：印出那份 filing 的報表 dataframe，如果整張表裡沒有那個 concept，就是這一類，**改 concept 名字沒用**

D0-2. **多股別公司抓不到期末流通股數**：PLTR／GOOGL／META `company.get_facts()` 裡 `dei:EntityCommonStockSharesOutstanding` **0 筆**（只有 `EntityPublicFloat`），因為 Class A/B/C 是分開標的。TSLA 61 筆、COHR 62 筆、BRK.B 7 筆正常。連帶 BVPS／FCF per Share／流通股數 YoY 空白。`output/_final/META.xlsx` 現在就有這個洞。
   - **研究記錄（2026-08-12）**：`fetcher_gaap.py` 只查單一 XBRL concept，多股別公司這欄位本來就是空的（各股別分開報）。修法要改成按股別分別抓再加總，但程式碼裡已有註解記錄過「連單一股別公司這欄位都不乾淨」的踩坑史（抓到的是財報日後幾週的股數，非期末當天）——多股別可能根本沒有乾淨來源。**風險：中偏大**，搞不好要接受「這幾家就是空」當已知限制，或退而求其次抓 `EntityPublicFloat` 換算（精確度打折）

D0-5. **期間標籤是公式、沒有快取值 → 來源檔關著時跨檔案 `MATCH` 抓不到**（2026-08-08 確認）
   - CTH 已確定用**跨檔案讀取**當日常工作方式（不在 `Data_*` 裡加欄、另開工作檔），所以這條從邊角問題變成主要使用路徑上的坑
   - 成因：第 1、3、4 列是 Excel 公式，openpyxl 不算公式也不寫 `<v>` 快取值。來源檔**開著**時 Excel 會重算（`fullCalcOnLoad = True`）沒問題；**關著**時外部參照只讀得到檔案裡的值，那裡是空的
   - 第 5 列（期末結算日）是靜態文字，不受影響——目前的建議解法就是叫使用者拿第 5 列當 `MATCH` 的 key
   - 可能修法：寫檔時同時寫入 Python 算好的快取值（`fiscal_input` 已有 `fiscal_quarter_of()` 參考實作，本來就是公式的規格）。openpyxl 不支援 formula + value 並存，要另外處理
   - **要 CTH 決定做不做**：不做就是文件講清楚用第 5 列；做了才能直接用 `FY2026Q1` 當 key
   - **研究記錄（2026-08-12）**：兩條路——(a) 存檔後直接改 xlsx 內部 XML 把快取值塞進公式旁邊，**風險中**，動底層 XML 較脆弱；(b) 乾脆改成純值不用公式，**風險小**，但會犧牲「改 Index B4 財年起始月即時連動 1/3/4 列」這個功能（而且 B4 本來就只能在 Excel 裡改，程式端本來就抓不到那個編輯動作去重算，這個「即時連動」的價值本身也可以重新評估）。這項是取捨題不是純技術題

D8. 金融股（GS/JPM 等）獨立模板：現行 IS/BS 模板對金融股部分欄位空白，需另建模板。低優先。
   - **研究記錄（2026-08-12）**：程式碼裡目前除了印一行警告訊息，**完全沒有任何金融股特殊處理**——現有科目對照表直接套用，對不上的欄位就是空，不是 bug 是模板天生沒設計給銀行股。真要做要重新研究銀行/券商專屬 US-GAAP 科目（存款、放款、備抵呆帳...），等於另建一整套模板。**風險：大**。這其實是「要不要用這工具抓銀行股」的產品決定，不是技術問題，建議先想清楚要不要再談技術

D9. **外國私人發行人（Foreign Private Issuer）抓不到財報**（2026-08-20 CTH 回報）
   - 現象：抓 NBIS（Nebius Group）出現 `[NBIS] 抓取失敗 -> ValueError`，GUI 只顯示例外類型不顯示訊息（避免洩漏 URL/key），看不出原因
   - **研究記錄（2026-08-20）**：實際重現拿到完整訊息是 `fetcher_gaap.py:2030` 丟出的 `No 10-Q filings found for ticker 'NBIS'`。根因是 NBIS 這類外國私人發行人不申報 10-Q/10-K，改申報 **20-F**（年報）與 **6-K**（相當於季報/重大訊息），現有程式只查 `form="10-Q"` / `form="10-K"`，架構上就抓不到。不是 bug，是已知限制
   - 要支援的話得另外解析 20-F/6-K 的 XBRL 結構（跟現有 10-Q/10-K 的科目對照、期間切分邏輯不保證通用），工程量不小。**有空再做，先記錄**
   - **第二個實際個案（2026-08-21）**：跨公司比較測試 INTC/NVDA/AMD/TSM 時，TSM（台積電 ADR）同樣抓取失敗（`ValueError`），根因跟 NBIS 一致——TSM 也是外國私人發行人，只交 20-F/6-K。**CTH 問：20-F/6-K 有沒有機會補進去？**
   - **研究結果（2026-08-21，實測 TSM／UMC／NBIS 三家）**：
     1. **20-F 有 XBRL，6-K 沒有——確認，不是猜測**。直接呼叫 `filing.xbrl()` 實測：TSM／UMC／NBIS 三家最新一份 20-F 的 `xbrl()` 都回傳有效物件（`True`）；三家最新一份 6-K 都印出 `No XBRL attachments found`，`xbrl()` 拿不到資料。這代表 **6-K（季報等效）這條路線基本不可行**——不是「格式不固定，有些公司才有結構化數字」的機率問題，是這三家的最新 6-K 全部只有 PDF/文字附件，沒有 XBRL 可解析，符合原本「6-K 常只是包法說會新聞稿附件」的猜測
     2. **20-F 用 IFRS 命名空間，不是 US-GAAP——這是最大的工程量落點**。實際拉 TSM 20-F 的 XBRL `element_catalog`，1114 個科目**全部**是 `ifrs-full_` 前綴（如 `ifrs-full_Revenue`），一個 `us-gaap_` 都沒有。現有 `fetcher_gaap.py` 的 `IS_TEMPLATE`／`BS_TEMPLATE`／`CF_TEMPLATE` 整套科目對照表是針對 `us-gaap:` concept 建的，對 IFRS 科目**完全不適用**——不是換個 `form="20-F"` 參數就抓得到，是要另外設計一套 IFRS 科目對照表、可能還要處理 IFRS 特有的報表結構差異（例如 IFRS 允許的資產負債表排列、揭露顆粒度跟 US-GAAP 不完全一樣）
     3. **結論：只做 20-F 年報支援，範圍與工程量都不小，6-K 季報這條路線目前看起來不可行（沒有結構化資料源）**。原本設想的「先做 20-F 再看要不要做 6-K」變成「6-K 這步大概率做不了，只剩 20-F 值不值得單獨做」的判斷——年報頻率能不能滿足需求是 CTH 要考慮的重點（只有年報、沒有季報的比較功能，對分析師的用途打了折扣）
   - **風險：中偏大**，本質上是另建一套 IFRS 科目對照表與解析邏輯，工程量與現有 10-Q/10-K US-GAAP 模板相當，不是小改。值不值得做要看 CTH 覆蓋的外國發行人數量多不多，以及「只有年報沒有季報」能不能接受
