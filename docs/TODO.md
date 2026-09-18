# TODO

> **專案定位（2026-08-03 確立）**：這個程式只做一件事——**把 SEC EDGAR 的原始財務資料抓好**。
> 後續的判讀、分析、報告一律交給外部 skill。任何「幫使用者判斷」的功能都不屬於本專案範圍。

> **本檔維護規則（2026-08-22 CTH 指示）**：**做完的條目直接刪掉**，內容搬進
> `docs/CHANGELOG.md`。這份只留「還沒做」與「還沒決定」的事，不累積已完成清單。
> 想查歷史看 CHANGELOG 或 git log。

## ★ 最優先（2026-09-18 CTH 指定）

P1. **中斷續跑的保證要守住，而且要有人定期驗證**
   本地資料庫抓滿一次要 11 小時，所以「半夜斷掉不能白費」是這套東西能不能
   用的前提，不是加分項。2026-09-18 實測確認三種中斷點都能自動接續：

   | 中斷在哪 | 重跑行為 | 實測數字 |
   |---|---|---|
   | 公司與公司之間 | 已完成的整家跳過 | 0.7 秒/家（218 家約 2.5 分鐘掃完） |
   | 某家抓到一半 | 只補缺的那幾份 | 移走 5 份 → 命中 40、只下載 5 |
   | 檔案寫到一半 | 半截檔不被採信 | 殘留 `.tmp` 與壞掉的 JSON 都判無效 |

   靠的是三件事，**任何一件被破壞都會讓中斷續跑失效**：
   1. `save_filing()` **逐份即時落檔**，不是整趟跑完才寫
   2. `atomic_write_json()` 走 tmp + `os.replace()`，正式檔永遠是完整的舊版
      或完整的新版，沒有中間狀態
   3. `plan_ticker()` 用「SEC 清單 vs 本地 accession」比對，所以本地有多少
      就從多少接下去

   **待辦**：
   - (a) ✅ **2026-09-18 完成**：四條測試釘住（`test_local_db.py` 的
     `test_an_interrupted_company_resumes_from_what_is_already_on_disk`／
     `test_a_completed_company_is_skipped_entirely_on_rerun`／
     `test_a_truncated_cache_file_is_never_trusted`／
     `test_orphaned_tmp_files_are_not_mistaken_for_filings`）。
     **做過反向驗證**：把 `plan_ticker()` 改成一律跳過 → 2 條變紅；
     `.tmp` 那條要把 `glob("*.json")` 與 `ACCESSION_RE` **兩道防線都拆掉**
     才變紅（拆一道另一道還擋得住），確認測試釘的是行為不是實作
   - (b) 殘留的 `.tmp` 檔會累積（不影響正確性，會被忽略，但佔空間）。
     `update-db` 開始前順手清掉同 ticker 目錄下的 `*.tmp` 即可
   - (c) 重跑的固定成本約 3-4 分鐘（掃過已完成的 218 家 ＋ 每段 Python 啟動）。
     可接受，但如果之後名單擴到上千家就要重看

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
   - **2026-09-18 查證：上面兩個猜測都查過，都不成立**
     1. `_preview_scan_worker`（`main.py`）的 `ticker` 是執行緒**啟動當下用參數
        傳進去**的，不是完成時回讀輸入框——「回讀當下值」這條排除
     2. `net_retry.py` 整個模組沒有任何 module-level 共用狀態，
        `with_retry`／`is_network_error` 都是純函式——「搶 `net_retry` 狀態」
        這條也排除
     3. 掃描鈕本身**確實有擋重複點擊**：`_run_preview_scan` 一開始就把按鈕
        disable，`_update_button_states`（`main.py:2388`）也會在
        `_scan_running` 為真時持續讓按鈕維持不可按，回讀輸入框的
        FocusOut 查詢（`_tab1_lookup_worker`）本身把所有例外吞掉、不會冒出
        網路錯誤訊息給使用者看——目前程式路徑上找不到「中途改輸入框」能
        直接觸發使用者看到網路錯誤的機制
   - **重現優先，還沒重現前不要再猜方向**：先想辦法穩定重現（記下打了哪兩家、
     間隔多久、`logs/app.log` 當下那幾行），再談修法。錯誤訊息長什麼樣、是不是
     真的網路層例外，是第一個要確認的事——`net_retry.is_network_error()` 會沿
     `__cause__` 走訪，非網路問題也可能被包成網路錯誤顯示

E11. **關閉視窗前沒有未儲存提示**（2026-08-17 CTH 提出，先不做；存檔按鈕
   位置那半已於 2026-08-18 完成，見 `docs/CHANGELOG.md`）
   - 沒存檔就切走 Tab3 或關主視窗，目前沒有任何提示，可能遺失編輯
   - 這條需要先問 CTH 範圍（只比對 Tab3 欄位、還是含 Tab1/2 的執行參數？後者
     通常不算「設定」，多半不用比），也要查主視窗 `WM_DELETE_WINDOW` 現在有沒有
     處理常式，動手前先問清楚再做

## F. 跨公司比較功能的延伸需求

> **⚠ 主體（F1）早就做完了，不要被下面的條目誤導。** 跨公司比較（GUI 第 4 個分頁、
> `Compare_Data` / `Compare_Notes` 兩張 sheet、日曆季↔財季對應表、缺季留空白欄、
> 期間篩選支援月／日、圖表 X 軸真日期軸）已陸續完成並合併 master，內容看
> `docs/CHANGELOG.md` 搜尋「跨公司」與 F6／F7／G2／G7／G6 幾條。**這一段只剩下面
> 這條，還在等前提（股價／市值資料來源）才能動。**

F2. **估值倍數（P/E、EV/EBITDA、P/B 等）**（2026-08-20 CTH 提出，記錄用，未確認方向）
   - 前提：需要先有股價/市值資料來源，工具目前完全沒有市場數據，是比比較功能本身更大的
     擴充（要接股價 API／資料源）
   - **待研究，未動手**

## G. 2026-08-22 期間對齊／缺值／效能系列

> **⚠ 動手前必讀：`docs/superpowers/design-2026-08-22-period-alignment-and-gaps.md`**
> 那份是 CTH 逐項確認過的**完整規格**（每項含：為什麼、規格、動哪些檔案、測試要釘什麼、
> 怎麼驗收、風險、最容易踩的坑）。下面剩下的 G 條目只是索引與決策結論，細節不重複寫，
> 兩邊有出入時**以設計書為準**。

G0. **執行順序與相依性**
   - 剩下的順序：**G8 → G4/G5**
   - G8 會改變所有既有期間的資料來源優先序，前面幾項的驗收基準會跟著漂，所以排在最後

G4. **`Other (as reported)` overflow 區天生就會斷斷續續**（2026-08-22 查 G3 時一併確認，**可能不是 bug，是設計取捨，要 CTH 決定要不要處理**）
   - overflow 列的 key 是 XBRL concept name。公司隔幾年換一個 concept 或改標籤，同一個經濟意義的科目就會長出**第二列**，舊列從此整片空白。實測 NVDA 檔案裡 `Deferred income taxes` 出現兩列（一列缺 48 期、一列缺 18 期，兩列互補）、`Tax benefits from stock-based compensation` 也重複兩列
   - 另外 `_synthesize_q4()` 明講**只補模板列、overflow 列留空**（D0-1 就記錄過），所以每個合成 Q4 欄位在 overflow 區整片是空的
   - 要處理的話有兩條路：① 用 synonym 把同義 concept 合併成一列（風險：合錯就是把兩個不同科目加在一起）；② 維持現狀但在 Excel 上把 overflow 區標示成「原樣呈現、不保證跨期連續」。**建議先做 ②**

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

G13. **同一個期末日出現兩次（重複列）** ⚠ **成因已查明（2026-09-04），比原本記的嚴重；修法未做**
   - 實測 SNOW 的 `Data_Financials(Q)` 有兩欄期末日都是 `2022-01-31`
   - **原本的猜測（「同一期被兩份 filing 用不同財季標籤收進來」）是錯的。**

   ### 真正的成因（逐份實測，證據在下面每一條）

   問題出在**單一一份 filing**：`0001640147-22-000044`（2022-06-03 申報，
   **本該是 FY2023Q1、期末 2022-04-30**）。edgartools 解出來的三張表長這樣：

   ```
   IS: 期間欄 ['2022-01-31 (FY)']              → _current_q_col 挑到它
   CF: 期間欄 ['2022-01-31 (FY)']              → 同上
   BS: 期間欄 ['2022-04-30', '2022-01-31']     → 裸日期，_current_q_col 回 None
   ```

   **IS/CF 那兩張表根本沒有 Q 欄**——這是 edgartools 對這份 10-Q 的解析結果，
   上游就這樣。接著我們這邊：

   1. `_is_q_col()`（`fetcher_gaap.py:875`）把 `"FY"` 也算成期間欄 →
      `_current_q_col()` 回傳 `2022-01-31 (FY)`
   2. `_col_to_quarter_label()`（`fetcher_gaap.py:819`）走 `period.upper() == "FY"`
      分支 → 標籤 `FY2022`，期末日 `2022-01-31`
   3. **BS 的標籤是跟 IS 借的**（`fetcher_gaap.py:1487`）：
      ```python
      label = _col_to_quarter_label(is_q_col, ...) if is_q_col else _col_to_quarter_label(bs_col, ...)
      ```
      值取自 `bs_col`（＝**2022-04-30**），標籤卻取自 IS 的 `FY2022`

   ### ⚠ 所以症狀不是「看到兩個一樣的日期」，是「對的資料掛錯日期」

   那個 `FY2022` 欄**不是空的，也不是重複**——它有 36 格，全部是資產負債表的
   時點值，而且**那些是 2022-04-30 的餘額**：

   | 列 | `FY2022` 欄 | 真正的 `FY2022Q4` 欄 |
   |---|---|---|
   | Goodwill | 502,614,000 | 8,449,000 |
   | Accounts Receivable | 277,559,000 | 545,629,000 |
   | Total Non-current Assets | 2,680,966,000 | 2,051,055,000 |

   Goodwill 502,614,000 **正是 BS dataframe 裡 `2022-04-30` 那一欄的值**（已逐格核對）。
   SNOW 2022 年 3 月併購 Streamlit，商譽跳到 5 億是併購後才有的，2022-01-31 當天
   確實只有 8,449,000。兩欄共有值的 31 列裡 **28 列數字不同**。

   **使用者看到的是：兩欄標著同一個日期 `2022-01-31`，資產負債表互相矛盾，
   而且正確的是比較舊的那一欄。** 另外 FY2023Q1 這個標籤整個不見了
   （`_build_is_table` 的 dedup 是 `if label in periods: continue`）。

   ### 發生頻率

   掃 201 家的答案卷快取：「季表出現純年度標籤 `FY\d{4}`」與「期末日重複」
   **都只有 SNOW 一家，且兩者一對一重合**——同一個成因的兩個症狀。
   G6 的補欄規則不受影響（`round(0/91)=0`）。

   ### 修法未做——(a) 案的風險（2026-09-04 實測，**這是最重要的一段**）

   (a) ＝「季表的 builder 直接拒收 `(FY)` 欄」。看起來最直接，實測有四層風險：

   **風險 1（致命）：直接改 `_is_q_col()` / `_current_q_col()` 會打爆年表。**
   實測 SNOW／AAPL／JPM 各 3 份 10-K（共 9 份），**IS dataframe 的期間欄
   全部只有 `(FY)`，沒有任何一份有 Q 欄**：

   ```
   AAPL 0000320193-25-000079  cols=['2025-09-27 (FY)', '2024-09-28 (FY)', '2023-09-30 (FY)']
   JPM  0001628280-26-008131  cols=['2025-12-31 (FY)', '2024-12-31 (FY)', '2023-12-31 (FY)']
   ```

   **年表能建起來，完全就是靠 `_is_q_col()` 接受 `"FY"` 這件事。** 拒收 → 年表全空
   → `_synthesize_q4()`（Q4 ＝ 年報 − Q1 − Q2 − Q3）算不出來 →
   **每家公司每一年的 Q4 全部變空**。這不是「中風險」，是整條年度路徑歸零。

   **風險 2：要做成「只有季表路徑拒收」，改動面比想像大。**
   `_build_is_table`／`_build_bs_table`／`_build_cf_table` **季表年表共用同一組函式**，
   差別只在餵 `filings_k` 還是 `filings_q`（`fetcher_gaap.py:2692-2694` vs
   `2697-2699`，另有 `2737-2739` 的重試路徑）。要加一個「這趟是季表還是年表」的
   參數，得穿過那 3 個主 builder ＋ `_build_template_table`／`_build_dynamic_table`／
   `_build_segment_tables`，涵蓋 `_current_q_col()` 的 **8 個呼叫點**
   （1135／1231／1249／1455／1624／1638／2250／2331）與 **11 個 builder 呼叫點**。
   其中 1249／1455／1624 是**跨表借標籤**的用法（BS/CF 拿 IS 的欄來定標籤），
   模式沒跟著傳下去的話，BS 還是會繼承 IS 的 FY 標籤，等於白改。

   **風險 3：拒收之後 BS 的標籤會退化成「裸日期字串」。**
   `fetcher_gaap.py:1487` 的 fallback 是 `_col_to_quarter_label(bs_col, ...)`，
   而 `bs_col` 是裸日期 `2022-04-30`；`_col_to_quarter_label()` 的正則要
   `(Qn)`／`(FY)` 後綴，配不到就**原樣回傳**——那一欄的標籤會變成字串
   `"2022-04-30"`，不是 `FY2023Q1`。欄位排序（`sorted(periods.keys())`）與下游
   靠固定列位／標籤的 `MATCH` 都會怪。**所以 (a) 不能只做「拒收」**，還要補一條
   「BS 沒有 IS 標籤可借時，從裸日期推算財季」的路徑（`fiscal_quarter_of()` 現成的）。

   **風險 4：(a) 的結果是把「掛錯日期」換成「整季消失」，不是把資料救回來。**
   那份 filing 的 IS/CF 本來就只有年度欄，拒收之後 `_build_is_table` 直接
   `continue`；BS 那 36 格真實的 2022-04-30 資料，除非做了風險 3 的補救，
   否則會一起被丟掉。**換句話說 (a) 單獨做是「用資料變少換掉資料變錯」。**

   ### 其餘候選方向

   - **(b) 挑欄時優先找 `(Qn)`，全表只有 `(FY)` 時才退而求其次，並記進抓取帳本。**
     比 (a) 溫和（年表照常走 FY），但同樣要區分季/年模式，改動面跟風險 2 一樣大
   - **(c) 維持現狀，只在 Index 標紅。** 成本最低，但使用者看到的仍是兩欄矛盾的
     資產負債表
   - **(d) 只修 BS 的標籤來源（風險 3 那條路徑），不動 `_is_q_col`。**
     ⚠ 這是實測之後才浮出來的方向，**沒有評估過**：BS 有正確日期 `2022-04-30` 在手，
     不去借 IS 那個壞掉的標籤就對了。範圍只有 `_build_bs_table` 一處，
     但會改到所有公司的 BS 標籤來源，回歸面很大

   **要 CTH 決定的是：** 先做 (c) 讓使用者至少看得到警示，還是直接評估 (d)。
   **(a) 依實測結果不建議單獨做。**


## H. 單一公司抓取的核心體檢（2026-08-22 CTH 指定為最重要的任務）

> CTH 的五個要求：① 數字都要對且沒有缺漏 ② 有缺漏要能發現 ③ 呼叫要快
> ④ 既有公版項目要正確 ⑤ 確認公版完整性

H0. **已有的體檢數字（2026-08-24 重跑，**201 家**實測，基線見 `docs/template-coverage-baseline-2026-08-24.md`）**
   - ⚠ **這組數字是 H6／INTC Cash／G10（2026-08-25）之前的**。那三輪合計在 201 家上
     多救回 61 家次（H6 37、Cash `label_fallback` 11、D&A 13），達標列數應該會往上走，
     但**基線沒有跟著更新**——`gen_template_coverage_baseline.py` 吃的是
     `output/_spike/gaap_*.pkl`（已經比對完的結果），要先重跑
     `spike_derive_mapping.py`（201 家、真的抓 SEC）才反映得出來。下次有理由重建那批
     pkl 時再一起更新
   - 現行路徑：97 個模板列中，**44 列達標**（門檻改成比例：≥85% 的公司有值 且 填滿率中位數 >90%）
   - **樣本從 52 → 102 → 201 家，結論高度穩定**：達標列數 47（52 家）→ 46（102 家）→ 44（201 家）；每格三分類「我們抓到 / 真缺口 / 公司真的沒有」三次量測都是 **72~73% / 9% / 18%**
   - 44 vs 46 那兩列是**門檻邊緣的自然浮動**，不是回歸：`SBC` 87%→84%、`Diluted Shares` 85%→80% 掉出，`Additional Paid-in Capital` 81%→88% 補進
   - 「模板不適用」9 家：AFL／AIG／AMP／AXP／BAC／COF／GS／MET 都是金融股（D8 已知限制），**AZO（AutoZone）是唯一的非金融股，原因未查**
   - **「現行抓不到但 companyfacts 抓得到」原本有 7 列，H3 修掉其中 3 列**：`Other Operating Expense`（0 → **13**）、`Deferred Revenue, current`（3 → **28**）、`Current Portion of LT Debt`（23 → 21，但矛盾判定從 25 家降到 3 家，見 CHANGELOG H3）
   - **剩下 4 列改 concept 名字救不了**（公司沒在報表表面單獨列，只有附註有）：`Accrued Compensation`、`Op. Lease Liabilities, current`、`Operating Lease ROU Assets`、`Amortization of Intangibles`。詳見 H3-1
   - 資料在 `output/_spike/`（52 家的 facts JSON 與答案卷快取），重跑體檢不用再打網路。**注意答案卷的抓取窗不一致**：AAPL/ADBE/AMD/AVGO/COST/GOOGL/INTC/META/MSFT/NVDA/TSLA/WMT 十二家是全部 filing（44~69 期），其餘 40 家是 `max_filings=16`（約 21 期）。重建時要沿用同樣的參數，不然逐列覆蓋率沒得比

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
   - **G11 已定案不切換**，公版列的來源（concept 對照）不會換，這個前提已經解除，不再是阻擋動手的理由——但 H2 本身仍是「保留方向未來再處理」，不是目前最急迫的事

H4. **⚠ 公司自訂延伸 tag（`nvda_` / `tsla_` / `goog_` 這種）完全抓不到——模板沒有 label 比對層**（2026-08-23 端到端實測 NVDA 發現，**優先度高**）
   - **實測證據**：NVDA 的 `Capex` 在 57 期裡只有 36 期有值，**年報 17 年裡有 13 年整年抓不到**。根因是 NVDA 從 FY2019 到 FY2023 用自己的延伸 tag `nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets`，FY2024 起才改用標準的 `us-gaap_PaymentsToAcquireProductiveAssets`
   - **連鎖影響**：年報那一格空白 → `_synthesize_q4()` 算不出 Q4（年報 − Q1 − Q2 − Q3）→ **2014~2023 每一年的 Q4 Capex 都空**，`Free Cash Flow` 跟著一起空
   - **為什麼現行架構救不了**：`_match_is_row()` 有三層（std_concept → concept 正則 → label），但**模板的 6-tuple 只餵得進前兩層**。延伸 tag 的 concept 名字每家公司都不一樣，只有 label 對得上——NVDA 那幾年的 label 一直是「Purchases related to property and equipment and intangible assets」，穩定得很
   - **✅ 第一步已完成（2026-08-23）**：模板 tuple 加了第 7 欄 `label_fallback`，97 列全部補上，四個呼叫點都接線。目前只有 `Capex` 真的填了正則（`^purchases (?:of|related to).*propert`），因為那是唯一有實證的案例。非 live 測試 1114 passed
   - **⏳ 第二步（3b，數值指紋自動連結）：2026-08-25 量化完成，建議降級成長期項目，暫不做**。設計見
     `docs/superpowers/specs/2026-08-23-concept-rename-linking-design.md`。
     - **量化做法（3a）**：20 家（含 NVDA/TSLA/GE/PG）× 最新 6 份 10-Q，跑一次模板比對記下空格，
       在同一份 dataframe 未消化列裡數「不是 us-gaap_ 開頭、有值」的延伸 tag，取
       `min(當期空格數, 當期候選數)` 當理論上限（避免 1 個候選灌水算成救了好幾個空格）
     - **結果**：4013 個模板空格裡，理論上限只有 614 格（**15.3%**）。逐家差異大：
       GE 44%／JNJ 41%／TSLA 33% 較高，NVDA 只 5.6%（第一步 label_fallback 已經吃掉大半，
       印證第一步效果超預期、第二步邊際效益變小的猜測）；AAPL／HD 是 0%（overflow 沒有延伸 tag）
     - **這個 15.3% 還是高估**：人工抽查候選后，多數文不對題（JNJ「Other Operating Expense」
       空格配到 `jnj_GrossProfitPercentToSales`——毛利率百分比，跟營業費用無關；META「Inventories」
       配到 `meta_NonmarketableEquitySecuritiesCarryingValue`——非流通股權投資，跟存貨無關）。
       跟已經否決的「文字相似度配對」踩到同一類坑，只是方向相反（這次是字面沾邊但語意不同）
     - **少數看起來真的對得上**：MSFT `D&A` 空格配 `msft_DepreciationAmortizationAndOther`；
       AMZN `R&D Expense` 空格配 `amzn_TechnologyAndInfrastructureExpense`（AMZN 本來就不用
       R&D 這個名字）。但這種只占候選裡一小部分
     - **結論**：5~7 小時、風險中（動主流程取值）換來的可能只有個位數 % 的真實改善，且要另外
       設計怎麼濾掉像 JNJ／META 那種假候選。投報率明顯不如第一步。建議先不做，除非之後有
       具體案例（像 NVDA Capex 那種）逼出真正的需求
   - **companyfacts 沒有延伸 tag**（102 家實測，只有 us-gaap/dei/srt/ffd/ecd/invest 這些 SEC 標準 taxonomy）→ 第二資料源救不了這題，而且基線的〔真缺口〕KPI **低估了**這一類

H3-1. **剩下的模板列缺漏：多數是「公司沒在報表表面單獨列」，不是 concept 對照錯**（2026-08-23，H3 主體已完成搬進 CHANGELOG，覆蓋率 40/97 → 47/97）
   - **`Op. Lease Liabilities, current` 14/52 家**（最大的一批）：實測 AMD／AMZN／ARLO／COST／GOOGL／MSFT／MU／NVDA／ORCL／PANW／SWKS／TSLA **十二家全部**只在資產負債表表面列「非流動」那條（`OperatingLeaseLiabilityNoncurrent`），流動部分併在「其他流動負債」裡、只有附註才拆開。**改 concept 名字救不了**，現行逐份解 filing 的路徑結構上拿不到
   - `Change in Inventories` 10 家（ADI／CVX／KO／MCD／NEE／PFE／XOM 等）同一類：現金流量表沒有單獨的存貨變動列，整包在「changes in operating assets and liabilities, net」
   - **CTH 2026-08-23 決定：先當「暫時性已知限制」，不動手。** 這是**擱置、不是結案**——資料是拿得到的（`fetcher_facts` 走 companyfacts 讀得到附註層的 fact：`Op. Lease Liabilities, current` 44/52、`Deferred Revenue, current` 31/52、`Accrued Compensation` 35/52），只是現行逐份解 filing 的路徑結構上碰不到。想解的時候有現成的第二資料源可以接
   - **重啟這條的時機**：(a) 使用者實際回報這幾列的空白造成困擾；(b) 之後有其他理由要動 `fetcher_facts` 的接線時順手一起做
   - **受影響的列（Index 會一直標紅，這是預期行為不是 bug）**：`Op. Lease Liabilities, current` 14 家、`Change in Inventories` 10 家、`Op. Lease Liabilities, LT` 4 家；另外 `Accrued Compensation`、`Operating Lease ROU Assets`、`Amortization of Intangibles`、`Finance Lease Liabilities, LT` 的覆蓋率偏低也是同一個成因

H6-1. **hint 放寬後仍抓不到的案例——已診斷，還沒決定要不要修**（2026-08-25，
   H6 主體與已決定的部分都在 `docs/CHANGELOG.md`）
   - 原始資料仍在（**不要重跑**，一輪 12 分鐘）：`output/_hintsweep_201/hintsweep_201_result.txt`
     （201 家逐列逐家掃描）、`classification.md`（人工分類表）
   - 重跑指令：
     ```
     TICKERS=$(cat output/_hintsweep_201/tickers_joined.txt)
     ./venv/Scripts/python.exe scripts/diag_hintsweep.py "$TICKERS" > out.txt 2>&1
     ```
   - **⚠ 那份分類表有兩處判讀是錯的**（2026-08-25 查原始 10-Q 訂正，照著用之前先看這段）：
     ① 「CS&APIC 那 7 家是 concept 層失守」錯——ABT/AMP/AXP/COP/KR/MPC/UNP 七家全部是
     `us-gaap_CommonStockValue(Outstanding)` / `std_concept=CommonEquity`，concept 層好好的；
     ② 「SLB 退到 fallback_suffix 層」錯——`us-gaap_Cash` 一樣有 `std_concept=CashAndMarketableSecurities`
   - **① Common Stock & APIC 剩 COP 與 MPC**：COP 的 label 只有「Par value」、MPC 是
     「Issued – 995 million and 994 million shares (par value $0.01 per share…)」，
     **措辭裡完全沒有股票字樣**，放寬措辭救不了。要救只能改成「候選只有一列
     `CommonStockValue` 時不套 hint」這種結構性規則，那會影響所有公司，風險等級不同
   - **② NEE 的 Capex 是空的**（H6 之前抓到的是錯的數字）：它唯一的 `CapitalExpenses`
     候選是「Accrued property additions」（非現金揭露），真正的 capex 在
     `nee_CapitalExpendituresOfPublicUtilitiesFPLConsolidated` 這個延伸 tag。要救屬於
     **H4 第二步**的範圍，而且同一張表上還有 `nee_CapitalExpendituresOfFPL`（子公司）
     與 `Other capital expenditures` 兩個相似候選，不能無腦用 `^capital expenditures`
   - **③ 更零星的個案**（材料不足，不建議單獨修）：`Other Non-current Assets` IBM
     「Investments and sundry assets」／ISRG「Long-term investments」、`Dividends Paid`
     AMT／CHTR（只有少數股權分配）、`Accounts Receivable` SO（gross 口徑）、
     `Change in Inventories` AEP、`Finance Lease Liabilities, LT` ON、`Other Current Assets` CVX


## I. 本地 filing 快取（2026-09-03，Task 1-11 全部完成並已併入 master）

> **動手前必讀**：規格 `docs/superpowers/specs/2026-09-03-local-filing-cache-design.md`、
> 實作計畫 `docs/superpowers/plans/2026-09-03-local-filing-cache.md`（11 個 task，
> 逐條含測試碼）。執行紀錄與所有已判決的取捨在
> `.superpowers/sdd/2026-09-03-local-filing-cache/progress.md`（git-ignored，只在本機）。
> 已併入 master。已判決的取捨與效能天花板搬到 `docs/ARCHITECTURE.md`
> 「快取層：已判決的取捨與效能天花板」一節。

**這件事在做什麼**：抓同一家公司不要每次都重新對 SEC 打 20 年份 filing 再解析一次。
快取卡在**解析層與比對層之間**——存的是 edgartools 解出來的三張 DataFrame，
`IS/BS/CF_TEMPLATE` 那套科目比對永遠在快取之上即時重跑。所以以後改 hint regex、
加比率、調 Q4 合成邏輯都**不會**讓快取失效；但 **edgartools 升版會**，那條軸線靠
每份快取檔裡的 `edgartools_version` 欄位擋。詳細機制、四道閘、實測數字見
`docs/ARCHITECTURE.md`「本地 filing 快取」一節。

**目前狀態**：快取**已經生效**、已驗收、**已合併 master**（`83c3b62`）。
Tab3 的清除面板已接完（`main.py`），golden 逐格比對 0 格不同，
全套測試 1381 passed（`not slow`）＋ slow 58 passed。倍數不要引用單一數字
（冷 55~110s／熱 8~17s，隨 SEC 與本機負載變動），細節見
`docs/ARCHITECTURE.md`「本地 filing 快取」一節。

I5. **修正案（10-Q/A、10-K/A）現況本來就不抓**（承快取設計書，獨立議題）
   - `_list_filings()` 目前呼叫時 `amendments=False`（`fetcher_gaap.py:304` 附近），
     所以公司重編財報開的那份新 filing **現在就不在抓取清單裡**，不管有沒有快取
     都一樣抓不到。快取只保證「清單裡查得到的 filing」會被正確、即時補齊，
     不擴大也不縮小現有抓取範圍
   - 要不要處理是獨立的產品判斷題：抓修正案等於同一期會有兩份來源，
     要先決定「以哪一份為準」以及重編後舊數字要不要覆蓋——牽動 D11 的
     「as reported vs restated」既有立場（`Compare_Notes` 目前明講保留原始申報版）

## 執行順序建議（2026-08-22 更新）

> **維護規則**：做完的條目**直接從本檔刪除**，內容搬進 `docs/CHANGELOG.md`。
> TODO 只留「還沒做的」與「還沒決定的」，不累積已完成清單，否則會無限變長、
> 接手的人分不出哪些還有效。歷史紀錄查 CHANGELOG 或 git log。

| 順位 | 項目 | 需要 API？ | 需要人？ | 說明 |
|---|---|---|---|---|
| **0** | **P1 中斷續跑的自動化防線** | 否 | 否 | **CTH 2026-09-18 指定最優先**。現在只有手動實測，沒有測試擋著 |
| 1 | B2 skill 端抽取 | 否 | 是（skill 設計） | Non-GAAP 抽取改由 Claude Code 做 |
| 2 | G8（比較欄補洞） | — | 部分 | 風險最高，會動所有既有期間的資料來源優先序，所以排在最後 |
| 3 | D8／D0-2／D0-5／D9 一次決策 | 部分否 | 是 | 都是「已知限制要不要修」的產品判斷題，建議合併成一次決策對話 |
| 4 | E 系列 GUI 細節（E2/E3/E5/E11） | 否 | 部分要 | 多半已標「先不做」或「待重現」，不影響資料正確性 |
| — | G4 overflow 標示 | 否 | 是 | 建議只做標示，不做 synonym 合併 |
| — | F2 估值倍數 | 是（要股價來源） | 是 | 待研究，未確認方向 |
| — | I5 修正案要不要抓 | 否 | 是 | 產品判斷題，牽動「as reported vs restated」的既有立場 |

## J. 本地財報資料庫（2026-09-04 CTH 指定為本專案核心能力）

> 設計書：`docs/superpowers/specs/2026-09-04-local-filing-db-design.md`
> **結論：不需要重新構思架構。** 現有 `filing_cache.py` 已是正確形狀，
> 缺的是三塊「狀態與體驗」。
>
> **J1-J9 已完成並在 master 上**（見 `docs/CHANGELOG.md`）。落地在
> `src/local_db.py`（狀態層）＋ `src/cli.py` 的 `update-db`／`db-status`
> ＋ GUI 的「資料庫總覽」分頁。**`fetcher_gaap` 的抓取迴圈一行都沒動**
> ——「到底了沒」在迴圈外面推導。下面只留還沒做完的 J6。

J6. ⚠ **「抓到底」實際上是「抓到這個 CIK 的底」，不是公司歷史的底**
   （2026-09-05 跑 J5 batch 1 時發現，**這是新發現，設計書沒有預期到**）
   - **症狀**：134 家裡至少 9 家的可用歷史遠短於 18 年，而系統判定是「完整／已到底」
   - **實測**（ticker / cik / 最舊 10-Q）：

     | ticker | cik | 最舊 10-Q | 實際歷史 |
     |---|---|---|---|
     | BLK | 2012383 | 2024-11 | **只有 2 年** |
     | GEHC | 1932393 | 2023-04 | 3 年（2023 分拆，合理） |
     | DIS | 1744489 | 2019-05 | 7 年 |
     | DOW | 1751788 | 2019-05 | 7 年 |
     | CI | 1739940 | 2019-05 | 7 年 |
     | AVGO | 1730168 | 2018-06 | 8 年 |
     | LIN | 1707925 | 2017-10 | 9 年 |
     | DD | 1666700 | 2017-11 | 9 年 |
     | IR | 1699150 | 2017-08 | 9 年 |
     | （對照）CMCSA | 1166691 | 2008-05 | 18 年 |
     | （對照）KO | 21344 | 2008-04 | 18 年 |

   - **成因**：公司改組／合併／換註冊地會產生**新的 CIK**，舊 CIK 的申報不在新 CIK
     的 filing 清單裡。`Company("DIS")` 拿到的是 2019 年為了併購 Fox 成立的
     TWDC Holdco，併購前的迪士尼在另一個 CIK。BLK 是 2024 年 GIP／HPS 那兩樁交易
     之後的新控股實體，所以只剩兩年
   - ⚠ **這是靜默的，而且判定本身沒有錯**：那個 CIK 確實沒有更舊的申報，
     `reached_bottom = "no_more_filings"` 是對的，GUI 顯示「已到底」也是對的——
     但使用者會以為拿到 18 年、實際只有 7 年。**對財報工具來說這是最糟的失效模式：
     不報錯，只是資料比你以為的少**
   - 設計書把 `no_more_filings` 解釋成「公司比較晚上市」（META 2013、ARLO 2018），
     **那個解釋不完整**——同一個判定還涵蓋「改組換 CIK」，兩者的意涵完全不同：
     前者是「本來就沒有」，後者是「有，但在別的地方」
   - **可能方向（① ③ 都沒評估過）**：
     1. 用 SEC 的 former-name／predecessor 關聯去找舊 CIK 併進來
        （正確性風險：改組前後不一定是同一個經濟實體，DowDuPont 拆成三家就是反例）
     2. ✅ **2026-09-18 完成**：`Data_Meta` 新增 `Oldest Period`／
        `Oldest Period End`（`fetcher_gaap._build_meta_table`），Index 第一頁
        頂端 metadata 列也加了「資料最早自：{period}」（`excel_formatter.py`，
        放在「資料最新至」前面，最顯眼位置）。四個 locale 都補了翻譯，
        `zh_tw`／`zh_cn` 的說明文字直接寫明「⚠ 公司若改組換過 CIK，這不等於
        公司真正上市以來的歷史」。**這條只是把既有資料攤開來顯示，不改變
        任何抓取或判定邏輯**——`reached_bottom` 照舊，只是使用者現在不用
        自己拿 Quarters Available 回推涵蓋年數
     3. 在 `_meta.json` 記一個「疑似改組」旗標（最舊申報 > 2009 且不是已知的晚期 IPO）
   - **① ③ 2026-09-18 CTH 決定不做，維持現狀**——② 已經把「歷史有多長」
     攤在總覽的「年數」欄與 Excel 的 `Oldest Period`，使用者看得到就夠了。
     ③ 只能是啟發式警告（分不出「換過 CIK」與「比較晚上市」），① 有實質
     正確性風險（DowDuPont 拆三家＝併錯了不會發現）。要做的話 ① 建議走
     **手動指定前身 CIK**，把正確性判斷留在分析師手上，不讓程式猜。
   - ~~① ③ 還沒做，要 CTH 決定要不要做~~——② 已經把靜默變可見，① 有正確性
     風險（見上），③ 是進一步的主動提示，成本比①低但比②高

## D. 待 CTH 決定的已知限制

D11. **抓取結果會被當下的 SEC 狀況影響——同一家公司重抓可能得到更完整的資料**（2026-08-24 實測發現）
   - **實測證據**：`fetch50` 那一輪連續抓 50 家時，AXP／BMY／C／COP／ETN／HON／LLY／MRK／MS／PSX／WFC 的 `Operating Income` 只有 16/21 期；隔幾小時單獨重抓 LLY 兩次都是 **20/21 期**。DUK 的 `Revenue`／`Gross Profit` 同樣情形
   - **根因**：連續大量抓取時 SEC 偶發失敗 → `_filing_obj()` 重試用盡 → `_note_gap()` 記帳後跳過那份 → 少了年報就 `_synthesize_q4()` 合成不出 Q4 → **靜默少掉 4 格**（一年一格 × 5 份年報）
   - **同一份程式碼本身是決定性的**：單獨重抓 AAPL 兩次，逐格完全相同（含 overflow）。不確定性來自網路，不是程式
   - **為什麼會漏看**：帳本有記，但 Index 的完成度只會顯示成「整欄稀疏」或「中間有洞」，看起來像資料本身缺，不像「這次抓壞了」
   - **可能修法**：(a) 抓完自動偵測「有帳本缺漏就重試那幾份」——**2026-08-24 CTH 決定直接做**，見下方 D11-B；(b) ✅ **已完成（2026-08-17）**——Index 第一頁 A3 橘底那列＋`Data_Meta` 的 `Fetch Gaps` 就是這個，跟下面「資料完整度」區塊刻意分開顯示，見 `fetch_ledger.py`／CHANGELOG 2026-08-17；(c) 降低連續抓取的速率，未做
   - **對驗證樣本的影響**：`output/_spike/` 那 201 家裡可能有數家帶著這種缺漏，做逐列覆蓋率統計時要記得這是雜訊來源之一
   - **~~待查證的候選名單~~ 2026-08-24 查證：這條是誤記，作廢**。原文說「log 出現『自動修復：發現 N 期缺失需要修復重疊』的有 AIG／ALL／BK／CB／COF／DD／DOW／HIG／MET／NEM／OXY 共 11 家」——全庫搜尋「重疊」兩字，程式碼裡完全沒有這句話。實測 COF 等 6 家，log 印的其實是 `[COF] 警告：['Revenue', 'Capex'] 在 EDGAR 中無對應概念`，那是 **Override Engine**（關鍵指標抓不到、找不到相似科目可補）的警告，跟 filing 抓取失敗、跟「重疊」都無關，是完全不同的既有資料缺口類別。這 11 家不能當 D11 的證據名單，D11 唯一的實測證據還是回到最上面那組（AXP/BMY/C/COP/ETN/HON/LLY/MRK/MS/PSX/WFC/DUK）


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
   - **2026-08-25 追加一個具體需求：Cash 這一列對銀行要能「加總兩列」**
     （`CashAndDueFromBanks` + `InterestBearingDepositsInBanks`，有時再加
     `FederalFundsSoldAndSecuritiesPurchasedUnderAgreementsToResell`）。JPM 實測
     24.7bn vs 285.1bn，只取第一列等於填 8%（比留空更糟）。現行模板一列只對一個
     XBRL 列、不會加總——**「支援加總」是金融股模板的硬需求，不是選配**。
     會計原則（Reg S-X Article 9 / ASC 230）與實測數字見 `docs/CHANGELOG.md`
     「銀行現金口徑定案」那條
     - 現在仍空白的 6 家：AXP／BK／C／COF／JPM／WFC。**BAC 已經有值且口徑正確**
       ——它自己在 BS 上列了小計（229.7 = 28.1 + 201.6），不是我們特別處理的
   - **這些空白是正確行為，不要「修」**（2026-08-25 H6 確認，避免下次重做一次分析）：
     銀行／保險／交易所／鐵路／REIT 的 `Cost of Revenue` 共 29 家（AMT/AON/AXP/BAC/BK/
     BKNG/BLK/C/CCI/CME/COF/CSX/FDX/GS/HCA/ICE/JPM/MCD/MCO/MS/NDAQ/NSC/ODFL/PLD/
     SCHW/UNP/UPS/V/WFC）——概念上本來就沒有 COGS，填進人事費就是製造錯誤數字
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

   - ### ⚠⚠ 2026-09-18 實測推翻上面兩條結論（拿 ARM 測，前人測的是 TSM／UMC／NBIS）

     **外國私人發行人要分兩類看，不能一概而論。** 上面「20-F 用 IFRS」與
     「6-K 沒有 XBRL」對 TSM／UMC 可能成立，但 **ARM 完全不是那樣**：

     | | D9 原本的說法 | ARM 實測 |
     |---|---|---|
     | 20-F 的命名空間 | IFRS，要另建對照表 | **us-gaap 1,248 筆**（＋arm 自訂 180） |
     | 6-K 有 XBRL 嗎 | 「沒有」 | **33 份全都有**，其中 **9 份含完整財報** |
     | 6-K 能解析嗎 | — | **可以**，走既有 `filing.obj().financials` 就出三表 |

     - **實測**：2026-07-29 那份 6-K → 損益表 51 列、資產負債表 65 列、
       現金流量表 47 列，**結構跟 10-Q 拿到的一模一樣**
     - **季度覆蓋是完整的**：ARM 財年 3 月底結束，Q1→8 月、Q2→11 月、
       Q3→2 月各一份 6-K，**Q4 併在 20-F 年報**。所以 3 份季度 6-K ＋ 1 份
       20-F ＝ 一年四季齊全
     - **科目對照層一行都不用動**——已經是 us-gaap 了

   - ### 怎麼自動判斷一份 6-K 能不能解析（33/33 完全分離）

     6-K 是個大雜燴：同樣是 6-K，有的是季度財報、有的是股東會通知、
     有的是人事公告。33 份裡只有 9 份有財報。判準實測比較過三種：

     | 判準 | 準確度 | 說明 |
     |---|---|---|
     | **`R*.htm` 檔數量 > 1** | **33/33** ✅ | 有財報的 61~72 個，沒有的**全部都是 1 個** |
     | `EX-101.CAL` 存在 | 32/33 | 2024-08-12 那份有 CAL 卻沒財報（假陽性） |
     | 主文件檔名 | 未量化 | 有財報的用期末日（`arm-20260630.htm`），沒有的用申報日 |

     **R 檔是 SEC 的財報檢視器從 XBRL 三表產生的**——有三表才會有一堆 R 檔，
     沒有就只有封面那一個。這是最直接的訊號。

   - ### 兩條可行路線（都還沒做，要 CTH 決定）

     1. **預先過濾**：列 6-K 清單時看附件裡有幾個 `R*.htm`，只有 >1 的才下載。
        省掉 24 份無用下載，但每份要多一次 index 請求
     2. **讓既有機制處理**：直接把 6-K 加進抓取清單，沒有財報的那 24 份會走
        既有的「`has_financials=False` 負向快取」路徑——**第一次多下載 24 份，
        之後永遠跳過**。程式改動最小，`_save_to_disk_cache()` 本來就這樣處理
        pre-XBRL 舊申報

     ⚠ **範圍**：以上是 **ARM 一家**的觀察，下面那張表是六家的完整實測。

   - ### 六家實測（2026-09-18，`scripts/probe_foreign_issuer.py`）

     CTH 提供 ADR 分級框架後補測，**框架完全成立**：

     | 公司 | ADR 等級 | SEC 有資料 | 準則 | 20-F 解析 | 6-K 有季報 |
     |---|---|---|---|---|---|
     | NTDOY 任天堂 | **Level I** | ❌ 查無代號 | — | — | — |
     | HMC 本田 | Level II | ✅ | IFRS | ✅ 94／182／61 列 | ❌ |
     | NMR 野村 | Level II | ✅ | **us-gaap** | ✅ 88／160／85 列 | ❌ |
     | TM 豐田 | Level III | ✅ | IFRS | ✅ 56／146／52 列 | ❌ |
     | SONY | Level III | ✅ | IFRS | ✅ 124／160／77 列 | ❌ |
     | ARM | Level III | ✅ | **us-gaap** | ✅ | ✅ 9／33 份 |

     **① Level I 走 Rule 12g3-2(b) 豁免＝SEC 根本沒資料。** NTDOY 查無此代號，
     跟 IFX 同一種症狀——「查無代號」就是 Level I 的指紋，跟「有資料但格式
     不同」要分開看。

     **② ⚠ D9 說「要另建一套 IFRS 解析邏輯」是錯的。** edgartools **已經**
     解析得了 IFRS 的 20-F，TM／SONY／HMC 三表全部出得來。真正的工程量在**下游**
     ——`IS/BS/CF_TEMPLATE` 是用 us-gaap 科目名比對的，IFRS 的列名對不上。
     **所以要建的是科目對照表，不是解析器**，這兩件事的成本差很多。

     **③ 準則跟國籍無關。** 同樣是日本公司，**野村用 us-gaap、本田用 IFRS**。
     跟「註冊地猜不準」是同一個教訓——只能實測。

     **④ ARM 的 6-K 季報是特例。** 日本四家各掃 12 份 6-K，**0 份有財報**
     （R 檔都是 1）。所以「6-K 裡有季報」目前只在 ARM 身上看到。

   - ### 現行比對層對這三種來源的命中率（2026-09-18 實測，損益表 24 列）

     | 來源 | 命中 | 缺的是什麼 |
     |---|---|---|
     | **ARLO 10-Q**（對照組，正常美股） | **19/24** | D&A(CF memo)／Total Costs／Minority Interest／NI incl. NCI／SBC |
     | **ARM 6-K**（us-gaap 季報） | **17/24** | 同上 ＋ **Basic EPS／Diluted EPS** |
     | **TM 20-F**（IFRS 年報） | **10/24** | 同上 ＋ **Gross Profit／R&D／Interest Exp／Interest Inc／Other Non-op** 等核心科目 |

     **① 對照組自己也只有 19/24**——缺的那 5 列是正常的（不是每家都在損益表
     表面列 D&A、少數股東權益）。**19/24 就是「好」的狀態**，不要誤以為 24/24
     才叫成功。

     **② us-gaap 的 6-K 幾乎不用改**：17 vs 19，只差 EPS 兩列。現行
     `IS/BS/CF_TEMPLATE` 直接可用，剩下的工作是「怎麼把 6-K 納入清單」而不是
     「怎麼解析」。

     **③ IFRS 確實要另建對照表**：10/24，缺的是毛利、研發、利息收支這些核心
     科目——IFRS 的科目名跟 us-gaap 不同，不是漏報。這證實了工程量在科目對照
     表而不是解析器。

   - ### ⚠ 這對「B 路線」的影響（原本傾向的做法）

     B 路線（把 6-K 無條件加進抓取清單、靠既有負向快取過濾）**對日本公司會
     很糟**：它們的 6-K 根本沒財報，而 **Sony 有 1,055 份 6-K、野村 1,244 份**
     ——無條件加進去等於每家多下載上千份無用檔案。

     所以要做的話**判準必須是動態的**（列清單時用 R 檔數量過濾，即 A 路線），
     不能無條件把 6-K 加進清單。A 路線多的那次 index 請求，比起省下上千份
     下載完全划算。
   - **⚠ 2026-09-18 實測 218 家全掃（`scripts/check_filer_type.py`）**：實際有
     4 家抓不到，已全部移出更新名單（218 → 214 家）：

     | 代號 | 公司 | 狀況 |
     |---|---|---|
     | ARM | ARM Holdings PLC /UK | 只有 20-F（3 份），0 份 10-Q |
     | ASML | ASML Holding NV | 只有 20-F（27 份） |
     | STM | STMicroelectronics N.V. | 只有 20-F（32 份） |
     | IFX | — | **SEC 查無此代號**（Infineon 在美國是 OTC ADR `IFNNY`） |

   - ⚠⚠ **不可以靠註冊地或公司名稱判斷，兩個方向都會錯**（這是這次掃描最重要
     的教訓）：
     - NXPI（荷蘭）、LIN（英國）、ACN（愛爾蘭）、TEL（瑞士）註冊地都在國外，
       **但全部報 10-Q**——SEC 認定它們是美國國內申報人
     - 反過來 ARM 2023 才在 NASDAQ 上市、看起來像正常美股，**實際只有 20-F**。
       先前憑印象列「三家有問題」就漏掉了它
     - 決定申報身分的是股東結構與管理層所在地，不是註冊地。**只能問 SEC**
   - **加新公司進名單前先跑 `scripts/check_filer_type.py`**，比事後在 10 小時
     的抓取裡看它失敗便宜得多
