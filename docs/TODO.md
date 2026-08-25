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

B6. ~~**8-K 零下載規則的剩餘風險：`fiscal_year_end` 會不會隨時間變**~~
   ✅ **2026-08-25 驗完，風險已量化，決定接受**（CTH 決定「驗」）
   - **結果：201 家裡只有 2 家改過財年（1.0%）**，用 `scripts/check_fye_drift.py`
     離線量的（零網路請求，吃 `output/_spike/facts_*.json`）：
     | 公司 | 改制 | 改制前受影響的季 | 標錯幾季 |
     |---|---|---|---|
     | **LHX** | 2019 年從 6 月底改成 12 月底（差 177~182 天） | 30 季 | **30 季（100%）**，一律差 2 季 |
     | **MSCI** | 2010 年從 11/30 改成 12/31（差 31 天） | 1 季 | **0 季**（位移不足以跨季） |
     - 其餘 199 家：181 家最大偏移 0~4 天、18 家 5~9 天，全部是 52/53 週制的正常浮動
   - **決定：接受這個風險，不做特別處理。** 理由：① 只影響 1% 的公司，而且只影響
     **改制以前**的申報（LHX 是 2019 年以前）；② 真的要修得為每一期存一份「當時的
     FYE」，那要嘛下載歷史 10-K（正好抵銷零下載的意義），要嘛另建一張歷史 FYE 表；
     ③ `cli.py` 下載後的 `label_agrees_with_fiscal_label` 旗標抓得到選進來的那幾份
   - **⚠ 交接文件原本寫的「加一道 0~70 天 sanity check 就接得住」是恆真式**：候選季末
     永遠相隔 89~92 天，選中的必然落在 `[-tol, 91-tol)`，tol=21 時上界剛好 70，
     一次都攔不到（200 份實測 lag 範圍 -2~58）。程式碼保留它（`max_lag_days`），
     擋的是參數被改壞與畸形輸入，**不是 FYE 漂移**
   - **偵測手段仍有一個洞**：`--years` 篩在下載之前，被漂移害到而**根本沒被選中**的
     那幾份不會被下載，也就不會被比對到。抓得到「選進來的有問題」，抓不到「該選進來
     卻被漏掉」
   - **重驗方式**：`./venv/Scripts/python.exe scripts/check_fye_drift.py [門檻天數]`
     （預設 14 天）。樣本換了、或想拉更多公司進來時重跑
   - 仍未測到的：樣本只到最近 8 季（2024~2026），更早期沒驗；2004-08 之前 Item 2.02
     這個編號還不存在，那段本來就抓不到


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

## G. 2026-08-22 期間對齊／缺值／效能系列（G1／G2／G3／G6／G7／G9／G10 已完成並移入 CHANGELOG）

> **⚠ 動手前必讀：`docs/superpowers/design-2026-08-22-period-alignment-and-gaps.md`**
> 那份是 CTH 逐項確認過的**完整規格**（每項含：為什麼、規格、動哪些檔案、測試要釘什麼、
> 怎麼驗收、風險、最容易踩的坑）。下面剩下的 G 條目只是索引與決策結論，細節不重複寫，
> 兩邊有出入時**以設計書為準**。

G0. **執行順序與相依性**（2026-08-25 更新：G1／G2／G3／G6／G7／G9／G10 已完成並移入 CHANGELOG）
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

> **G11（改用 companyfacts API 取數）已於 2026-08-23 定案：不切換，主力維持
> edgartools。** 決策理由與完整報告見 `docs/CHANGELOG.md`（搜尋「G11 決策」）與
> `docs/superpowers/report-2026-08-22-g11-companyfacts.md`。下面 G10／H2 提到
> 「等 G11 決策」都已經解除，可以直接動手。

G13. **同一個期末日出現兩次（重複列）**（2026-08-22 做 G6 判定規則分析時發現）
   - 實測 SNOW 的 `Data_Financials(Q)` 有兩欄期末日都是 `2022-01-31`
   - 52 家、1,482 對相鄰期間裡只有這 1 筆，屬於低頻但真實的資料問題
   - 要查：是同一期被兩份 filing 用不同財季標籤收進來（label 沒撞號但期末日撞了），還是別的原因
   - G6 的補欄規則不會被它影響（`round(0/91)=0`），但重複欄本身會讓使用者看到兩個一樣的日期

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

H6-1. **hint 放寬後仍抓不到的案例——已診斷，待 CTH 決定**（2026-08-25，H6 主體
   已完成並移入 `docs/CHANGELOG.md`，前後數字看那條）
   - 原始資料仍在（**不要重跑**，一輪 12 分鐘）：`output/_hintsweep_201/hintsweep_201_result.txt`
     （201 家逐列逐家掃描）、`classification.md`（人工分類表，⚠ 見下方訂正）
   - 重跑指令：
     ```
     TICKERS=$(cat output/_hintsweep_201/tickers_joined.txt)
     ./venv/Scripts/python.exe scripts/diag_hintsweep.py "$TICKERS" > out.txt 2>&1
     ```
   - **⚠ 分類表有兩處是錯的（2026-08-25 查原始 10-Q 訂正）**：
     ① 「CS&APIC 那 7 家是 concept 層失守」錯——ABT/AMP/AXP/COP/KR/MPC/UNP 七家全部是
     `us-gaap_CommonStockValue(Outstanding)` / `std_concept=CommonEquity`，concept 層好好的；
     ② 「SLB 退到 fallback_suffix 層」錯——`us-gaap_Cash` 一樣有 `std_concept=CashAndMarketableSecurities`
   - **① 銀行的「Cash and due from banks」——✅ CTH 2026-08-25 決定：維持空白，
     併入 D8 一起處理。原本 7 家，同日 `label_fallback` 上線後剩 6 家
     （AXP/BK/C/COF/JPM/WFC）**
     - **BAC 已自動解決，而且拿到的是正確口徑**：它在 BS 上自己列了一條小計
       `Cash and cash equivalents` = **$229.7bn = 28.1（due from banks）+ 201.6
       （存放同業）**，正是 ASC 230 現金流量表定義的銀行現金。JPM 沒列這條小計，
       所以還是空的——**差別在公司有沒有在報表表面列小計，不是我們的規則**
     - **會計原則**：Reg S-X Article 9 規定銀行 BS 第一列是 `Cash and due from banks`
       （庫存現金、在途收款、**不生息**的存放同業）；生息的那部分（主要是存在 Fed 的
       準備金）是另一列 `Interest-bearing deposits with banks`。而 ASC 230 現金流量表
       底下 reconcile 的「現金及約當現金」，銀行實務上是**兩列相加**（有時再加
       fed funds sold）。Compustat `CHE`、Bloomberg cash & near cash 對銀行也是相加
     - **實測（JPM 2026-06-30，本專案快取的 10-Q）**：
       ```
       Cash and due from banks          $24.7bn
       Deposits with banks             $285.1bn   ← 11 倍
       Federal funds sold / resale     $446.1bn
       ```
       只抓 `due from banks` 等於填了真實現金的 **8%**——**比留空更糟**：留空使用者知道
       沒資料，填 8% 看起來像正常數字，拿去算 net debt 或流動性會整個歪掉
     - **為什麼不能直接做對**：模板一列只對到一個 XBRL 列、**不會加總**，湊不出銀行口徑
       的 cash。這正是 **D8（金融股另一套模板）** 存在的理由，等 D8 一起做才有辦法
   - **② Common Stock & APIC 剩 COP 與 MPC**：COP 的 label 只有「Par value」、MPC 是
     「Issued – 995 million and 994 million shares (par value $0.01 per share…)」，
     **措辭裡完全沒有股票字樣**，放寬措辭救不了。要救只能改成「候選只有一列
     `CommonStockValue` 時不套 hint」這種結構性規則，那會影響所有公司，風險等級跟這輪不同
   - **③ NEE 的 Capex 現在是空的（H6 之前是錯的數字）**：它唯一的 `CapitalExpenses`
     候選是「Accrued property additions」（非現金揭露），真正的 capex 在
     `nee_CapitalExpendituresOfPublicUtilitiesFPLConsolidated` 這個延伸 tag。要救屬於
     **H4 第二步**（label_fallback／延伸 tag）的範圍，而且 NEE 同一張表上還有
     `nee_CapitalExpendituresOfFPL`（子公司）與 `Other capital expenditures` 兩個相似候選，
     不能無腦用 `^capital expenditures` 當 label_fallback
   - **④ Cost of Revenue 那 29 家維持空白是正確行為**（AMT/AON/AXP/BAC/BK/BKNG/BLK/C/CCI/
     CME/COF/CSX/FDX/GS/HCA/ICE/JPM/MCD/MCO/MS/NDAQ/NSC/ODFL/PLD/SCHW/UNP/UPS/V/WFC）
     ——銀行／保險／交易所／鐵路／REIT 概念上沒有 COGS，與 **D8** 同一類。列在這裡只是備查
   - **⑤ ~~INTC 2022~2025 的 `Cash` 有 15 期抓不到~~ ✅ 2026-08-25 已修**
     （CTH 決定照公司報表表面的列示抓，見 CHANGELOG「Cash 補 label_fallback」）。
     原始症狀：INTC 那幾年把現金 tag 成
     `us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents`
     （std_concept 是 `CashAndCashEquivalents`），我們模板的 std 是
     `CashAndMarketableSecurities`、fallback 是 `CashAndCashEquivalents`——**注意
     那個 concept 名字裡沒有 "And"**（`CashCashEquivalents...`），所以兩層都比不中；
     label 明明就是「Cash and cash equivalents」，但這一列沒有 label_fallback。
     2026 年那幾份又改回 `CashAndCashEquivalentsAtCarryingValue`，所以只有中間那段空
     - **採用②（補 `label_fallback`）**：ASU 2016-18 只要求**現金流量表**的期初期末
       總額含受限現金，資產負債表沒有要求合併列示，INTC 印在 BS 上的那行字就是
       「Cash and cash equivalents」——抓公司報表表面列示的那一行是對的口徑。
       否決①（改 concept fallback）是因為它會一併吃進真的把受限現金塔得很大的公司
     - **副作用實測（201 家最新 10-Q）：新增命中 11 家、換答案 0 家**
       （BAC/CSX/DOV/EL/GILD/HSY/LULU/MMM/OMC/PG/SBUX）
   - **⑥ 更零星的個案**（材料不足，不建議單獨修）：`Other Non-current Assets` IBM
     「Investments and sundry assets」／ISRG「Long-term investments」、`Dividends Paid`
     AMT／CHTR（只有少數股權分配）、`Accounts Receivable` SO（gross 口徑）、
     `Change in Inventories` AEP、`Finance Lease Liabilities, LT` ON、`Other Current Assets` CVX


## 執行順序建議（2026-08-22 更新）

> **維護規則**：做完的條目**直接從本檔刪除**，內容搬進 `docs/CHANGELOG.md`。
> TODO 只留「還沒做的」與「還沒決定的」，不累積已完成清單，否則會無限變長、
> 接手的人分不出哪些還有效。歷史紀錄查 CHANGELOG 或 git log。

| 順位 | 項目 | 需要 API？ | 需要人？ | 說明 |
|---|---|---|---|---|
| 1 | B2 skill 端抽取 | 否 | 是（skill 設計） | **B5 與 H6 都已於 2026-08-25 完成，見 CHANGELOG** |
| 2 | G8（比較欄補洞） | — | 部分 | 風險最高，會動所有既有期間的資料來源優先序。**G10、B6、銀行 Cash 口徑、INTC Cash 都已於 2026-08-25 收掉** |
| — | ~~B2~~（已升到第 1） | 否 | 是（skill 設計） | 介面是 `cli.py press-release --json`。Non-GAAP 現在整個關閉，E2 等後續 GUI 工作卡在這條後面 |
| 4 | D8／D0-2／D0-5／D9 一次決策 | 部分否 | 是 | 都是「已知限制要不要修」的產品判斷題，建議合併成一次決策對話 |
| 5 | E 系列 GUI 細節（E2/E3/E5/E11） | 否 | 部分要 | 多半已標「先不做」或「待重現」，不影響資料正確性 |
| — | G4 overflow 標示 | 否 | 是 | 建議只做標示，不做 synonym 合併 |
| — | G8 | — | — | **G11 已定案不切換，不會白做**（G1/G2/G3/G6/G7/G9/G10 都已完成） |
| — | F2 估值倍數 | 是（要股價來源） | 是 | 待研究，未確認方向 |

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
     24.7bn vs 285.1bn，只取第一列等於填 8%。現行模板一列只對一個 XBRL 列、不會加總
     ——**「支援加總」是金融股模板的硬需求，不是選配**。理由與會計原則見 H6-1 第 ① 點
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
