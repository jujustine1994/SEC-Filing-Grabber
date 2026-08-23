# Changelog

## 現狀

- Phase 1 (GAAP)：萬能模板完成 ✅
- Phase 2 (Non-GAAP)：完成 ✅
- Phase 3 (Excel 美化)：完成 ✅
- Phase 4 (多語言)：完成 ✅

## 功能清單

### 已完成
- [x] **H4 第一步：模板 tuple 加第七欄 `label_fallback`，NVDA 的 Capex 從 39/68 期
      補到 67/68（2026-08-24）**：
      - **起因**：端到端實跑 NVDA 的 Excel 才發現的（只跑測試看不出來）。`Capex`
        年報 17 年只有 4 年有值，連帶 FY2014~FY2023 每一年的 Q4 都空、
        `Free Cash Flow` 跟著空
      - **根因**：NVDA 的 10-K 從 **FY2013 到 FY2023 共 11 年**用自己的延伸 tag
        `nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets`（逐份 filing 查證）。
        `_match_is_row()` 前兩層都比對 concept 名稱，**延伸 tag 的名字每家公司自己取**，
        比不到。它其實有第三層（label 比對），但模板的 6-tuple 餵不進去，等於死碼
      - **做法**：`_T` 加第七欄，97 列全部補上，四個呼叫點都接線。目前只有 `Capex`
        真的填了正則 `^purchases (?:of|related to).*propert`——**那是唯一有實證的案例**，
        其他列要填之前先拿資料證明。正則刻意錨在 `^purchases`，第三層後面沒有任何
        東西再擋它，寫寬就會吃到「Proceeds from sales of property」與
        「Depreciation of property」（兩個反向測試釘住）
      - **驗收**：102 家逐格比對，**模板列回歸 0 格**、補上 116 格；全套測試（含真連
        SEC）1172 passed / 7 skipped / 0 failed
      - 第二步（數值指紋自動連結）設計已定案未實作，見
        `docs/superpowers/specs/2026-08-23-concept-rename-linking-design.md`
- [x] **驗證樣本 102 → 201 家，三次量測結論一致（2026-08-24）**：再加 99 家
      （FI／HES 因 edgartools 的 ticker 對照表過期抓不到，略過）。
      | | 52 家 | 102 家 | 201 家 |
      |---|---|---|---|
      | 達標列數 | 47/97 | 46/97 | 44/97 |
      | 我們抓到／真缺口／公司真的沒有 | — | 73/9/18% | 72/9/18% |
      - 44 vs 46 是**門檻邊緣的自然浮動不是回歸**：`SBC` 87%→84%、`Diluted Shares`
        85%→80% 掉出，`Additional Paid-in Capital` 81%→88% 補進
      - 「模板不適用」9 家：AFL／AIG／AMP／AXP／BAC／COF／GS／MET 都是金融股
        （D8 已知限制），**AZO 是唯一的非金融股，原因未查**
      - **踩到並記錄的比對陷阱**：`Net Income` 與 `SBC` 在 IS 和 CF 模板各有一列，
        比對腳本用列名當字典鍵只會留最後一個，變成拿 IS 那列比 CF 那列，憑空生出
        **3,659 個假異動**、一度誤判成嚴重回歸。改用 `(列名, 第幾次出現)` 當鍵之後
        真實回歸是 0。overflow 區的 `Other`／`Accrued expenses` 這類重複列名更嚴重
- [x] **驗證樣本 52 → 102 家，模板覆蓋率不變（2026-08-23）**：新增 50 家跨產業
      公司（醫療 LLY/ABBV/MRK/AMGN…、工業 HON/DE/RTX/BA…、消費 PEP/HD/TGT…、
      金融 V/MA/AXP/MS/C/WFC、能源 COP/SLB/EOG、公用事業 DUK/SO、REIT AMT/EQIX、
      EDA SNPS/CDNS）。窗口統一 `max_filings=16`，50 家全部成功、0 失敗。
      - **達標列數 47/97（52 家）→ 46/97（102 家）**，樣本翻倍且產業分布大幅
        擴散之後幾乎不變 → 模板不是只對半導體／軟體有效
      - 「XBRL 裡到底有沒有」三分類也幾乎不動：**我們抓到 73%／真缺口 9%／
        公司真的沒有 18%**，52 家與 102 家兩次量測完全相同的比例
      - 「模板不適用」從 4 家（BAC/GS/SCHW/PLD）變 3 家（BAC/AXP/GS）——SCHW 與
        PLD 因為 H3 的修復脫離了門檻，新進的 AXP 遞補。金融股仍是已知限制（D8）
      - **踩到並修掉一個量測陷阱**：覆蓋門檻原本寫死「>=45 家」，那是照 52 家訂的
        （≈87%）。樣本擴到 102 家時同一個「45」等於門檻鬆到 44%，達標列數會從
        47 假性跳到 **74**。改成 `MIN_CO_RATIO = 0.85` 隨樣本數走，並把這個坑寫進
        常數旁邊的註解——**這正是文件第零節在防的那種誤導**
- [x] **H3-2：「中間有洞」拆出「零星有值」，假警報砍掉約一半（2026-08-23）**：
      B 判斷（首末有值之間不能有空格）對**流量列**會過度報警——公司沒併購、
      沒借款、沒買回的那一季，通常是整條列不寫，不是寫 0。
      - **門檻是量出來的，不是拍的**：拿 companyfacts 當真值驗 52 家、2,906 個洞，
        逐洞問「這個期末日在 companyfacts 裡到底有沒有數字」：
        | 首末之間填滿率 | 洞數 | 真的漏抓 |
        |---|---|---|
        | 0~30% | 269 | **8%** |
        | 30~50% | 575 | **18%** |
        | 50~70% | 682 | **22%** |
        | 70~85% | 1,179 | 54% |
        | 85~100% | 201 | 46% |
      - 70% 是乾淨的分水嶺 → `_SPORADIC_FILL_RATIO = 0.7`。低於門檻的列改歸
        新的 `SporadicRow`（零星有值），**Index 上照樣列出來但不用 ⚠、不算進問題數**
      - **不照列名硬編一張「流量列」名單**：同一列在不同公司的性質不一樣
        （AAPL 幾乎每季都有併購支出，ONTO 三年一次）。改用**這家公司自己在這一列
        的填滿率**判斷，跟 C 判斷「不跟別家比」的原則一致
      - `_SPORADIC_MIN_SPAN = 8`：期數太少看不出型態（3 期缺 1 期就是 67%，
        但那只代表資料太少），低於 8 期不降級
      - 效果：`Acquisitions` 標紅 25 家 → **13 家**、`Debt Proceeds` 24 → **11**、
        `Debt Repayments` 16 → **12**
- [x] **H3-1：只在附註、沒印在報表表面的科目 → 暫列已知限制 D10（CTH 2026-08-23）**：
      `Op. Lease Liabilities, current` 實測 AMD／AMZN／ARLO／COST／GOOGL／MSFT／MU／
      NVDA／ORCL／PANW／SWKS／TSLA **十二家全部**只在資產負債表表面列非流動那條，
      流動部分併在「其他流動負債」裡。**改 concept 名字救不了**。
      **標「暫時性」是有意義的**——`fetcher_facts`（companyfacts）讀得到這些
      （該列 44/52），資料拿得到、只是現行路徑碰不到，隨時可以接
- [x] **H3：模板 concept 對照的系統性修復，覆蓋率 40/97 → 47/97（2026-08-23）**：
      52 家實測基線 `docs/template-coverage-baseline-2026-08-23.md` 重跑後，
      **16 個模板列的有值公司數上升、只有 1 列下降 1 家（刻意的，見下）**。
      - **最大的一批「矛盾」其實是判斷規則誤判，不是抓漏**：
        `Current Portion of LT Debt` 原本 25/52 家被判矛盾，一度以為是 concept
        名字錯。實際查 INTC／PG／XOM／MU／NVDA／QCOM／PFE／ORCL 八家的資產負債表，
        **表面只有一條流動借款列**（tag 成 `us-gaap:DebtCurrent` 或
        `NotesPayableCurrent`，`standard_concept` 是 `ShortTermDebt`），一年內到期
        的長期負債就併在那一條裡——而那一條我們早就抓進 `Short-term Debt` 了，
        資訊沒有掉。`data_quality` 新增 `_COHERENCE_EXCUSES`（「空白有正當理由」
        的反證，優先於 `_COHERENCE`），**矛盾 25 家 → 3 家**。留下的是 AMZN 這種
        表面連流動借款列都沒有的公司，那才是真的抓不到
      - **`label_hint` 太窄是最大的系統性成因**。掃描 22 家最新 10-Q，統計「hint
        把一個本來命中的列變成沒命中」的次數：`Cash Taxes Paid` 14 家、
        `Deferred Revenue, current` 12、`Share Repurchases` 9、
        `Accounts Receivable` 8、`Capex` 7、`Cash Interest Paid` 6、
        `Long-term Debt` 5、`Change in Inventories` 4。
        `_pick()` 濾空之後**整個優先層被跳過**，不會退回去用 concept 比對，
        所以一個過窄的 hint 等於把那一列整個關掉。典型例子：多數公司寫
        「Treasury stock purchases」不寫「repurchase」；PG 三個現金流小計都寫
        「TOTAL OPERATING ACTIVITIES」；hint 寫複數 `inventories`，但公司寫單數
        `Inventory`
      - **`Shares Outstanding` 43/52 家「中間有洞」，成因單一**：期末股數走封面頁
        的 `dei:EntityCommonStockSharesOutstanding`，**10-K 那筆標的是 `fp='FY'`**，
        對出來的標籤是 `FY2025`，季表要的是 `FY2025Q4` → **每一年的 Q4 都是洞**。
        `_shares_map_from_records()` 現在把年度那筆同時補進 Q4（公司自己有標
        `fp='Q4'` 就不覆蓋）。**有洞 43 家 → 5 家**
      - **`Other Operating Expense` 52 家掛零**：模板猜的 `OtherOperatingExpenses`
        （std）與 `OtherOperatingExpense`（fallback）**沒有任何一家公司在用**。
        實際 tag 的是 `OtherCostAndExpenseOperating`（KO）與
        `OtherOperatingIncomeExpenseNet`（MCD／CAT）。**0 家 → 13 家**
      - **`Debt Proceeds` / `Debt Repayments` 改由既有的加總後處理獨佔**：模板比對
        到一條之後 `consumed` 會讓加總跳過它，加總再覆蓋回去就只剩「其他幾條的
        和」——所以這兩列的模板 tuple 改成不比對，值一律由 `_sum_matching_rows()`
        算。同時**排掉 `ProceedsFromRepaymentsOf...` 這種淨額**（借款減還款，可正
        可負）：WMT 原本「還款」是 **+26 億正數**、PG 少算 10 億、MCD 少算 4.9 億。
        唯一下降的一列就是這個 trade-off——GE 只有一條淨額列，排掉之後那一格空白
        （落到 overflow，資料沒消失）
      - 其餘一併修好的列：`Accounts Receivable`（hint `accounts receivable` →
        `receivable`）、`Cash`（`cash and (cash )?equivalents`，MCD 寫
        「Cash and equivalents」）、`Other Current Assets`、`Other Non-current Assets`
        （MCD 寫「Miscellaneous」）、`Common Stock & APIC`（GOOGL 寫「Class A,
        Class B, and Class C stock and additional paid-in capital」）、
        `Long-term Debt`（fallback 改成排除流動部分的正則，AAPL 寫「Term debt」、
        CRM 寫「Noncurrent debt」）、三個現金流小計、`Capex`、
        `Cash Taxes Paid`／`Cash Interest Paid`（原本的 std_concept 分別會命中
        **遞延所得稅費用**與 **AVGO 的債務清償損失**，那都不是付出去的現金）
      - **新增 53 個測試**，每一個都釘住「哪一家公司、哪個 concept、哪個 label」，
        字串是從真實 `to_dataframe()` 逐字抄的。含反向測試（放寬 hint 之後不能
        誤抓：TSLA 的流動負債不能進長期負債、XOM 的利息要取合計不是營運那段、
        AAPL 的 Non-trade receivables 不能冒充其他流動資產）
      - **28 家真實 10-Q 前後逐格對照：新命中 82 格、失去 1 格**（GE 的淨額列）
- [x] **G5：Index 的「完成度」重寫成四個判斷（2026-08-22，CTH 回報「很多東西
      沒出現也是滿分，而且這標準不是我設的」）**：
      - **舊標準形同虛設**：只看 9 個關鍵列、而且**只在「最近 4 期全部是 None」
        才算缺**。同一份 NVDA 檔案舊版顯示 `9/9 ✓`，實際上 95 個欄位裡有 27 個
        幾乎全空、3 季完全沒有資料
      - **新增 `src/data_quality.py`（26 個測試）**，四個判斷，可信度由高到低，
        **全部不需要「同業基準表」**：
        | | 判斷 | 誤判率 |
        |---|---|---|
        | A | 季度斷層 `round(天數差/91)-1` | 0 |
        | D | 整欄稀疏（欄位在、整排幾乎都空） | 0 |
        | B | 中間有洞（同一列有些期有、有些沒有） | 0 |
        | C | 整列全空且與相關欄位矛盾 | 低 |
      - **CTH 直接點出的誤判風險**（「基準表會不會剛好某間公司沒有某項目、
        結果永遠標紅」）就是不採用同業普及率的原因。改用**同一家公司的相關
        欄位互相驗證**：NVDA 有 74 億長期負債、有利息費用，卻沒有任何借還款
        現金流——這家公司自己的資料就對不起來。反過來若負債類欄位全空，
        沒有借還款紀錄是一致的，不標紅
      - **三個實測踩出來的細節**（都寫進程式註解）：
        ① A 不能用固定門檻——52 家 1,482 對相鄰期間裡，111~150 天的 16 筆
        **全部是 COSTCO**（16 週的第四季），固定門檻會全部誤判；
        ② B 只看「首末有值之間」——`Operating Lease ROU Assets` 只有 28/67 期
        是因為 ASC 842 從 2019 才適用，不是漏抓；
        ③ B 要排除 overflow 列與整欄稀疏的期間——不排除的話 NVDA 報 85 列有洞，
        排除後 14 列
      - **自動判定「模板不適用」**：稀疏欄超過一半期數時直接這樣講。52 家實測
        BAC/GS/SCHW/PLD 全數觸發（稀疏欄佔 90~100%），因為 IS/BS/CF 模板是為
        製造業設計的（TODO D8 已記錄）。對這種情況列 21 行稀疏欄沒有意義
      - Index 的完成度欄與明細區塊整個換掉，12 個新 i18n key × 4 語系
- [x] **模板體檢基線文件（2026-08-22）**：新增
      `scripts/gen_template_coverage_baseline.py`，產出
      `docs/template-coverage-baseline-<日期>.md`——52 家的逐列覆蓋率、
      最常出問題的列、現行路徑 vs companyfacts 的差距。不打網路，吃
      `output/_spike/` 的快取。**現行路徑達到「≥45 家有值且填滿率 >90%」的列
      只有 40/97**，這是之後改 concept 對照時要往上推的數字
- [x] **G12：現金流量表的「期末現金餘額」被當成期間值做減法（2026-08-22）**：
      52 家裡 50 家的 `Ending Cash` 是錯的。實測 AAPL 顯示 2.55 億甚至負數，
      實際是 450 億。
      - **同一個 bug 有兩個實例**，修完第一個才發現第二個：
        ① `_build_cf_table()` 的 YTD 拆算把餘額減成「現金變動額」；
        ② `_synthesize_q4()` 的 Q4 合成把餘額做「年報 − Q1 − Q2 − Q3」減成負數
      - **修法**：`_CF_POINT_IN_TIME_IDX` 標出流量表裡混著的時點值列，兩處都跳過
        相減。`_synthesize_q4()` 加 `point_in_time_idx` 參數（整張表
        `is_balance=False` 時，個別餘額列仍要直接取年報值）
      - **驗收（真實 AAPL）**：`2025-09-27` 從 −58,796,000,000 修成
        35,934,000,000、`2026-03-28` 從 255,000,000 修成 45,572,000,000，
        六期全部正確且 Q4 對得上 Apple FY2025 年末現金
      - 日後在 `CF_TEMPLATE` 加時點值列時要記得加進 `_CF_POINT_IN_TIME_IDX`
      - 4 個新測試（含「修時點值不可以順手把流量項的 YTD 拆算也關掉」）
- [x] **G9：一次抓取內同一份 filing 只解析一次，加速 1.4~1.9 倍（2026-08-22，
      CTH 問「為何跑這麼久」）**：
      - **實測定位**：ARLO 25 份 filing 共 66 秒，`_filing_obj` 被呼叫 96 次
        （每份 3.8 次），其中 `financials`（XBRL 解析）19.9s、`to_dataframe` 28.4s。
        IS/BS/CF/segments 四個 build pass 各自對同一批 filing 重解析一次
      - **edgartools 不會跨呼叫快取**：同一支 ticker 在同一個 process 連跑兩次是
        64.5s vs 67.3s，完全沒變快——所以這個重複成本是真的
      - **修法**：`_parse_cache_scope()` 涵蓋一次抓取，`_filing_obj()` 與
        `_financials_of()` 在範圍內快取。**快取的生命週期只能是一次抓取**——
        跨 ticker 殘留會吃到別家資料，跨執行殘留會拿到過期申報
      - **放在哪很重要**：包在 `fetch_gaap_statements()` 外層（本體拆成
        `_fetch_gaap_impl()`），不是放在 `_ledger() is None` 那個分支裡。
        `main.py`／`cli.py` 會自己先開 `collect_gaps()`，那條路不會走到遞迴
      - **實測**：64.5s → 33.6s（冷）／44.0s → 32.4s（全熱）。全套測試也從
        21 分鐘降到 10 分 44 秒
      - **輸出完全沒變**：同一支 ticker 開／關快取各跑一次，**5,678 格逐格比對
        0 格不同**
      - 3 個新測試（含「不可跨執行殘留」與「沒開範圍也要能跑」）
- [x] **G11 spike：SEC companyfacts 取數路徑建好並驗證（2026-08-22，CTH 指示
      「完整跑 G11，讓資訊來源更穩定一點」）**：現行路徑對每份 filing 解析 4~7 次
      （IS/BS/CF/segments 各一次），實測每家 7.5 分鐘；companyfacts 一家一個
      request，**0.34 秒，快 215 倍**，而且涵蓋範圍還更早（NVDA 2008-07 vs 2009-07）。
      - **新增 `src/fetcher_facts.py`（40 個測試）與 `src/facts_mapping.py`**。
        現有程式**一行都沒動**，這是平行路徑
      - **對照表是反推出來的，不是手填**：模板的 `std_concept` 欄是 edgartools
        正規化過的名字（`NetIncome`），不是原始 us-gaap element name
        （`NetIncomeLoss`）。拿現行路徑已知正確的數字當答案卷，對 52 家 ×
        每個 concept 算命中率，最高的就是正確對應。95 列中 90 列自動對上、
        3 列人工補、2 列合理沒有對應。每列的證據寫在行末註解
      - **驗收（52 家逐格）**：61,069 格中 56,686 格數字相同（**92.82%**），
        其中 1,543 格只差正負號 → **符號對齊後 95.35%**
      - **獨立驗證（24 家，不看現行路徑）**：資產=負債+權益 95%、四季加總=年度 95%、
        **SEC 官方 `frame` vs 我們的期中點判準 59,564/59,564 零例外全數一致**
        （這是對 F6 跨公司對齊決策的官方背書）、實質重編佔 5.4%
      - **四個踩到的坑**（都寫進程式註解）：① 一個 concept 只能當一列的主人
        （`OtherNonoperatingIncomeExpense` 差點變成 `Operating Income` 的備援）；
        ② `negate` 要跟隨首選不能用 `all()`（否則 31 家的 Capex 全部正負相反）；
        ③ 備援的符號慣例必須與首選一致否則要剔除；④ `unit`/`taxonomy` 指錯一律
        回空不可退回預設（EPS 是 `USD/shares`、流通股數在 `dei`，退回 USD 會讓
        每股盈餘抓到金額且看起來很合理）
      - **尚未切換**。完整報告與待決事項見
        `docs/superpowers/report-2026-08-22-g11-companyfacts.md`
      - 四支 spike 腳本登記於 `scripts/README.md`
- [x] **G1：日曆季判準統一成「一份實作、兩個具名基準點」（2026-08-22）**：同一件事
      原本有三套算法散在三個檔案——① `fetcher_gaap._calendar_quarter()` 完全不內縮
      （把 INTC 結束在 2023-04-01、實際涵蓋 1~3 月的那一季算成 `2023Q2`）；
      ② `fiscal_input` 的 −15 天；③ 跨公司用的 −45 天期中點。
      - **不是併成一個數字**：「這季結束在哪一季」與「這季主要落在哪一季」是兩個
        問題，全改 −15 會讓跨公司對齊偏一格，全改 −45 會讓第 4 列對一個 7 月結束
        的季說「2025Q2」、跟正下方第 5 列的 `2025-07-27` 自相矛盾
      - **改法**：①直接刪掉（改成委派）；②③ 合併成
        `calendar_quarter_of(period_end, *, basis)`，`basis="end"`／`"span"`。
        **`basis` 刻意不給預設值**，不給就噴 `TypeError`——強迫每個呼叫端表態，
        這才是「統一」的實質保障
      - `calendarized_quarter_of()` 保留成 `basis="span"` 的薄包裝，讓
        `comparison.py` 讀起來是「取對齊季」而不是「傳個字串參數」；
        測試釘住薄包裝與 `span` 逐格等價，不可各自演化
      - 機器鍵 `Calendar Quarter` **沒有動**（改了會打斷下游腳本，見 TODO D0-5），
        只改四個 locale 的 B 欄說明，寫明「該期**結束**在哪個日曆季（跨公司比較
        用的是另一套判準）」
      - 12 個新測試於 `test_fiscal_input.py` / `test_fetcher_gaap.py`
- [x] **G3：IS 區的 SBC 與 D&A 只有 Q1 有值，改用 CF 表已拆算好的單季值（2026-08-22，
      CTH 回報「有些財務項目有時出現有時不出現」）**：
      - **根因**：`IS_TEMPLATE` 裡 `source == "CF"` 的兩列走
        `_current_q_col(cf_df)` 直接找現金流量表的單季欄，但 **10-Q 的現金流量表是
        YTD 累計**，Q2/Q3 的 filing 根本沒有單季欄 → 整片空白。實測 NVDA 缺 51/68，
        缺的全是 Q2/Q3（Q1 有值，因為 Q1 的 YTD 就等於單季），連帶
        `_synthesize_q4()` 也算不出那兩列的 Q4
      - `_build_cf_table()` 早就做了 YTD 拆算（本季 YTD − 上季 YTD），所以 CF 區
        同名兩列一直是好的（缺 1/68）——**問題只在 IS 這條路沒走那段**
      - **改法**：新增 `_backfill_cf_sourced_rows()`，用 CF 表依 `quarter_labels`
        回填 IS 的 `SBC` 與 `D&A (CF memo)`。**不在 IS 再寫一份 YTD 拆算**——
        那就是第二份會漂移的實作
      - **一定要放在 `_synthesize_q4()` 之後**：CF 表的 Q4 是那一步才補上的
      - 保留 IS 原本讀 `cf_df` 那段（提供 B 欄的公司自報標籤），回填只在 CF 有值
        時覆蓋——比設計書原本寫的「整段拿掉」嚴格更好
      - **驗收（真實 NVDA 資料）**：IS `SBC`／`D&A (CF memo)` 從缺 51/68 降到 1/68，
        與 CF 區的 `SBC`／`D&A` 完全一致。剩下那 1 期是最早的合成 Q4，材料是
        pre-XBRL 申報
      - 4 個新測試（依 label 對照不靠位置、CF 是 None 不蓋掉 IS 已有的值、
        缺列不拋例外）
- [x] **圖表三個標題壓在別的東西上面（2026-08-22，F3 第三輪）**：CTH 截圖回報
      Y 軸標題「Revenue ($mm)」壓在「50,000.0」刻度數字上、X 軸標題「期間」
      掉進日期標籤那一排裡面。**跟前一輪圖例那個 bug 完全同一類**：openpyxl 的
      `Title` 沒有 `overlay` / `layout` 屬性時完全不寫這兩個元素，Excel 拿到
      「沒寫」的標題會直接畫在既有內容上面，而不是另外撥一條專屬空間。
      實測存檔後的 `xl/charts/chart1.xml`：`<legend>` 有
      `<overlay val="0"/>`（上一輪補的），三個 `<title>` 一個都沒有，也都沒有
      `<layout/>`。原生 Excel 輸出的每個標題一定帶 `<c:layout/>`（空元素＝
      明講用自動版面）加 `<c:overlay val="0"/>`。
      - **修法**：新增 `_pin_title_layout()`，對 `chart.title`／`x_axis.title`／
        `y_axis.title` 三個都設 `overlay = False` + `layout = Layout()`。
        **必須放在三個 title 都指派完之後**——openpyxl 每次指派 `title` 都會
        重新造一個 `Title` 物件，先設屬性會被後面的指派蓋掉
      - **驗證**：PowerShell 呼叫 Excel COM 把同一組資料的圖表匯出成 PNG
        前後對照。修復前 Y 軸標題確實疊在 50,000.0 上、「期間」浮在繪圖區
        中間偏下（20171029 附近）、圖表標題壓進頂端格線；修復後三個標題
        各自佔一條專屬空間，都不重疊。附帶把「圖表標題壓進繪圖區」這個
        沒被回報但同樣存在的問題一起修掉
      - 2 個新測試於 `test_comparison_writer.py`，其中一個直接解開存檔後的
        `.xlsx` 讀 `xl/charts/chart1.xml`，數 `<overlay val="0"/>` 有 4 個
        （三標題 + 圖例）、`<layout/>` 有 3 個——**Excel 讀的是檔案不是
        Python 物件，只斷言物件屬性不算驗證**
- [x] **財季編號改由期末日反推，修掉「季度資料被靜默丟棄」與連鎖的 Q4 缺洞
      （2026-08-22，D0-6）**：`_col_to_quarter_label()`（`fetcher_gaap.py:431`）
      原本直接採信 edgartools 欄名裡的 `(Qn)` 標記。實測那個標記對 52/53 週
      財年制的公司會標錯，而且**同一家公司相鄰兩季會標到同一個號碼**：
      ```
      NVDA 2010-05-02 → (Q2)   實際 FY2011Q1
      NVDA 2010-08-01 → (Q3)   實際 FY2011Q2   ← 跟下一行撞號
      NVDA 2010-10-31 → (Q3)   實際 FY2011Q3
      INTC 2023-04-01 → (Q2)   實際 2023Q1
      INTC 2023-07-01 → (Q3)   實際 2023Q2     ← 跟下一行撞號
      INTC 2023-09-30 → (Q3)   實際 2023Q3
      ```
      filings 是新到舊排序，`_build_*_table()` 的 dedup（`if label in periods:
      continue`）於是把**舊的那一季靜默丟掉**，沒有任何 log 或警告。少了 Q1，
      `_synthesize_q4()` 的「Q1/Q2/Q3 都要有」判斷式自然跳過該年 Q4 合成——
      這才是「Q4 常常抓不到」的真正根因，不是 SEC 端缺資料（實測那幾份
      10-Q 的 `filing.xbrl()` 都拿得到）。
      - **更嚴重的是沒被丟掉的那些也貼錯格**：`Compare_Data` 的
        `INTC / FY2016Q2 = 13,702` 其實是 INTC **2016 Q1** 的營收，整條線
        往後平移一季，而同一欄的 AMD 是真正的 Q1——兩家在圖上根本沒對齊
      - **修法**：`(Qn)` 現在只用來分辨「季度欄 vs 年度欄」，編號一律用
        `_col_to_period_end()` 解析出的實際日期，交給
        `fiscal_input.fiscal_quarter_of()`（既有的參考實作，先把期末日內縮
        15 天再取月份，剛好吃掉 52/53 週制的漂移）反推。沒有另外發明仲裁
        規則——edgartools 的標記不再參與計算
      - 4 個新測試於 `tests/test_fetcher_gaap.py`（NVDA/INTC 實際踩到的欄名，
        以及「同一財年三季必須算出三個不同 label」的碰撞回歸測試）。兩個既有
        測試的期望值跟著改：它們的 fixture 是 `2025-12-27 (Q1)` 配預設
        `fy_end_month=12`，那個組合在現實不存在（12 月結算的公司 Q1 不會在
        12 月底結束），改成 `FY2025Q4`
- [x] **跨公司比較改用日曆季／日曆年對齊（2026-08-22）**：`comparison.py`
      原本直接拿各公司自己的財季標籤當共同欄位鍵。財年結束月不同的公司，
      同一個標籤指的不是同一段時間——NVDA 一月結算，它的 FY2026Q2 實際結束
      在 2025-07-27，卻跟 AMD 結束在 2026-06-27 的 FY2026Q2 疊在同一欄，
      **整條 NVDA 線在日曆時間上偏移約一年**。
      - **對齊判準用期中點**（期末日往前推 45 天，即「這一期的多數天數落在
        哪個日曆季」），不是期末日落在哪一季。理由：NVDA 7 月底那一季要跟
        AMD/INTC 6 月底那一季擺同一欄（同一波財報、分析師就是這樣比），
        用期末日會把它推到跟 AMD 9 月那一季同欄，錯一格
      - 年度模式同理：財年結束在 1-5 月掛前一個日曆年（NVDA FY2026 結束在
        2026-01，內容是日曆 2025 年），6-12 月掛當年（MSFT FY2025 就是 2025）
      - 新增 `fiscal_input.calendarized_quarter_of()` / `calendarized_year_of()`，
        **刻意不動既有的 `calendar_quarter_of()`**——後者是單一公司
        `Data_Financials(Q)` 第 4 列在用的，語意是「期末日落在哪一季」，
        跟跨公司對齊是兩件事，共用一個函式以後改哪邊都會誤傷另一邊
      - `Compare_Data` 的期末結算日列改取同欄各公司**最晚**的那個日期
        （同一日曆季各家期末日不再相同）：Snapshot 拿它做「不晚於 B1」的
        判斷，取早的那個會讓使用者把 B1 設在 7/1 就看到 NVDA 還沒結算完的
        那一季數字
      - 期末日抓不到時退回原本的財季標籤，不因為算不出日曆季就整季丟掉
      - 11 個新測試（`test_fiscal_input.py` 8、`test_comparison.py` 3、
        `test_comparison_writer.py` 1），5 個既有 `test_comparison.py` 測試的
        期間鍵跟著從 `FY2024Q1` 改成 `2024Q1`
- [x] **跨公司比較 `Chart_<指標>` 圖表三個隱藏渲染 bug（2026-08-22，CTH
      截圖回報「中間斷線＋圖例被吃＋沒有單位」，F3 完成後才發現的殘留問題）**：
      前一版 F3 只調了尺寸/圖例位置/numFmt 這些「看得到的屬性」，實際存進
      Excel 檔案的圖表卻完全沒有座標軸文字（Y 軸數字、X 軸日期都不見），
      圖例整條疊在 X 軸標籤上——**這三個 bug 光看 openpyxl 文件或程式碼猜
      不到，是因為本機剛好有裝 Office，改用 PowerShell 呼叫 Excel COM
      自動化，把「openpyxl 產出的圖表」跟「Excel 原生 `ChartObjects.Add()`
      建出來的圖表」兩份 XML 逐項比對，才抓出三處 openpyxl 完全不寫、但
      Excel 原生輸出一定會寫的屬性**：
      1. `chart.set_categories()` 不管儲存格內容永遠寫成 `<c:numRef>`，但
         期末結算日是文字（"20240331"），Excel 拿數值參照指向文字儲存格
         解析不出來，類別軸整個讀不到值。改成手動把每個 series 的
         `cat` 換成 `AxDataSource(strRef=StrRef(f=...))`
      2. openpyxl 完全不寫 `<c:delete>` 元素（regardless大小寫都沒有）。
         OOXML 規格上沒寫預設是 `false`，但 Excel 實際渲染時對「沒寫」跟
         「明講 `delete=0`」待遇不同——沒寫會保守地整排不畫刻度標籤。兩軸
         都要 `chart.x_axis.delete = False`／`chart.y_axis.delete = False`
      3. `Legend` 沒有 `overlay` 屬性時完全不寫該欄位，Excel 拿到「沒寫」
         會讓圖例跟 X 軸標題/刻度標籤擠在同一條窄帶、直接疊在一起（畫面上
         看起來像圖例文字蓋在日期上）。原生 Excel 輸出一定帶
         `overlay="0"`，這裡明講 `chart.legend.overlay = False`
      - 同時補上 `axPos`（catAx 該是 `"b"` 不是 openpyxl 預設的 `"l"`）、
        `tickLblPos="nextTo"`、`crosses="autoZero"`，都是原生 Excel 輸出
        必有、openpyxl 不會主動寫的欄位，一次補齊不留模糊地帶
      - **驗證方式**：`PowerShell` 呼叫 Excel COM，用真實 Excel（不是
        LibreOffice 或其他工具）開啟並匯出圖表成 PNG 圖片，肉眼比對前後
        差異；並用二分法（一次拿掉一個自訂屬性）排除誤判，找到真正觸發
        問題的屬性。最後用 INTC/NVDA/AMD 真實資料重新產出，確認 Y 軸數字
        （0 到 90,000.0，帶 `$mm`）、X 軸日期（跳 4 期顯示一次）、圖例
        （AMD/INTC/NVDA 三色分開不重疊）都正常顯示。新增 3 個測試於
        `test_comparison_writer.py`
- [x] **跨公司比較 Snapshot 改用真日期輸入 + 最近一期查找（2026-08-21，CTH
      回報「打數字抓不到資料」並要求日期不用剛好對到期末結算日）**：
      `comparison_writer.py::write_snapshot_sheets()` 的 B1 原本是純文字，
      要求剛好打中 `YYYYMMDD`——使用者打數字會被 Excel 自動轉成數值型別，
      跟 Compare_Data 期末結算日列（文字）型別對不上，`MATCH` 抓不到值。
      改成真正的 Excel 日期型別（`number_format = "yyyy/mm/dd"`），使用者
      可以直接打一般日期（如 `2024/7/15`），不用湊 8 碼數字。同時把查找
      邏輯從「精確比對」改成「不晚於這天的最近一期」（`SUMPRODUCT(MAX(...))`
      公式），符合分析師「這個時間點看得到的最新數字是什麼」的直覺——未來
      才公布的財報不會被拿來回填過去的時間點。開發過程中用 PowerShell
      呼叫真實 Excel COM 自動化實測（本機沒有 LibreOffice，改用已安裝的
      Office），抓到並修掉兩個公式在 Python 端寫測試測不出來的 Excel 執行期
      bug：① 用「期末結算日聯集列」判斷儲存格是否空白會誤判（應該看該公司
      自己那一列）；② `INDEX(range,0)` 在 Excel 不會報錯而是回傳整個範圍的
      隱含第一格，`IFERROR` 接不住，要用 `IF` 明確擋掉 offset=0 的情況。
      新增測試於 `test_comparison_writer.py`
- [x] **跨公司比較 X 軸資料密度控管 + 圖表版面調校（2026-08-21，TODO F3 全部
      5 項完成，含新發現的第 6 項）**：`comparison_writer.py::write_chart_sheets()`
      補齊尺寸（`chart.width=30`／`height=15`，openpyxl 預設 15×7.5cm 偏小，
      CTH 要求雙倍）、圖例移到下方橫排（`legend.position="b"`，CTH 確認選
      這個而非右側——公司數量不固定，右側直排公司一多會被擠進繪圖區，下方
      橫排可隨數量自動換行）、Y 軸數字格式改跟 `Compare_Data` 儲存格同一套
      規則（`excel_formatter.unit_format_for()`），軸標題金額類指標加
      `($mm)` 單位。另外修了一個原本 5 項清單沒涵蓋、這次修復自己帶出來的
      新問題（第 6 項）：接上 D0-1 Q4 合成後跨公司比較的時間跨度可以拉到
      60-70 欄，全部日期標籤硬擠在 X 軸會疊字看不清楚，加
      `x_axis.tickLblSkip`／`tickMarkSkip = max(1, n_periods // 15)`，目標
      約 15 個可視標籤。真實抓 INTC/NVDA/AMD 資料驗證：X 軸 69 期正確跳 4
      期顯示一次。新增 5 個測試於 `test_comparison_writer.py`
- [x] **跨公司比較「選擇比較內容」視窗 3 項 UI 問題（2026-08-21，TODO F4）**：
      `main.py::_open_compare_selection_window()`。① 新增模組層級純函式
      `match_company_cache()`：公司自動完成原本只比對 ticker 開頭字串，打
      公司全名（如「Intel」）搜不到，改成 ticker 開頭**或**公司名稱含有輸入
      字串符合其一即收錄，8 個測試於 `test_match_company_cache.py`。②③
      新增 `pack_wrapped_chips()`（已選公司/指標的 chip 依容器寬度自動換行，
      不再用單純 `pack(side="left")` 導致超寬被視窗邊緣裁掉看不到）與
      `make_scrollable_frame()`（整個視窗內容包進可捲動 Canvas，確認/取消
      按鈕改成先 pack 在 `win`／`side="bottom"`，永遠貼在視窗底部看得到，
      不會被下面可捲動內容擠出視窗）。這兩個都是 GUI 版面調整，沒有無頭
      測試框架可用，改用真實 Tk（Windows 本機有顯示，`Tk()` 建得起來）寫
      驗證腳本：塞 20 家公司/20 個指標的長標籤進選擇視窗，確認 chip 換行
      成多行、確認/取消按鈕維持在視窗可視高度內（y=605，視窗高 640）、
      而被擠出的內容確實在可捲動區域內（y 一路到 1163）而非消失或報錯
- [x] **跨公司比較季度模式收尾補一次進度到滿格（2026-08-21，TODO F5，CTH
      回報「Excel 已產出」訊息在進度條還沒跑完時就先跳出來）**：
      `main.py::_compare_worker()` 全程只靠 `_on_company_start()`（每次開始
      抓下一家時把進度設成 `current-1`，最後一家開始抓時進度只到
      `total-1`，設計上就不會自然到滿格）與 `_on_filing_progress()`（借用
      同一條進度條顯示逐份 filing 進度，但最後幾家一開抓就失敗時連一次都
      沒機會回報）更新進度，寫完檔案前後沒有任何地方補一次「全部做完」的
      進度更新。單一公司抓取與批次抓取都有在收尾呼叫
      `self._set_progress(total, total, ...)`，跨公司比較這裡漏了同樣的
      收尾呼叫。補上後跟另外兩條路徑行為一致
- [x] **跨公司比較的季度模式接上 D0-1 Q4 合成（2026-08-21，CTH 回報圖表中段大量缺
      Q4）**：`comparison.py::build_comparison()` 原本把「顯示頻率」跟「該不該抓
      年報」錯誤綁成互斥——`fetch_annual=(frequency == "annual")`，選季度比較時
      年報完全不抓，D0-1 的 Q4 合成（`fetcher_gaap._synthesize_q4()`，靠「年報－
      Q1－Q2－Q3」湊出 10-Q 沒有的 Q4）沒有材料可用，整個跳過，日曆年結束的公司
      （AMD/INTC 等）日曆 Q4 那一整欄空白。單一公司 Tab1 抓取用兩個獨立勾選框
      （`fetch_q`／`fetch_k` 預設都 `True`）沒有這問題，是跨公司比較這邊漏接。
      修法：`fetch_annual` 固定傳 `True`，年報只當 Q4 合成材料用，不影響輸出頻率
      （季度/年度模式的 `fetch_quarterly` 仍照原邏輯區分）。代價是每家公司多抓
      年報份數，跨公司比較耗時會拉長。新增測試
      `test_build_comparison_quarterly_frequency_still_fetches_annual_for_q4_synthesis`
      於 `test_comparison.py`，全套測試皆過。實測 INTC/NVDA/AMD 跨公司比較，
      Q4 缺洞大幅減少（剩下的缺洞是真的沒有 XBRL 可抓的舊財報，屬於 SEC 端限制
      非程式問題，見 `docs/TODO.md` D9 附近說明）
- [x] **`google-generativeai` 汰換為 `google-genai`（2026-08-21，TODO D7）**：
      舊 SDK 官方已終止支援。改動三處同樣的 pattern——`fetcher_nongaap.py::_ai_request()`、
      `override_engine.py::_llm_call()`、`main.py::_test_ai_worker()`（設定面板
      「測試連線」按鈕）：`genai.configure()` + `genai.GenerativeModel(model)
      .generate_content(prompt)` 改成 `genai.Client(api_key=...).models
      .generate_content(model=model, contents=prompt)`。回傳物件同樣有 `.text`
      屬性，`errsafe.py::_exc_status()` 原本探測的 `code` 屬性名剛好對上新版
      `google.genai.errors.APIError.code`，不用跟著改。`requirements.txt` 的
      `google-generativeai>=0.8.0` 改成 `google-genai>=1.0.0`，`launcher.ps1`
      的安裝說明文字同步更新。新增測試：`test_fetcher_nongaap.py` 與
      `test_override_engine.py` 各一個，用 `unittest.mock.patch("google.genai
      .Client")` 驗證呼叫參數正確；`main.py::_test_ai_worker()` 屬 GUI 背景
      執行緒方法，本專案原本就沒有 GUI 層測試覆蓋，維持現況未補測試。
      本環境沒有真實 Google API Key，先只驗證了 SDK 呼叫介面／參數正確；
      **CTH 已用真實金鑰在「進階設定」按「測試連線」實測成功**，端對端驗證完成
- [x] **`Chart_<指標>` X 軸改用期末結算日，修正折線亂連線（2026-08-21，
      TODO F3 第 3 項）**：`comparison_writer.py::write_chart_sheets()` 原本
      用財季標籤（`FY2024Q3`）當 X 軸類別——不同公司財年結束月不同（AAPL 9
      月結束、NVDA 1 月結束），同一欄標籤在不同公司代表不同月份，且 Excel
      對缺值期間預設直接連到下一個有值的點，畫出誤導折線。改用
      `Compare_Data` 已有的期末結算日列（絕對日期，如 `20240331`，緊接在
      資料列上方）取代財季標籤，並設 `chart.display_blanks = "gap"` 讓缺值
      顯示為斷點不連線。屬於資料正確性修正，非排版美觀。新增 2 個測試於
      `test_comparison_writer.py`，全專案 954 個既有測試皆過。F3 其餘 4 項
      （圖表尺寸、座標軸單位、margin、圖例位置）未動，見 `docs/TODO.md`
- [x] **跨公司財務比較功能（2026-08-20，TODO F1，CTH 提出）**：新增 Tab4
      「跨公司比較」，可同時選多家公司、多個財務指標，輸出獨立的新格式
      Excel，跟現有單一公司抓取流程完全分開。架構：`src/comparison.py`
      對每個 ticker 呼叫既有 `fetch_gaap_statements()`，用
      `ratios.py::build_ratio_table()` 算比率後重組成
      `{指標:{公司:{期間:值}}}`；`src/comparison_writer.py` 用
      `openpyxl.chart`（本專案首次使用）寫出 `Compare_Data`（原始資料，
      每指標一區塊，帶靜態「期末結算日」列）／`Snapshot`（活的，
      `INDEX`/`MATCH` 公式對日期輸入格即時算快照，黃底提示可編輯）／
      `Snapshot_Manual`（空白，供人工貼值凍結，給未來程式讀）／
      `Chart_<指標>`（每指標一張折線圖，長條圖靠 Excel 原生「變更圖表
      類型」切換，不重複產圖）。Tab4 選擇視窗支援公司 ticker 自動完成
      （沿用 `company_cache.json`）＋逗號批次貼上、指標分類下拉（跨分類
      累加勾選）、季度/年度切換。`ratios.py` 的 `RATIO_DEFS` 加上
      category 欄位（英文機器鍵，供分類下拉分組），新增 21 個比率
      （Debt Ratio、D&A/Revenue、EBITDA($mm)、ROIC 等）。
      `excel_formatter.py` 新增 `($mm)` 單位後綴與可共用的
      `unit_format_for()`。新增測試：`test_comparison.py`（5）、
      `test_comparison_writer.py`（10）、`ratios.py`/`excel_formatter.py`
      各自補測試，全數通過；拿 NVDA/AMD 真實資料端對端驗證過。設計文件
      見 `docs/superpowers/specs/2026-08-20-cross-company-comparison-design.md`，
      實作計畫見 `docs/superpowers/plans/2026-08-20-cross-company-comparison.md`
- [x] **`Data_Financials(Q)` 補上 Q4（2026-08-20，TODO D0-1，CTH 提出「Q4
      資料不見」）**：SEC 沒有 Q4 的 10-Q，公司只交 Q1/Q2/Q3，過去
      `Data_Financials(Q)` 因此永遠缺 Q4，連帶 TTM 類比率（ROE／ROA／FCF per
      Share／淨負債EBITDA）湊不到連續四季，多半是空的。`fetcher_gaap.py`
      新增 `_synthesize_q4()`：有 10-K 年報可用時，IS/CF（流量項）用
      「年報 − Q1 − Q2 − Q3」反推，BS（存量項）直接取年報值。只補模板列，
      overflow 列（公司特有科目，季報/年報列序不保證對齊）留空；只抓季報
      不抓年報時優雅跳過；不處理 NG/segment 表；Excel 上不特別標示 Q4 為
      推算值。細節見 `docs/ARCHITECTURE.md`「Q4 推算」章節。7 個新測試
      （`tests/test_fetcher_gaap.py`），156 個 fetcher_gaap 測試、848 個
      非 live 測試、906 個含真連線 SEC EDGAR 的 live 測試（NVDA/AVGO/PLTR/
      ARLO 等）全過，0 failed
- [x] **輸出設定「設為預設」按鈕說明、SEC Identity 為何要填（2026-08-18，
      TODO E17/E18）**：Tab1 輸出設定的「設為預設」按鈕原本沒有任何說明，
      不知道是設定什麼，下方加一行固定顯示的灰字說明（只套用這次執行、按了
      才會存成下次初始值），不用點/hover 才看得到。Tab3 SEC EDGAR Identity
      區塊的格式提示上方加一句解釋，講清楚是 SEC EDGAR Fair Access 政策
      要求、姓名/信箱只當識別用不會挪作他用、沒填會抓不到資料。Tab1/Tab2
      的未設定警告維持簡短，原因說明集中放在 Tab3，沒有兩處重複寫一樣的話。
      四語都補了
- [x] **「管理 Watchlist」按鈕搬到 Tab2（2026-08-18，TODO E19）**：原本是
      跨頁籤都看得到的持久按鈕（`root` 層級，Tab1/Tab2 都有），但這顆只跟
      批量更新有關，Tab1 單一公司用不到。搬進 `_build_tab2` 最上方，Tab1
      不再顯示
- [x] **GAAP 抓取進度提示（2026-08-18，TODO E12）**：`progress_bar` 元件早就
      有，但 `fetch_gaap_statements` 抓取期間完全沒回報進度，整個 GAAP 階段
      進度條停在同一格不動、跟卡死一樣。仿照缺漏帳本（`collect_gaps`）同一套
      ContextVar 手法新增 `report_progress(cb)`，搭 `_build_is_table`／
      `_build_bs_table`／`_build_cf_table` 既有的 `_note_ok()`/`_note_gap()`
      掛鉤點回報「第幾份/共幾份」，不改這三個函式的簽名。GUI 顯示「抓取 GAAP
      財報中（N/M）」，單一公司／批次都有。範圍不含 segment/dynamic 兩個較
      小的次要建表流程。實測：COHR 30 次 tick 精確對上算好的分母；CRWV 透過
      GUI 真的抓一次，進度條從 1/15 跑到 15/15
- [x] **`launcher.ps1` 安裝進度可見（2026-08-18，TODO E7）**：根因是
      `uv pip install` 兩處都帶了 `-q`（靜音）旗標，把 uv 預設會印的
      「正在下載哪個套件＋進度條」關掉了——不是沒寫進度條，是進度條被自己
      關掉。對照 `windows-tool-templates.md` 範本才發現這條規定，兩處 `-q`
      都拿掉；補上範本的 `Write-Step`（帶時間戳的進度提示），下載 uv／建
      虛擬環境／裝套件前都會印一行帶時間的狀態。首次安裝說明（要裝哪些
      元件）原本就有，不是這次新增的
- [x] **「查可用期間」掃描補上耗時 log（2026-08-18，TODO E3 查證用）**：
      `_preview_scan_worker`（`main.py`）原本完全沒寫 `logs/app.log`，CTH 回報
      「第一次按只顯示代號、第二次才查到期間」但這個環境三個角度都測不出來
      （清空 edgar 本機快取模擬全新安裝耗時仍 ~1.5 秒；模擬點擊測試第一次就
      正常顯示；程式邏輯讀過沒找到「兩次才生效」的路徑），懷疑是 CTH 那台
      機器第一次連線特別慢（防火牆/Proxy/防毒攔截），但這邊重現不了也排除
      不了。先補上起訖 log（含失敗時的完整例外內容，比 GUI 上只顯示例外類型
      名詳細），下次重現時貼 log 就能直接判斷是連線慢還是另一種 bug
- [x] **Watchlist 視窗可用性 3 項（2026-08-18，TODO E13，「組 C」）**：群組列
      「重新命名／刪除群組」改成照 Windows 慣例靠右對齊，群組名不管多長都不會
      把按鈕擠出去；公司列的 📁／[x] 圖示按鈕原本沒有任何文字說明，改成清單
      上方加一行說明講一次用途，不用每列都重複塞字；「名稱庫：尚未建立」重新
      措辭成「公司名稱快取：尚未建立」，講清楚是什麼（Ticker 自動轉公司全名
      用的）、沒建立會怎樣（新增公司變慢幾秒但還能用）、要怎麼辦（點「更新
      名稱庫」），四語都改了
- [x] **視窗尺寸與內容撐開（2026-08-18，TODO E4/E6/E11/E16，「組 B」）**：
      視窗寬高從寫死的 900×720 改成算工作區比例（寬 2/3、高 0.8，各自夾在
      900~1300／780~900px 之間），CTH 的螢幕上實測會落在 2/3 寬度左右，不會
      隨螢幕大小固定死一個像素數。真正解掉「視窗變寬但內容沒撐開」的根因是
      Tab1／Tab3 的內容欄位一直沒設 `columnconfigure(0, weight=1)`——只放大
      視窗本身沒用，欄位要有 weight 才會分到多出來的寬度；連帶輸出資料夾路徑
      欄位改成撐滿剩餘寬度（不再固定 26 字元，長路徑不會被切字）、SEC Identity
      說明文字的 `wraplength` 從寫死 520px 改成跟著容器寬度即時調整。高度拉高
      後 Tab3 的固定高度可捲動容器 (`_TAB3_HEIGHT`) 重新量測校準（342→355），
      log 區可視高度明顯變高。Tab3 的存檔按鈕搬到頁籤最下面的固定 footer（不
      放在可捲動內容區裡，不用捲到底才看得到），新增「還原」按鈕把畫面欄位
      改回上次已存檔的值。順手清掉 `frame_log` 上兩行對 `pack()` 容器沒作用的
      無效 `rowconfigure`/`columnconfigure`。**未做**：關閉視窗前的未儲存提示
      （E11 的另一半，範圍待跟 CTH 確認）
- [x] **GUI 易用性小修 6 項（2026-08-18，TODO E8/E9/E10/E15）**：抓取中連「查可用
      期間」按鈕一起 disable，原本會跳警告視窗的重複點擊檢查整段刪掉；設定頁與
      Sheet 面板共用的固定高度容器補上滑鼠滾輪綁定；Ticker 輸入欄加寬（12→18）；
      Tab1/Tab2 的「進階設定」小按鈕改名「報表類型」，不再跟 Tab3 頁籤撞名；
      Tab1/Tab2 新增 SEC Identity 未設定警告（點了跳去 Tab3）；輸出資料夾/檔名
      格式不再靠改 radio、瀏覽資料夾就悄悄存成全域預設，改成 Tab1 明確按「設為
      預設」才存檔，Tab2 加一行唯讀文字顯示目前的全域預設在哪
- [x] **打包改成雙擊自助 + 排掉內部文件（2026-08-17，同日後續）**：新增
      `scripts\打包.bat`（薄 BAT）+ `scripts\pack.ps1`，CTH 雙擊就能自己壓，
      不必開 PowerShell 也不必叫 AI。12 項驗證跑在畫面上，**任一項沒過就刪掉 zip
      並 exit 1**——寧可不產出，也不要產出一包不知道能不能傳的東西（FAIL 路徑
      有實測：在 `src/` 埋一個 `.log` 進去，第 5 項抓到、zip 被刪、exit 1）。
      另把 `docs/` 排除出 zip（只留 `README.md` 有連到的 `8k-period-off-by-one.md`）
      ——`CHANGELOG.md` 一個就 84 KB，全是內部開發紀錄，收件人一個字用不到。
      包從 340 KB 降到 **187 KB**。
      順手清掉兩個誤入版控的檔：`company_cache.json`（405 KB，程式會自己重建）與
      `20260814 sec tool.zip`（124 KB，舊打包產物），`git rm --cached` 停止追蹤、
      檔案留在磁碟。**`.gitignore` 對已追蹤的檔案無效**，光加規則不夠。
      追蹤內容從 1906 KB 降到 ~1377 KB
- [x] **打包散布給非技術使用者（2026-08-17）**：`docs/PACKAGING.md` 是給 AI 照著
      執行的作業指示（白名單複製、產出 `dist/SEC-Financial-Fetcher-YYYYMMDD.zip`、
      解壓到暫存目錄跑 11 項自我驗證）。zip 內附 `先讀我.txt`（來源
      `docs/RECIPIENT-README.txt`），只講兩件事：填 SEC EDGAR Identity、
      首次啟動選語言。機敏資訊天生在專案外（`config.json` 在 `%APPDATA%`），
      驗證步驟仍實際 grep 一次金鑰與信箱樣式
- [x] **launcher 改成 uv 優先，不再依賴 winget 與系統 Python（2026-08-17）**：
      原本 `winget install --id Python.Python.3` **已經失效**（實測 exit 20
      No package found，winget 上只剩 `Python.Python.3.10`~`3.14`），沒 Python
      的電腦跑到那裡必定卡死。改成讓 uv 自己下載 Python
      （`uv venv venv --python 3.13`，python-build-standalone，約 20MB）。
      連帶消掉三個卡點：死掉的 package ID、裝完 PATH 沒刷新要再雙擊一次、
      舊 Win10 沒 winget 就叫使用者自己去 python.org。步驟從 3 步變 2 步。
      另補 TLS 1.2（舊 Win10 的 PS 5.1 可能還預設 1.0/1.1，抓 astral.sh 會直接失敗）。
      實測：uv 下載的 CPython 3.13.12 上 `tkinter` 開得了視窗，`edgar`／`openpyxl`
      都裝得起來
- [x] **抓取缺漏會主動報告（2026-08-17）**：某幾期沒抓到時那幾期留空、其餘照常
      產出，並在三處講明缺了哪幾期——GUI 橘字、`logs/app.log`、**Excel 的
      Index 第一頁橘底那列**（外加 `Data_Meta` 的 `Fetch Gaps`）。GUI 的 log
      關掉就沒了，但三天後重開這份 Excel 警告還在，那才是使用者會搞混的時點。
      網路類例外會退避重試 2/4/8 秒救閃斷；連續 3 期失敗就停止重試（不再乾等，
      整個網路斷掉時 40 份財報各等 6 秒等於白等 4 分鐘）但**不中止**。
      「是網路還是資料」由失敗當下實際戳一次 SEC 決定，不猜例外類別名稱。
      一期都沒抓到才不寫檔——空殼 Excel 會蓋掉舊檔，是唯一不可逆的傷害
- [x] **覆蓋既有輸出檔前先提示（2026-08-17）**：講明蓋掉哪個檔、`My_*` 分析頁
      會保留。批量一次列出「12 檔中有 8 檔已存在」問一次，不逐檔跳
- [x] **進階設定改成第三個頁籤（2026-08-17）**：原本是彈出視窗
- [x] **視窗置中且不再跳動（2026-08-17）**：走 Win32 `SPI_GETWORKAREA` 取得扣掉
      工作列的可用區域自己算座標（原本只給大小不給座標，Windows 用階梯式落點
      越開越低，最後掉到工作列底下）。尺寸固定 900×720，掃描完不再從 650 撐到 800
- [x] **掃描鍵與執行鍵不再共用 ▶（2026-08-17）**：改成 `🔍 查可用期間` 與
      `⬇ 開始抓取`，掃描期間顯示「約需 10 秒」
- [x] **多語言（2026-08-14）**：GUI 與 Excel 顯示文字支援繁中／简中／英文／日文。
      進階設定最上方選語言，重開程式生效（選完會跳英文視窗問是否重啟）。
      Excel 只有 B 欄與 Index 版面隨語言變——A 欄英文機器鍵與 C 欄公司原文
      任何語言下都不變，既有跨檔案公式不受影響。新增語言只要複製一份
      `src/locales/*.py` 再加一行登錄。詳見 `docs/ARCHITECTURE.md`「多語言」
- [x] **首次啟動選語言（2026-08-15）**：第一次開程式跳一個純英文的 `Language` 視窗，
      四個按鈕點一下就記住，之後不再出現。判斷依據是 `config.json` 的 `language`
      不是合法代號，所以既有使用者也會被問一次
- [x] `cli.py --lang`：讓 skill 直接指定產出的 Excel 語言，不必先去改 GUI 設定
- [x] B1 Overflow Rows：IS/BS/CF 三表各自追加未被模板消耗的 XBRL rows，確保不遺漏任何財務數字
- [x] Non-GAAP overflow 自動分流至 `Data_Financials_NG(Q/Y)` 獨立 sheet（label 含 "adjusted"/"non-gaap"/"excluding" 等）
- [x] Auto-repair override engine（新 ticker 首次 fetch 自動診斷並修復缺失 key rows）
- [x] config.json 搬到 %APPDATA%\SEC Financial Tools\（不進 git，啟動時自動 migrate）
- [x] Watchlist 每間公司獨立輸出路徑（📁 按鈕，存於 watchlist item output_dir）
- [x] Excel 自動美化（深藍色 header、交替底色、section 分隔、subtotal 粗體）
- [x] 財務數字自動 ÷1M（EPS 除外），套用千分位格式
- [x] Index sheet（第一頁，列出所有 sheet 用途 + 最早/最新期間）
- [x] Data_Financials(Q)（季報）+ Data_Financials(Y)（年報 10-K）雙 sheet
- [x] Per-ticker output path memory（ticker_paths in config.json）
- [x] Non-GAAP fetching from 8-K press releases（Data_EPS_Recon + Data_NonGAAP）
- [x] nongaap_cache.json 增量快取（每季 AI 呼叫結果本機快取）
- [x] 單一公司 GAAP 財報抓取
- [x] Excel 輸出（Data_Financials 三表合一 + Data_Seg_* + Data_Meta）
- [x] 批量更新 (Watchlist)
- [x] Watchlist 管理 popup
- [x] 進階設定 popup（AI config, identity, output dir）
- [x] max_filings 設定（Advanced Settings，預設 80 筆 = 約 20 年）
- [x] Ticker 標識（每個 sheet A1）
- [x] IS 固定 22 行模板（含 D&A/SBC/Minority Interest/Total Non-op）
- [x] BS 固定 41 行模板（完整 Assets / Liabilities / Equity）
- [x] CF 固定 25 行模板 + Free Cash Flow 衍生計算
- [x] B 欄 Original Item（公司的 XBRL 原始標籤）
- [x] ProfitLoss fallback（BA、TSLA、XOM、WMT 用 ProfitLoss 報 Net Income）
- [x] D&A label fallback（TSLA std_concept = nan 情況）
- [x] Total Non-op DERIVED fallback（Pre-tax − Operating Income）
- [x] Gross Profit DERIVED fallback（Revenue − COGS，修復 COHR）
- [x] GOOGL encoding fix（非 ASCII 字元 NFKC normalize）
- [x] match="first"|"last" + label_hint 精確比對（解決 BS 重複 std_concept 問題）
- [x] `_match_is_row` 串接優先級（cascading）：label_hint 不符時跳至下一優先級而非誤匹配
- [x] CF Net Income fallback regex OR（`NetIncomeLoss|ProfitLoss`）
- [x] CF 彙總行 label_hint 改為 `^net cash|^cash`（同時支援 AAPL 與 XOM 排除）
- [x] 實機測試（GAAP）：AAPL ✅ TSLA ✅ BA ✅ XOM ✅ NVDA ✅ COHR ✅（2026-04-23）
- [x] 快速掃描顯示最新季度＋送件日（2026-08-13）

### 待辦
- [x] 實機測試（Non-GAAP）：NVDA ✅（正規化後每季 6 指標整齊；修前 34 個稀疏 row）；AAPL ⚠️ 少報屬預期行為（2026-04-19）
- [x] main.py 舊名稱掃描：確認無 Data_IS/BS/CF 殘留參照 ✅（2026-04-19）
- [x] Non-GAAP：nongaap_cache.json ticker 隔離（多公司共用 output_dir 安全）✅（2026-04-25）
- [x] CF：Q2/Q3 YTD overflow 跨 filing 減法 — overflow 行與模板行同邏輯 ✅（2026-04-26）
- [ ] Non-GAAP：Data_EPS_Recon 從未產生（edgartools eps_reconciliation 對 AAPL/NVDA 均回傳空）
- [x] Non-GAAP：NVDA 指標名稱正規化（剝除期間 token、去重比較期間、過濾展望指標）✅（2026-04-26）
- [x] CF：Q2/Q3 YTD→季度換算（Q2 = Q2_YTD − Q1；Q3 = Q3_YTD − Q2_YTD）✅（2026-04-23）
- [x] 金融股警告（GS/JPM 等）：fetch 完成後 log 顯示模板限制警告 ✅（2026-04-26）
- [x] 批量更新（Tab 2）加入 Non-GAAP 支援：新增 checkbox + `_worker_batch` 支援 Non-GAAP ✅（2026-04-26）

---

## 更新記錄

### 2026-08-14（晚間）

**修「處理進度」log 消失——快速掃描面板一展開就把 log 擠不見（實測根因，非猜測）**

CTH 回報點「執行」後看不到進度視窗，附圖顯示「處理進度」LabelFrame 底下
完全空白。用 `winfo_reqheight()`/`winfo_height()` 實際量測 `SECFetcherApp`
的 grid 佈局才找到根因：

- 主視窗高度鎖死 650px 不會自動撐大（`main.py` `__init__` 的
  `geometry()`，2026-08-12 的既有設計，防止可選 Sheet 清單把視窗撐爆）
- 快速掃描找到 segment 表、展開「可選 Sheet」面板時，Tab 1 實際需要的高度
  比閒置時多約 160px，但視窗高度不會跟著長；`處理進度` 是**唯一**
  `rowconfigure(weight=1)` 的列，這 160px 的缺口全部由它吸收
  ——實測面板一展開，`log_text` 的可視高度直接被壓到 **1px**，等於整個
  消失，正是回報看到的症狀。這是 **pre-existing 問題**（本次的
  `preview_sheets` 快速掃描顯示最新季度功能另外加了一行 Label，一度把它
  改得更糟，已在同一次修正中拿掉，改寫進 LabelFrame 自己的標題列，
  不佔額外一行、高度成本 0）
- 修法：**視窗高度跟著面板展開/收合動態切換**，而非死守 650px。
  `_build_sheet_panel()` 在面板展開時呼叫
  `root.geometry("700x800")`，收合（含重新掃描前的清空、掃描到 0 張表）
  時呼叫 `root.geometry("700x650")`。800 是實測值——`log_text` 在該高度
  下恢復到約 120px（4~5 行可視），比閒置時的 107px 還寬裕
- 驗證方法：寫一次性診斷腳本直接呼叫 `SECFetcherApp` 並用
  `winfo_height()` 量測 `frame_log`／`log_text` 的**實際分配高度**（不是
  `winfo_reqheight()`——那個在視窗尺寸鎖死時會回報「假想」的未裁切需求值，
  跟畫面上真正看得到的高度是兩回事，一開始被這點誤導繞了一圈）。
  idle→scanned→reset 三態各測一輪，log_text 107→120→107，不再出現 1px
  的塌陷
- 全套測試 673 過（GUI 無自動化測試，本次修正靠上述量測腳本驗證，非
  test suite 覆蓋）
- **待 CTH 實機驗收**：目前開著的視窗是舊 code 起的，需要關掉重開
  （或重新雙擊 `啟動器.bat`）才會套用新的動態視窗高度

### 2026-08-14

**打包給朋友前的地雷排查：`launcher.ps1` 加強 Python 偵測**

- 全新 Windows 電腦沒裝過 Python 時，PATH 裡常有內建的「App execution alias」
  `python.exe` 存根——`Get-Command python` 找得到它、看起來像已安裝，但實際
  執行只會跳出 Microsoft Store 頁面，不會印版本號。原本的偵測只看指令存不
  存在，會誤判成「已安裝」而跳過安裝流程，之後才默默失敗。改成二次確認
  `python --version` 的輸出符合 `^Python \d+\.\d+`，安裝完成後的判斷也套用
  同一套邏輯
- 實測：清空 venv 重新用 `uv venv` + `uv pip install` + 直接跑
  `python src\main.py`，GUI 存活未閃退，確認偵測邏輯改動沒有連帶弄壞正常
  安裝路徑
- 打包格式由 `.rar` 改 `.zip`——`.rar` 需要朋友電腦另外裝 WinRAR/7-Zip 才
  打得開（Windows 11 24H2 以下都沒有內建支援），`.zip` 全版本 Windows
  資源管理器原生可解壓縮，少一個「連第一步都過不去」的地雷
- 打包範圍維持最小化（`啟動器.bat` / `launcher.ps1` / `README.md` /
  `requirements.txt` / `src/`），不含 `venv`／`docs`／`tests`／`scripts`／
  `.git`／個人設定；`config.json`（Identity + API Key）本來就存在
  `%APPDATA%`，未被打包
- **未處理、留給朋友看說明文字自行判斷**：解壓縮/雙擊 bat 時 Windows 可能
  跳出「已保護您的電腦」的 SmartScreen 警告（檔案帶網路下載標記）；
  `uv` 安裝走 `Invoke-RestMethod | Invoke-Expression`，公司防火牆/防毒
  較可能擋下，擋下時只會看到「uv 安裝失敗」，沒有進一步診斷

### 2026-08-13

**快速掃描新增「最新季度＋送件日」顯示**

- `fetcher_gaap.preview_sheets()` 回傳型別由 `list[str]` 改為 `dict`：
  `sheets`（原本的可選 Sheet 清單）＋新增 `latest_label`（如 `FY2026Q1`）、
  `latest_period_end`（期末日）、`filing_date`（送件日）。三個新欄位全部
  從掃描時已經抓到的最新一筆 10-Q filing metadata（`period_of_report` /
  `filing_date`）算出，**不多打一次 API**
- 財年結束月走 `Company.fiscal_year_end` 屬性（一次請求），不用
  `_detect_fy_end_month()`——那個要 `filing.obj()` 抓 10-K 全文，快速掃描
  用不起
- `filing_date` 是 EDGAR 的送件／公開日，**不是** SEC 受理的精確時間戳
  （這版 edgartools 沒有 `acceptance_date` 欄位，要拿的話得繞過 edgartools
  直打 submissions JSON，本次未做）
- **循環匯入地雷**：`fiscal_quarter_of` 原本想在 `fetcher_gaap.py` 頂層
  `from fiscal_input import ...`，但 `fiscal_input.py` 會 `from
  excel_formatter import ...`、`excel_formatter.py` 又 `from fetcher_gaap
  import StatementTable`——三檔案兜成一圈，模組載入到一半互相要對方還沒
  定義好的名字就炸。改成在 `preview_sheets()` 函式內部延後 import
- GUI：「可選 Sheet（掃描後顯示）」面板上方加一行藍字
  `最新資料：FY2026Q1（期末 2025-12-27）｜送件日 2026-02-01`；抓不到
  財季時顯示「無法判斷財季」而非空白。快速掃描的「？」說明泡泡同步補充
- `tests/test_fetcher_gaap.py`：3 個既有 `preview_sheets` 測試改配合新
  dict 回傳型別，新增 1 個驗證財季換算的測試。全套 673 過（不含
  `test_live_snapshots.py`，那個連真實網路）

### 2026-08-12（目錄結構整理）

**根目錄改走 `windows-tool.md` 標準結構：MD 進 `docs/`，17 個 `.py` 分 4 批搬進 `src/`**

- 根目錄的 `20260812 sec工具.rar`（466KB，未追蹤於 git）確認為手動備份後直接刪除
- `ARCHITECTURE.md`／`CHANGELOG.md`／`PITFALLS.md`／`TODO.md`／
  `docs_statement_template_proposal.md` 全部 `git mv` 進 `docs/`；
  `doc-init-protocol.md` 本來就有「`docs/` 找不到就 fallback 讀根目錄」的相容邏輯，
  搬完不影響 AI 讀取；`pytest -q` 跑一次確認沒有測試寫死根目錄路徑
- 17 個 `.py`（含 `main.py`／`cli.py`／`fetcher_gaap.py` 等）分 4 批搬進 `src/`，
  每批跑一次 `pytest -q -m "not slow"` 全綠才搬下一批：
  1. `main.py`／`cli.py`（入口，沒人 import）
  2. `errsafe.py`／`metric_rules.py`／`override_engine.py`／
     `press_release_tables.py`／`zh_labels.py`（葉節點）
  3. `ratios.py`／`segments.py`／`nongaap_layout.py`／`excel_formatter.py`／
     `fetcher_nongaap.py`／`fiscal_input.py`／`excel_writer.py`／`output_tables.py`
  4. `config.py`／`fetcher_gaap.py`（被最多檔案 import，留到最後）
  - 全程維持 flat import（`from fetcher_gaap import ...` 語句本身沒改），
    靠 `conftest.py`／`scripts/*.py` 過渡期同時把根目錄跟 `src/` 塞進
    `sys.path`，搬完後拔掉根目錄那段
  - **踩到的坑**：`main.py` 的 `SCRIPT_DIR = Path(__file__).parent` 搬到
    `src/` 後從「專案根目錄」變成「`src/`」，`company_cache.json` 快取路徑、
    輸出資料夾預設路徑、舊版 `config.json` 遷移路徑三處都跟著算錯位置。
    改用既有的 `_find_project_root()`（往上找 `launcher.ps1`）算出
    `PROJECT_ROOT` 取代這三處的 `SCRIPT_DIR`
  - `tests/test_no_ai_by_default.py` 有一處靜態讀 `main.py` 原始碼文字做規則檢查，
    路徑寫死在根目錄，一併修正
- `launcher.ps1` 呼叫路徑改成 `python src\main.py`；用背景程序啟動驗證
  log 有正常寫入「環境就緒」且無 `[CRASH]`、程序存活有回應，**實際雙擊驗收
  待 CTH 確認**
- `conftest.py`、`company_cache.json`、`config.example.json` 判斷後留根目錄不搬：
  `conftest.py` 是 pytest 探索機制（不是被 import 的程式碼），搬進 `src/`
  會讓 pytest 找不到、marker 註冊失效；後兩者是資料/範本檔，規則沒明講歸屬，
  CTH 確認不動

### 2026-08-12

**BS 新增非流動資產／負債合計列；權益與 CF 合計列粗體修正；GUI 調整**

- `fetcher_gaap.py` BS_TEMPLATE 新增 `Total Non-current Assets`／
  `Total Non-current Liabilities`：先試接 XBRL 的 `AssetsNoncurrent`／
  `LiabilitiesNoncurrent`（`^us-gaap_...$` 精確錨定字串比對），查不到就用
  `Total − Current` 相減算。**踩到的坑**：一開始用裸子字串比對 `fallback`
  tag，`AssetsNoncurrent` 是 `OtherAssetsNoncurrent` 的子字串，結果新列的值
  跟「其他非流動資產」一模一樣——AAPL 實測揪出後改成錨定比對，驗證
  流動＋非流動＝總計，兩邊都對得起來
- `excel_formatter.SUBTOTAL_CONCEPTS` 補上新兩列；順手修一個既有 bug——
  原本寫的 `"Total Equity"` 對不上實際列名（`Total Equity — Parent` /
  `incl. NCI` / `Total Liabilities & Equity`），權益合計列從沒真的粗體過；
  也補上 `Investing Cash Flow`／`Financing Cash Flow`，跟既有的
  `Operating Cash Flow`／`Free Cash Flow` 一致處理
- `tests/test_live_snapshots.py` 修掉沒跟上 commit `71ed50e`（overflow 移到
  底部）版面重構的舊索引算法：`test_b1_overflow_rows_nonnull`、
  `test_b1_ng_sheet_structure`、`_cf_overflow_rows`（3 個 cf_overflow 測試
  共用）改用 `OVERFLOW_SECTION` 標題列位置直接定位合併後的單一 overflow
  區塊，不再假設 overflow 各自接在 IS/BS/CF 模板列後面——14 個假陽性失敗
  （測試沒跟上版面，不是資料品質退化）清掉。live test 覆跑驗證：14 failed → 14 passed
- GUI：主視窗改 700×650 並鎖高度（`geometry()` 只呼叫一次即關閉自動撐大），
  快速掃描跳出的可選 Sheet 清單改成固定高度可捲動容器、3 欄排版；快速掃描
  按鈕旁加「？」說明泡泡

### 2026-08-09（後半）

**8-K 季度標籤 off-by-one 實修：方案 B+（測試 656 → 672）**

CTH 從報告的三個選項選了 B+：**B 的成本 + A 的正確性**。列清單階段一行不動
（那裡零下載，正確性要求本來就低），改在**下載之後**重算。

`cli.py press-release --json` 每一季多三個欄位，舊欄位一個都沒改：

| 欄位 | 內容 |
|---|---|
| `period_end` | 從已解析的表格抓到的真實財期結束日 |
| `fiscal_label` | `period_end` + 財年結束月算出的正確財季，與 `Data_Q` 同一套慣例 |
| `fiscal_label_source` | 固定 `"period_end"` |

頂層另加 `fy_end_month`（來自 `Company.fiscal_year_end`，**一個 ticker 一次請求**，
不必為了問財年下載 10-K）。查不到就是 `None`、`fiscal_label` 留空——**不猜 12**，
非 12 月結算的公司會被整批標錯一到三季。

期末日的規則繞了一圈才定下來：

1. 第一版「取不晚於申報日的最新日期」——AMD 抓到發布日（安全港聲明寫
   `speak only as of August 4, 2026`），整家標錯一季
2. 第二版「優先採信 `ended` / `as of` 後面那個日期」——**更糟**，AMD／INTC／
   AVGO 三家全錯：關鍵字指到的是安全港聲明、資產負債表去年年底、上一季的
   註腳，而它們真正的期末日只是**沒有引導詞的表頭**（colspan 展開後只剩日期）
3. 定案：**單純取「不晚於申報日前 3 天」的最新日期**。3 天的緩衝排掉發布日，
   財報最快也要期末後兩週才發，不會誤傷

另外每一欄要**直向串起來**再找一次日期：NVDA／INTC 把表頭排成 `April 26,`
一列、`2026` 下一列，只看單一儲存格 NVDA 三季全空。

驗證 `scripts/verify_8k_fiscal_labels.py`（新增）：15 家 × 8 季 = 120 份，
期末日抓取率 **120/120**，每家的新舊標籤偏移都是常數，且與
`docs/8k-period-off-by-one.md` 用新聞稿內文獨立推出的偏移**完全一致**。

順手修掉一個必炸的洞：`cli.py press-release` 不給 `--json` 時在 cp950 主控台
只印得出 `失敗 -> UnicodeEncodeError`（`⚠` 編不進 cp950）。`main()` 開頭
強制 stdout/stderr 轉 UTF-8（`errors="replace"`）。新聞稿內文的重音字母、
`™` 之類同樣會炸，逐字元挑符號治不完。

**沒做**：`--years` 篩的仍是發布日換算的年份——篩選發生在下載前，那時讀不到
期末日。非 12 月結算公司在年份邊界可能差到 3 季，已寫進 `--help` 與 README。

---

### 2026-08-09

**8-K dedupe 由「保留最舊」改為「有附件優先、其次最新」（測試 652 → 656）**

`_list_earnings_filings()` 撞到同一季度標籤時原本保留最舊那份。實測 16 家 128 份撞到
2 次，**兩次都保留錯的**：

| 公司 | 舊規則保留 | 實際該留 | 後果 |
|---|---|---|---|
| WDC `FY2025Q1` | `0001193125-25-007725`（items `2.02,5.02`） | `0000106040-25-000005`（`2.02,9.01`） | 整季靜默消失 |
| QRVO `FY2025Q4` | `0000950103-25-013685`（2025-10-28 Preliminary） | `0001628280-25-048216`（2025-11-03 正式） | 拿到初步數字 |

新規則抽成 `_dedupe_by_label()`，兩層排序：

1. **有 Item 9.01 者優先**。9.01 是「Financial Statements and Exhibits」，新聞稿就是
   那個 EX-99 附件；只有 2.02 沒有 9.01 的那份根本沒有東西可以解析。WDC 那份壞的
   items 是 `2.02,5.02`（EDGAR 實際值），一條規則就分得開
2. **同層取最新**。preliminary 一定早於正式版，所以 QRVO 靠這條分

**仍然零下載**：兩層判斷都只讀 listing metadata（`items`），沒有 `obj()`，有測試釘著。
原本「保留最舊」的理由是「更正版的重新申報不該取代原始發布」，但
`get_filings(amendments=False)` 本來就把 8-K/A 排除掉了，這個理由早就不成立。

與季度標籤 off-by-one（TODO D4 後半）是獨立的兩件事，標籤怎麼改都不影響這條。

---

### 2026-08-08（晚間）

**D1 Excel 排版驗收通過（CTH 逐項親驗，執行順序表第 1 順位解除）**

標的是 `output/_final/` 七份（AAPL / NVDA / META / AVGO / MSFT / COHR / PLTR）。

通過：欄寬、三表底色與 5 列間隔、數字格式（÷1M／`0.0%`／每股兩位／`(x)`／`(days)`）、
新字型（文字微軟正黑體 + 數字 Consolas）、Index 新字級、`Index!B4` 連動第 1/3/4 列。

**凍結窗格維持 D3**——D0-4 提的「表頭有 5 列、往下捲看不到第 3~5 列」CTH 看過後
決定不改，該條結案。

**驗收中發現並修掉一個**：Index 提醒列（A5）會切到文字。查了才知道**原本就不夠**，
不是這次改字級才壞的——提醒文字顯示寬度 508 個半形當量、A~E 合併寬度 86，需要
約 6 行，而寫死的 28 只夠 2.4 行。合併儲存格不會自動調列高（Excel 的行為，不是
openpyxl 的限制），改成 `_wrapped_row_height()` 依文字長度與實際欄寬推算，實測 94.5。

自動複驗七份：字型只有微軟正黑體 + Consolas 兩種、九個關鍵列位七家全對、
A5 列高皆 94.5、品質 9/9（COHR 8/9 是舊有的 `Operating Income` 缺口，無 AI key
時 E2 診斷不啟動）。全程零 AI 呼叫。

**殘留待 CTH 判斷**：PLTR 最新季淨利率 54.9%，對其營運偏高，疑似一次性稅務項目，
無外部來源可對，不下定論。

---

### 2026-08-08（下午）

**排版：文字改微軟正黑體、數字改 Consolas、Index 字級放大一級（測試 635 → 651）**

以前所有 `Font()` 都沒帶 `name=`，Excel 各自套預設——英文 Calibri、中文新細明體，
同一張表混兩種字體。現在 `excel_formatter.FONT_NAME = "微軟正黑體"` + `_font(**kw)`
工廠，`excel_formatter` 與 `fiscal_input` 兩個模組所有字型統一走它，換字型只改一行。

**數字欄走 `NUMBER_FONT_NAME = "Consolas"`（等寬）**。原本 CTH 選的是「全部含數字欄
一種字型」，看過實際輸出後同日改成數字等寬——微軟正黑體的數字不等寬，`1,234,567` 與
`89,012` 上下位數對不齊，翻財報時很難掃。

判斷依據是**儲存格的值是不是數字**（`_is_number()`），不是欄號。用欄號會誤傷
`Data_Meta`（D 欄起是公司名這類文字）與 `Data_Segments`（長格式，D 欄起是分類名）。
副作用是第 2、5 列的日期字串（`2026-03-29`）留在微軟正黑體——它們是字串不是數字。

`_apply_row_styles` 改成 `_paint(cells, fill, **font_kw)`：同一列的文字版與數字版
字型只差 `name`，`bold` / `color` / `size` 完全共用，不會出現「換了字型結果粗體掉了」。

順手修掉一個洞：Index 表格的 B~D 欄（含中文說明）原本**根本沒設 font**，只設了底色，
所以那幾欄一直是新細明體——正是這次要修的症狀本身。

Index 字級整體放大一級：

| 位置 | 舊 | 新 |
|---|---|---|
| A1 公司抬頭 | 14 | **16** |
| A2 抓取日期／最新期間 | 9 | **10** |
| 表格（標題 + 內容 + 品質明細） | 10 | **11** |
| B4 財年起始月輸入格 | 11 | **12** |
| 財年核對提醒 | 9 | **10** |

提醒那列的列高一併 28 → 32：字級變大而列高不動，wrap 過的長字串尾巴會被切掉。

**測試寫成掃描式不是逐格斷言**：`test_every_font_on_data_sheets_uses_the_configured_family`
把整張 sheet 的字型收成集合比對，漏掉任何一處 `Font()` 就會紅。逐格斷言只會蓋到寫測試
當下想得到的那幾格，而這個 bug 的本質就是「有一處忘了帶」。合併範圍的填充格
（`A1:E1` 的 B1~E1）留在 openpyxl 預設 Calibri，Excel 只看左上角那格，不影響顯示。

**六份輸出重產**：AAPL / META / AVGO / MSFT / COHR / PLTR 已重抓（零 AI，只打 SEC EDGAR）。
NVDA 因檔案在 Excel 開著被鎖檔保護擋下——那正是 `check_output_writable()` 該有的行為。

---

### 2026-08-08

**財年起始月改成使用者可編輯，期間標籤全部用公式帶（測試 597 → 632）**

財年結束月是程式從 10-K 的 XBRL 欄名自動判讀的（`_detect_fy_end_month()`），
**會出錯**；出錯時整排財季標籤跟著錯，而使用者除了重跑程式沒有別的辦法，
看到的還只是一堆寫死的文字。

改法：`Index!B4`（黃底）放一格可編輯的財年起始月，定義名稱 `FY_START_MONTH`。
`Data_Financials(Q)/(Y)` 第 1、3、4 列改成引用它的 Excel 公式，改一格整本更新。
第 5 列（期末結算日）是 XBRL 真實日期，**永遠靜態**，是所有公式的錨。
Index 上有一段醒目的核對說明，明講怎麼對、以及哪些東西不會跟著變
（Index 表格的最早/最新期間、`Data_Ratios`、`Data_Meta` 是 Python 算好的靜態值）。

**順帶修掉 AAPL 真實存在的 off-by-one**

新公式與舊靜態值逐格比對 10 個活頁簿共 **389 格，只有 12 格有差，全部是 AAPL 的兩欄**：

| 期末結算日 | 舊標籤 | 新標籤 | 事實 |
|---|---|---|---|
| 2023-04-01 | `FY2023Q3` ❌ | `FY2023Q2` ✅ | Apple 自報 Q2 FY23，營收 94,836 |
| 2023-07-01 | `FY2023Q4` ❌ | `FY2023Q3` ✅ | Apple 自報 Q3 FY23 |

也就是說舊版的 AAPL 有兩欄標錯了一季。連帶效果：AAPL 原本是四家裡唯一
ROE／ROA／FCF per Share 有值的，但那四欄「看似連續」正是標籤錯造成的假象，
TTM 實際上跳過 Sep 季又重複 Dec 季——**數字是錯的**。改對之後正確地變成空白。

**52/53 週制：換算前先把期末日往前推 15 天**

美股期末日在月底前後浮動最多 6 天，WDC 的 FY2026 Q2 結束在 `2026-01-02`，
直接看月份會算成 Q3。往前推 15 天必定落回該季最後一個月。這正是
`docs/8k-period-off-by-one.md` 裡 COST／WDC／PANW 七份對不上的原因，
這次一併處理掉。

另加 `wb.calculation.fullCalcOnLoad = True`——openpyxl 不算公式、寫出去沒有
快取值，不強制重算 Excel 有機會直接顯示空白（看起來像整排標籤不見了）。

**四家實測（NVDA/AAPL/PLTR/AVGO，財年結束月 1/9/12/11 月）另外抓到的問題**

- **`Data_Financials(Q)` 永遠沒有 Q4**（Q4 沒有 10-Q，數字在 10-K）→ TTM 類比率
  湊不到連續四季，NVDA/AVGO/PLTR 的 ROE／ROA／FCF per Share／淨負債EBITDA 整列全空
- **多股別公司抓不到期末流通股數**：PLTR／GOOGL／META 的 `company.get_facts()` 裡
  `dei:EntityCommonStockSharesOutstanding` **0 筆**（Class A/B/C 分開標），
  TSLA 61 筆、COHR 62 筆正常。連帶 BVPS／FCF per Share／流通股數 YoY 空白
- 兩項都未修，列在 TODO D0

**跨公司列位再驗一次**：四家財年結束月完全不同，`Revenue` 8、`Gross Profit` 10、
`Operating Income` 17、`Net Income` 24、`Cash` 38、`Total Assets` 51、`OCF` 98、
`Capex` 99、`Free Cash Flow` 114 **完全一致**。

**文件同步**：README 有三處已經過期到會誤導的內容——欄位說明還寫「B 欄 =
Original Item、C 欄起 = 各季數據」（實際是 A/B/C/D 四欄、表頭 5 列）、
列位對照表寫 `Cash` 在 34 列（實際 38）、以及一整章「用 `Data_Std` 寫跨公司模板」
（那張 sheet 8/3 就刪了）。照現況重寫。

---

### 2026-08-07

**B1 CLI 工具層、B3 確定性表格解析、D4 前半調查（三項，測試 524 → 593）**

Non-GAAP 改走 skill 的三塊前置作業。全程零 API 呼叫——`gaap` 與
`press-release` 都只打 EDGAR。

**`cli.py`（新）**——給 skill 的入口，GUI 與核心函式一行沒動。

```
cli.py gaap AAPL --years 2023-2026 --xlsx out.xlsx
cli.py press-release ARLO --years 2025-2026 --tables --json
```

- `gaap AAPL --years 2023-2026`：23.7 秒，5 張 sheet。與 GUI 產的
  `output/_final/AAPL.xlsx` **逐格比對 4,089 格，差異 13 格全部是「抓取日期」**
  （08-03 vs 08-07）。輸出等價，不是「看起來一樣」
- `press-release ARLO --years 2025-2026`：5.9 秒抓 7 季全成功，每季
  2,863～6,629 字元（平均 4.1K），整份 JSON 107KB。同樣資料丟原文是 3.1MB
- 唯一重構：`main._append_ratio_table` 抽到 `output_tables.py`。`import main`
  會拉進 tkinter，但複製一份到 CLI 更糟——兩邊會慢慢長歪，而「GUI 產的
  Excel 跟 CLI 產的不一樣」極難發現。`main` 保留舊名別名
- 網路呼叫集中在 `_gaap_tables` / `_earnings_filings` / `_press_release_html`
  三個函式，24 個測試沒有一個碰網路。其中一個專門釘住「例外訊息不會流到
  stdout/stderr」

**`press_release_tables.py`（新）**——`pandas.read_html` + 版面規則，零 AI。

難點全在 Workiva 的版面雜訊：同一個數字重複寫進相鄰欄、`$` 與 `%` 各佔一欄、
期間之間插全空間隔欄。ARLO 的調節表因此是 24×30 網格，真資料只有 14×6。

> **關鍵地雷**：間隔欄要用「所有**資料列**都空」判斷，不能用「整欄都空」。
> Workiva 的表頭是 colspan 展開的，`Three Months Ended` 會把 15 欄（含中間的
> 間隔欄）全部填滿，用整欄判斷就一個間隔都找不到，三個期間的數字會併成一格。

實測 12 家最新 Item 2.02 8-K，全部零 ragged row、零異常寬表：

| 公司 | 原文字元 | 篩後字元 |
|---|---|---|
| ARLO | 450,392 | 4,372 |
| NVDA | 274,720 | 2,241 |
| MSFT | 901,488 | 2,102 |
| ORCL | 1,956,388 | 4,535 |
| AAPL / COST | — | **0（正確）** |

AAPL 與 COST 篩出 0 張不是 bug——這兩家本來就不報 Non-GAAP。

刻意保留而非丟棄的三種情況（資料異常要看得見）：同一格收斂出兩個不同數字時
併排輸出、落單的標題區塊仍輸出、含數字的財測區塊不會被當標題吃掉。

**8-K 季度標籤 off-by-one 調查（只查不修）**——`docs/8k-period-off-by-one.md`

TODO 原本記的是「幾乎都比實際財季晚一季」。查完發現**不是偏一季，是偏 −3
到 +1 季，偏多少由財年結束月決定**。16 家 128 份、成功比對 119 份，
**只有 16 份（13%）標對**。

| 財年結束月 | 公司 | 偏移 |
|---|---|---|
| 9 / 11 | AAPL、AVGO | **0** ✅（兩層誤差剛好抵消，不是程式有處理） |
| 12 | ARLO、AMD、INTC、NOW | +1 |
| 6~8 | MSFT、MU、COST、PANW、WDC | −1 |
| 3 / 5 | QRVO、ORCL | −2 |
| 1 | NVDA、CRM | −3 |

INTC `20260723` 標成 `FY2026Q3`、實際 FY2026 Q2 —— 與 TODO 裡手動查到的一致。

順手查到第二個問題，比標籤本身更嚴重：dedupe 的「同標籤保留最舊那份」撞到
2 次，**兩次都保留了錯的那份**——

- **WDC `FY2025Q1`**：保留 2025-01-10 那份（Item 2.02+5.02，**根本沒有新聞稿
  附件**），丟掉 2025-01-29 的正式財報 → **整季靜默消失**
- **QRVO `FY2025Q4`**：保留 2025-10-28 的「**Preliminary** ... Second Quarter
  Results」，丟掉 11-03 的正式財報 → **拿到初步數字**

方法上用兩個獨立來源交叉驗證（公司自述財期 vs 期末日換算），一致 97 /
不一致 7；不一致全部是 COST/WDC/PANW 這種 52/53 週制的公司——財年結束日會在
月底前後跳，看月份換算就差一季。**這是修的時候的地雷：不能用期末日的月份
推財季。**

沒有動任何程式。修法會改到 `nongaap_cache.json` 的 key、要重抓，三個選項與
代價寫在報告末尾。目前影響是**潛伏**的（`NONGAAP_ENABLED = False`），唯一
對外吐季度標籤的 `cli.py press-release` 每一季都帶 `label_warning`。

---

### 2026-08-07（早上）

**GAAP 路徑徹底移除 AI；Non-GAAP 從源頭停用避免白燒額度**

- `override_engine.E2_LLM_ENABLED = False`：GAAP 抓取不再呼叫 AI，即使 GUI
  照舊把 `ai_config` 傳進來也不會真的打 API。E1 模糊比對照常運作，找不到就
  警告，不叫 AI 猜。實測傳真實 `ai_config` 抓 COHR，AI 呼叫次數 0
- `main.NONGAAP_ENABLED = False`：兩個 GUI checkbox 停用並改標「暫停中，
  改由 skill 處理」。**差點犯的錯**：一開始只在輸出端過濾掉 `Data_NonGAAP`，
  但 checkbox 還能勾——會照常呼叫 AI 抓完 6 季**才被丟掉**，等於白燒額度；
  停用要停在源頭，抓取路徑本身也加了守衛
- 相關程式碼（`nongaap_layout` / `metric_rules` / 快取）全部保留，改回
  `True` 即可恢復
- 補一條測試釘住「預設不打 AI API」，避免旗標被改回去沒人發現

---

### 2026-08-02（晚間）

**新聞稿截斷 12,000 字元——Non-GAAP 調節表抓不到的真正原因**

實跑 ARLO 重抓後發現 GAAP 對照行有值、但調節表七項全空。查證後確認是 `_call_ai()` 的 `text[:12000]` 造成的：

- ARLO 新聞稿全長 **53,569 字元，prompt 只送前 12,000（22%）**
- 「Stock-based compensation」出現在 18,605 / 33,759 / 38,440 / 40,558，「Amortization」在 40,848——**全部在截斷之後，AI 根本沒看到調節表**
- 重點條列都在文件最前面（所以毛利率、EPS 一直抓得到），但調節表一律在文件尾端

`PROMPT_TEXT_LIMIT` 由 12,000 提高到 200,000（約 50K token，Gemini Flash context 為 100 萬 token）。這是舊 prompt 時代留下的保守值，不是有意的設計。

**ARLO 重抓前後對照**（密度 38/210 → **64/175**）：

| 列 | 修前 | 修後 |
|---|---|---|
| GAAP Gross Margin | 0/6 | **5/5** |
| GAAP Net Income | 1/6 | **5/5** |
| 股權獎酬 SBC | 0/6 | **5/5** |
| Free Cash Flow | 2/6 | **5/5** |

殘差驗證有效：最新一季 GAAP 淨利 14.88M + SBC 19.73M = 34.61M，Non-GAAP 淨利 30.96M，殘差 **−3.65M**（未具名的稅務影響等）——調節橋沒對平時會自己顯示出來。

**期末流通股數改走封面頁 dei fact**

原方案（BS 模板對映 `us-gaap:CommonStockSharesOutstanding`）實測 ARLO/AAPL/NVDA/MSFT/COHR 五家全部沒有 tag，股數只寫在 `CommonStockValue` 的 label 文字裡。改走 `Company.get_facts()` 的 `dei:EntityCommonStockSharesOutstanding`：

- 歷史序列完整（ARLO 32 筆、AAPL 70 筆，2009 年起逐季）
- fact 的 `fiscal_year` + `fiscal_period` 直接對得上本專案的 `FY{year}Q{n}` 標籤
- ⚠ 該 fact 的日期是封面頁「最近可行日期」，比財季結束**晚幾週**（ARLO FY2025Q1 財季結束 2025-03-30，股數是 2025-05-02 的 103,400,957）。這是公開資料裡最接近的時點股數，但不是財季結束當天
- 實跑驗證：ARLO 103.4M → 108.6M；AAPL 14.94B → 14.59B（回購使股數逐季下降，`流通股數 YoY` −1.66%）。`BVPS` 一併有值

**`Data_Std` 財季改用 FY/FQ 標記**

第 3 列由 `2026Q1` 改為 `FY2026FQ1`，與第 4 列的日曆季 `2026Q1` 在視覺上分得開。這兩列最容易被搞混，非 12 月結算的公司同一欄可能是 `FY2026FQ1` 但日曆 `2025Q4`，看錯就是整整一季的誤差。

- 另驗證 AAPL 37 個比率只有 1 個全空（利息保障倍數，該公司未 tag 利息費用），TTM 類的 ROE／ROA 都算得出來
- 測試 492 → **505**

---

### 2026-08-02（下午）

**`Data_Std`：跨公司固定版面表**

使用者回報「每間公司的 sheet 數量名稱都不同，很難用公式對照不同公司」。實測 `output/` 既有 11 個檔案後確認**三個軸同時在變**：sheet 數 10～30 張、`Cash` 落在第 28～56 列（overflow 行插在 section 之間把後面整段推移）、季度欄 4～50。唯一能跨公司直接參照的只有 `C4`。`VLOOKUP` 也不安全——`Net Income` 在 IS 與 CF 各出現一次。

- **`std_sheet.py`（新）**：`Data_Std` 三個保證——固定 sheet 名稱、固定列位（overflow 一律不進來）、固定機器鍵（B 欄 `IS.REVENUE` / `BS.CASH` / `CF.NET_INCOME` / `RATIO.毛利率`）。136 列，內容為 IS 22 + BS 42 + CF 26 + 比率 37 + 表頭 3
- **`FROZEN_ROW_NUMBERS` 列號凍結測試**：列號寫死在測試裡，任何人插入一列都會立刻紅。沒有這條，「固定列位」幾個月後會悄悄失效，而使用者的跨檔公式會靜默抓錯
- **兩種期間標籤**：查證後發現現行 `FY2026Q1` 是**公司財季不是日曆季**——`_col_to_quarter_label()` 對非 12 月結算的公司會把年份往後推。ARLO 的 FY2025Q1 是 2025 年 3 月、AAPL 的是 2024 年 12 月、NVDA 的是 2024 年 4 月，模板照財季標籤對齊會靜默比錯期間。`Data_Std` 因此加了第 3 列「日曆季」與第 4 列「期末年月」，跨公司對齊用第 3 列
- **`fetcher_gaap.py`**：`Data_Meta` 新增 `Fiscal Year End Month`（日曆季換算的依據）
- **實機驗收**：ARLO（12 月結算）與 AAPL（9 月結算）各跑一次，兩家的 `Data_Std` 都是 136 列、`IS.REVENUE` 都在第 7 列、`BS.CASH` 第 30 列、`CF.FREE_CASH_FLOW` 第 98 列、`RATIO.毛利率` 第 108 列。AAPL 的 `FY2025Q1` 正確標成日曆 `2024Q4`
- README 補「用 `Data_Std` 寫跨公司模板」章節，含 `INDEX`+`MATCH` 範例與財季/日曆季的陷阱說明
- **三種期間標籤（同日修正）**：使用者澄清季度標籤要照**公司財年**算、同時要有**日曆年的結算日**。因此第 3 列為財季（`2026Q1`，去掉 FY 前綴方便比對）、第 4 列為日曆季、第 5 列為**真實期末結算日**
  - 期末日原本是從財季標籤 + 結算月反推的，只到月份。美股多用 52/53 週制，ARLO 的 FY2025Q1 實際結束在 `2025-03-30` 而非 03-31、AAPL 的 FY2026Q1 在 `2025-12-27`。真實日期在 XBRL 欄名（`"2026-03-29 (Q1)"`）裡，`_col_to_quarter_label()` 解析完就丟掉了
  - `StatementTable` 新增 `period_ends` 欄位（有預設值，向後相容），`_build_is_table()` 記錄、`_merge_financials()` 併入、`Data_Std` 逐欄使用；沒帶到的欄位自動退回反推年月
- 測試 464 → **492**

---

### 2026-08-02

**`Data_NonGAAP` 固定模板、`Data_Ratios` 比率表、`Data_Segments` 長格式**

- **`scripts/survey_nongaap_metrics.py`**：調查 32 家（大中小型跨產業）8-K 新聞稿實際使用的 Non-GAAP 指標，純文字比對不呼叫 AI。結果決定了 Core 收哪些行：Net Income 79%、Diluted EPS 79%、FCF 76%、Operating Income 66%、Effective Tax Rate 62%、Gross Margin 59%、**Net Margin 0%**（沒有任何一家會寫，只能推導）。完全不報 Non-GAAP：AAPL / AMZN / COST
- **`nongaap_layout.py`（新）**：`Data_NonGAAP` 改為「固定模板 + overflow」，沿用 `Data_Financials` 已驗證的模式。15 行 Core（永遠存在，沒資料就空白）+ GAAP 對照行 + GAAP→Non-GAAP 調節表 + overflow 區 + 年度 (FY) 區。SaaS 專屬指標（ARR/RPO/Billings/NRR）刻意不進 Core——收了等於開產業別模板
  - **GAAP 對照行從同一份新聞稿抓，不從 `Data_Financials` 拉**：Non-GAAP 的季度標籤系統性晚一季，跨表拉會變成錯開一季的無聲比較。8-K 調節表本來就同時列 GAAP 與 Non-GAAP
  - **調節表的「其他」用殘差倒算**：`Non-GAAP 淨利 − GAAP 淨利 − 具名項目合計`。表因此會自己對帳——AI 漏抓某個調整項時殘差會變大，一眼看得出來
  - 具名調整項取實測覆蓋率 ≥59% 的七項：SBC 90%、重組資遣 79%、減損 69%、訴訟和解 66%、無形資產攤銷 66%、併購相關 62%、調整項稅務影響 59%
  - prompt 同步要求 AI 抓 GAAP 對照值與調節項目（帶號，加到 GAAP 淨利以得到 Non-GAAP 淨利）
- **`ratios.py`（新）**：`Data_Ratios` 37 個比率。寫算好的**值**不寫 Excel 公式（公式結果要 Excel 開過才寫進檔案，openpyxl 直接讀會拿到 None），B 欄寫算法文字，列名帶單位後綴 `(%)` / `(x)` / `(days)` / `($)`——後綴優先於關鍵字判斷，否則「流動比率 (x)」含「率」會被 ÷100、「DSO (days)」會被當金額 ÷1,000,000
  - **實跑抓到的嚴重 bug**：季度序列有缺口時（ARLO 實際是 FY2024Q1/Q2/Q3 → FY2025Q1，缺 Q4），YoY/QoQ/TTM 用「往前數 N 格」會取到錯誤基期。營收 `[100,200,300,400,260]` 的 FY2025Q2，正確 YoY 對 FY2024Q2 = +30%，位置法對到 FY2024Q1 = **+160%**，而且看起來完全正常。已改為依季度標籤對齊（`_lag_index`），基期不存在就留空
  - ROE = TTM 淨利 ÷ 期初期末平均權益；成長率基期 ≤ 0 時回 None
- **`segments.py`（新）**：`Data_Segments` 長格式，把所有 `Data_Seg_*` 併成固定名稱固定欄位的一張表，欄位取各軸季度聯集並依標籤對位。寬格式原樣保留，兩者同源
- **`fetcher_gaap.py`**：BS 模板新增 `Shares Outstanding`（41 → 42 行）
- 測試 423 → **464**

**待解**：期末流通股數實測 ARLO/AAPL/NVDA/MSFT/COHR 五家皆未在資產負債表 tag `CommonStockSharesOutstanding`，需改走封面頁 `dei:EntityCommonStockSharesOutstanding`（TODO 第 3 項）。GAAP 對照行與調節項目在既有快取上是空的——那些快取用舊 prompt 抓的，重抓後才會填上。

---

### 2026-08-01（晚間）

**輸出檔覆蓋防護 + max_filings 回補後重新裁切（TODO 第 8、4 項）**

- 新增 `check_output_writable()`，抓取開始前偵測檔案是否被 Excel 鎖住。
  原本失敗點在最後一步 `wb.save()`，使用者要白等一分多鐘才看到
  `PermissionError`；single/batch 兩條路徑都接上，批次模式只跳過該家
- `write_statements()` 改為寫暫存檔再 `os.replace()`，save 中途失敗時原檔
  完好；覆蓋既有檔前留一份滾動備份 `.bak.xlsx`（`KEEP_BACKUP` 可關）——
  年份範圍變窄時舊季度會整批消失，這是唯一的後路
- `_recover_missing_quarters()` 補回缺季後重新套用 `max_filings` 裁切。
  原本要 4 季、保留區間有 2 個缺口會實際處理 6 季，每多一季多一次 AI
  呼叫；裁切保留最新的
- 新增 `scripts/survey_nongaap_metrics.py`：調查 32 家公司 8-K 新聞稿實際
  使用的 Non-GAAP 指標，純文字比對不呼叫 AI，供決定 `Data_NonGAAP` 固定模板
- 測試 358 → 372

---

### 2026-08-01（下午）

**`Data_NonGAAP` 資料品質修復（TODO 第 2 項）——整張 sheet 由不可用變為可用**

採方案 (c)：prompt 改英文 **＋** 保留中英對照層。實查 `nongaap_cache.json` 後才確定不能只做其中一半——**AI 回中文還是英文是隨機的，同一個 ticker 內都會混**（CRM `FY2026Q2` 回中文、`FY2026Q1` 回英文），所以只改 prompt 沒有防線，只補中文規則則因為混語仍然合併不起來。

- **新增 `metric_rules.py`（唯一可調整處）**：期間 token 樣式（中英）、guidance 詞表（中英）、中文→英文詞彙表、同義名合併表、Excel 數值分類關鍵字，全部集中於此。規則作用在**讀取快取**階段而非寫入階段，因此改表後重跑即生效，**不需要重抓 8-K、不需要刪快取**
- **`fetcher_nongaap.py`**：
  - `_clean_metric_name()` 加剝中文期間 token（`2024年第四季` / `2025年第四季度` / `2026財年第三季度` / `第一季` / `2024全年度` / `2025年全年度`）
  - guidance 過濾中文用「包含」比對而非 `startswith`——中文的 guidance 詞常在名稱中間（`2026財年預期 Non-GAAP 營業利潤率上限`），沿用英文的 `startswith` 會整批漏掉
  - 新增 `_canonicalize_metric_name()`（詞彙替換可組合：`自由現金流`+`利潤率` → `Free Cash Flow Margin`，長詞優先）與 `_metric_merge_key()`（忽略大小寫與標點）
  - `_build_nongaap_table()` 改為讀取時重跑正規化＋對照＋跨季合併，既有中文快取因此自動救回
  - prompt 改英文並明確要求英文指標名、排除 guidance、**只取當期**（見下方 FCF）
- **`excel_formatter.py`**：新增 `FMT_PERCENT` 與百分比類別（÷1M 豁免）。分類順序為每股 → 百分比 → 股數 → 金額，關鍵字表讀 `metric_rules.py`
- **實作中查出並修掉的第三個缺陷**：關鍵字若用裸子字串比對，`Operations` 含 `ratio`、`Corporate` 含 `rate`、`Steps` 含 `eps`——XBRL overflow 行的 label 常含 `Operations`，會被誤判成百分比而**不再除以 1M**，三表金額直接錯 6 個數量級。ASCII 關鍵字改為一律要求詞界，並補 4 條迴歸測試釘住
- **實跑中查出並修掉的第四個缺陷**：`Free Cash Flow` 列混進全年度／TTM 數字（ARLO `FY2025Q1` 的 48.6M 配 9.5% margin 是年度值，單季應約 37%）。prompt 補「只取當期」約束後，該列只剩真正的單季值——**密度由 40/48 降為 32/42，但稀疏且正確優於密實且錯誤**
- **測試**：新增 51 個單元測試，總數 **276 → 327**（`pytest tests/ --ignore=tests/test_live_snapshots.py`），既有 276 條零轉紅。含 `tests/fixtures/arlo_nongaap_raw.json` golden test——直接用修復前 AI 真實吐出的髒中文快取當輸入（含 guidance、LTM、混語），不連網不呼叫 AI
- **實機驗收（ARLO 2025–2026，6 季）**：
  - 舊中文快取路徑（不呼叫 AI）→ 8 列全部合併成功，`Non-GAAP Gross Margin` 與 `Non-GAAP Diluted EPS` 皆 6/6 密實
  - 刪快取重抓（新英文 prompt）→ 快取 100% 英文、無 guidance 列、無期間 token；兩條路徑數值完全一致
  - 對 8-K 原文（`2026-05-07` 申報）逐項核對全中：非 GAAP 毛利率 50.1%、EPS $0.28、Adjusted EBITDA $30.4M／margin 20.2%、訂閱與服務毛利率 85.4%、FCF $25.4M／margin 16.9%
  - Excel 實檔驗證：毛利率顯示 `50.1%`（修前 `3.75e-05`）、EPS `0.28`（修前 `1e-07`）、Adjusted EBITDA `30.4`（正確 ÷1M）
**同日追加：跨公司驗證（CRM／PANW）與快取污染修復**

拿 CRM、PANW 的實際輸出回頭驗規則表，又抓到兩個獨立缺陷：

- **AI 呼叫失敗會污染快取（嚴重，資料永久遺失）**：`_call_ai()` 失敗時回 `{}`，與「AI 有回應但新聞稿真的沒有 Non-GAAP 指標」無法區分，該季照樣寫進 `nongaap_cache.json`；下次執行 `lbl not in cache` 命中，**那季永遠不會再被抓，且全程無聲**。實跑 PANW 時撞到 Gemini `HTTP 429`，6 季有 2 季就這樣被寫成空白。已改為失敗回 `None`、不寫快取、下次執行自動重試，並在 stderr 明示；「真的沒有指標」仍照常寫快取以免重複付費呼叫 AI。補 6 條測試（含「第一趟失敗、第二趟成功要真的重抓」）
- **英文 guidance 詞在名稱中間時漏掉**：英文原本只用 `startswith`，CRM 的 `Non-GAAP Diluted Net Income Per Share Guidance (Low)` 抓不到，預測數字直接混進時間序列。已加 `GUIDANCE_SUBSTRINGS_EN`（`guidance` / `forecast` / `(low)` / `(high)`）
- **規則表擴充**：CRM 的 SaaS 用語（`恆定匯率`／`固定匯率` → Constant Currency、`當期剩餘履約義務` → cRPO、`成長率` → Growth、`年增率` → YoY Growth、`與支援` → and Support），並把「恆定匯率」與「固定匯率」兩種說法併成同一列。`PERCENT_KEYWORDS` 的 `growth %` 放寬為 `growth`
- 測試 327 → **342**

**同日追加（二）：依「工具只負責照實落地 8-K 數字，不做判斷」的原則調整四項**

- **年度值不再填進季欄位**（`metric_rules.FY_ONLY_HANDLING = "label"`）：舊行為是當季缺值時把全年數字填進該季欄位且無標記——那是替資料下判斷。新行為是另成一列、名稱加 ` (FY)`。第三種選項 `"drop"`（直接丟）也保留，但預設不用，因為丟掉等於刪掉 8-K 裡確實存在的資料
- **百分比改存 Excel 原生比例**（`PERCENT_AS_EXCEL_RATIO = True`）：37.5 存為 0.375、格式 `0.0%`、顯示 37.5%。在 Excel 拉公式與畫圖直接可用
- **`服務毛利率` 與 `訂閱與服務毛利率` 不再合併**：查 ARLO 原文確認公司在 2025 年改過名——`FY2025Q1` 報 "non-GAAP service gross margin 81.7%"（配 Service revenue $64.1M），`FY2025Q2` 起改成 "subscriptions and services gross margin 83.1%"（配 Subscriptions and services revenue $68.8M），名稱與營收基礎同時變動。認定改名前後是同一條線屬於判斷，改為兩列並陳。單複數／連接詞差異仍合併（那是寫法不是定義）
- **AI 呼叫加退避重試 + 跑完統計**：`_ai_request()` 抽出成獨立接縫，`_call_ai()` 包重試（`AI_MAX_ATTEMPTS = 3`、退避 5s／15s）。次數壓低是因為 Gemini 的每日配額型 429 重試必敗、只是白等；每分鐘限流型則有效。跑完會列出未取得的季度並推給 `progress_cb`（GUI 使用者看不到 stderr）
- **查證紀錄（回答「為何會漏資料」）**：ARLO `FY2025Q1` 缺 Adjusted EBITDA，翻原文確認**數字在 8-K 裡**（`Adjusted EBITDA $9,765 / margin 8.0%`），是 AI 沒抓到——那一季 ARLO 只把它放在後段調節表，沒寫進前面的重點條列。屬 AI 抽取召回率問題，歸 TODO 第 3 項處理；補洞邏輯救不了（只會拿年度值去蓋）。GAAP 走 XBRL，完全不受影響
- 測試 342 → **358**

- **本次自行決定、可調整處**（都寫在程式碼註解裡）：
  - 百分比存**原始數字** 37.5 而非 Excel 比例 0.375 → `excel_formatter.PERCENT_AS_EXCEL_RATIO = False`
  - `服務毛利率` 與 `訂閱與服務毛利率` 視為同一列 → `metric_rules.METRIC_ALIASES` 刪一行即可分開
  - 對照表未收錄的名稱**原樣通過**，不丟棄

---

### 2026-08-01

**GUI 實機驗收（ARLO）與既有缺陷紀錄**

分支 `feat/8k-scan-optimization` 已以 `--no-ff` 合併回 master。

- **驗收結果**：從 `啟動器.bat` 實跑 ARLO（年份 2025–2026，GAAP + Non-GAAP），一分鐘內完成。GAAP 取得 8 份財報，`Index` 品質欄 **9/9 ✓**；Non-GAAP 6 季中僅 2 季需重新呼叫 AI，其餘 4 季由 `nongaap_cache.json` 命中——快取行為正確
- **`Data_Financials(Q)` 120 行中 31 行全空屬正常**：3 行為 IS/BS/CF 分隔標題，其餘為 ARLO 實際沒有的項目（無有息負債、不配息、無少數股權）
- **`Data_NonGAAP` 目前不可用**（既有缺陷，非本次改動引入，已列入 TODO 第 2 項）：
  - 數值被誤除以 1,000,000——`excel_formatter.py:49` 的 ÷1M 豁免關鍵字只有 `EPS` / `Per Share` / `per share`，AI 回的中文指標名（「Non-GAAP 每股盈餘」「Non-GAAP 毛利率」）比對不到。實例：毛利率 37.5 → `3.75e-05`、EPS 0.10 → `1e-07`
  - 指標名稱帶期間 token 未被剝除——`_normalize_nongaap_metrics()` 只處理 `Q4 FY26` / `FY2026` 等英文格式，中文的「2024年第四季」「2024全年度」不認得，導致同一指標每季各自成行，整張表對角線散開
  - 根因同一個：AI prompt 以中文撰寫，AI 回中文指標名，但下游正規化與格式化規則皆只認英文
- **重複抓取同一 ticker 的檔案行為**（已列入 TODO 第 8 項）：`write_statements()` 開啟既有檔後刪除所有 `Data_*` sheet 重寫，`My_*` 等自訂 sheet 保留，無備份無版本號。年份範圍變窄時舊季度直接消失；檔案被 Excel 開啟時在最後一步才拋 `PermissionError`

---

### 2026-07-31

**8-K 掃描效率優化**

設計文件：`docs/superpowers/specs/2026-07-31-8k-scan-optimization-design.md`

- **`fetcher_nongaap.py`**：
  - 新增 `_list_earnings_filings()`：改在 SEC 申報清單階段以 `items`（Item 2.02）與 `period_of_report` 完成篩選、去重、年份過濾與 `max_filings` 切割，全程不下載檔案
  - 新增 `_quarter_ordinal()` / `_ordinal_to_quarter()` / `_find_missing_quarters()`：偵測季度序列缺口
  - 新增 `_recover_missing_quarters()`：僅對缺季區間回退逐筆 `obj()` 深掃，用 `has_earnings` 找回未標 2.02 的財報；補不到的季度寫 stderr
  - 移除 `_get_earnings_filings()`（對全部歷史 8-K 逐筆下載）
  - `fetch_nongaap_statements()`：`obj()` 移進迴圈，只對未快取的季度下載；年份過濾改在 `max_filings` 之前套用
- **實測**：AAPL 全部 8-K 235 份、含 2.02 者 94 份；抓 4 季的下載次數由 235 降至 4
- **已知邊界**：SEC 自 2004-08-23 才啟用 2.02 編號，更早的財報 8-K（Item 12/5）不會被抓到
- **實機驗收（Task 5，`nongaap_cache.json` 未命中、逐季即時 AI 呼叫）**：CRM 76.5 秒（`~/.edgar` 已預熱）、PANW 71.3 秒（`~/.edgar` 4 份中 3 份已預熱）、ARLO 76.0 秒（兩層快取皆全冷），皆為 `max_filings=4`（4 季 = 4 次即時 AI 呼叫）。此結果超出設計文件原估的 60 秒目標，原因是 AI 呼叫本身耗時已主導總時間、非下載次數；但相較優化前單一 ticker 需 5–10 分鐘，仍是數量級改善
- **健壯性修正（最終審查後追加）**：`_recover_missing_quarters()` 原本未保護 `_period_to_quarter_label()` 呼叫，一份 `period_of_report` 為 8 字元以上非數字的申報即拋 `ValueError`，往上穿出 `fetch_nongaap_statements()` 只在 GUI 層被接住 → **整趟輸出失敗、連同一輪已抓好的 GAAP 三表一併拿不到 Excel**。缺季回補掃的正是「未標 2.02」這批 metadata 最雜亂的申報，暴露程度高於清單路徑。已補上與 `_list_earnings_filings()` 同款的 guard（例外只記 `type(exc).__name__` + `_exc_status(exc)`），畸形申報改為跳過
- **`tests/test_fetcher_nongaap.py`**：新增 26 個單元測試（清單篩選 8、缺季偵測 7、缺口補掃 5、下載時機 2、review 修正回合追加 2：`test_fetch_nongaap_recovers_gap_quarter_in_newest_first_order`、`test_list_earnings_filings_skips_non_numeric_period`、最終審查追加 2：缺季回補的畸形 `period_of_report` 防護），單元測試總數由 250 增至 276（`pytest tests/ --ignore=tests/test_live_snapshots.py`）
- 另在 `tests/test_live_snapshots.py` 新增 2 個 `slow` 連網驗收測試，不含在上述 250/276 計數內，預設指令不會執行。`test_live_listing_filter_matches_deep_scan_arlo` 另加 `len(deep_labels) >= 20` 下限斷言——該測試迴圈內有 `except Exception: continue`，若下載全面失敗（限速、identity 被拒）`deep_labels` 會是空集合而讓斷言假性通過，而它是全分支唯一守住「快速路徑不漏季」這個核心主張的測試

**本次順帶查出的既有缺陷（未修，已寫入 TODO 第 2 項）**

最終審查實查 EDGAR 後確認：Item 2.02 8-K 的 `period_of_report` 存的是**發布日**而非財期結束日，故 `Data_NonGAAP` 每一欄的季度標籤都比數字實際所屬財季晚約一季（INTC `20260723`→`FY2026Q3` 實報 FY2026 Q2；COST `20260528`→`FY2026Q2` 實報 FY2026 Q3）。同一根因另會漏抓：WDC 於 2025 日曆 Q1 發過兩份 Item 2.02 8-K（`20250110`、`20250129`），同標 `FY2025Q1`，去重「留最舊」把 1/29 那份丟掉。兩者皆長期存在、非本次引入（舊 `_get_earnings_filings()` 讀同一欄位）。

**已知殘留（parked）**

`_recover_missing_quarters()` 內針對 `_quarter_ordinal()` 回傳 `None` 的排序 guard，經複審實測為不可達路徑——既有的 `label not in wanted` 檢查已先濾掉畸形標籤，把程式碼退回修正前跑同一測試同樣通過。guard 本身無害，但對應測試是空測試，下次動到該檔時一併清除。

### 2026-06-10
- 修正：`winget install Python` 加入 `--override "/quiet PrependPath=1 Include_pip=1"`，確保靜默安裝後 Python 自動加進 PATH
- 修正：`launcher.ps1` 加入全域 `trap`，攔截未處理例外，防止執行失敗時視窗直接閃退

### 2026-05-05（Session 18）

**Index Sheet 品質檢測**

設計文件：`docs/superpowers/specs/2026-05-03-index-quality-check-design.md`

- **`excel_formatter.py`**：
  - 新增 4 個顏色常數：`QUALITY_GREEN / ORANGE / MISS_BG / MISS_FG`
  - 新增 `ALL_KEY_ROWS`（9 個 key rows：IS 4 + BS 3 + CF 2）
  - 新增 `_compute_quality(tables)` helper：找 `Data_Financials(Q)`，呼叫 `check_key_rows` for IS/BS/CF，回傳 `(score, total, missing_set)` 或 `None`
  - `_build_index_sheet()` 更新：
    - Header 合併範圍 `A1:D1` → `A1:E1`（同 A2）
    - Row 4 新增第 5 欄「完成度」
    - 每個 sheet 列的 E 欄：`Data_Financials(Q)` 顯示 `"9/9 ✓"`（綠）或 `"N/9 ⚠"`（橘），其餘顯示「—」
    - 表格下方加「品質明細」區塊：section header + 9 行逐一顯示 ✓ / ✗，缺失行底色淺橘
- **`tests/test_excel_formatter.py`**：新增 14 個 tests（`_compute_quality` × 4、E 欄 × 5、明細區塊 × 5）

**全套測試：250/250 PASSED**

**Commits**

```
ecddafd  refactor: clean up unused unpacking and duplicate import
f9aab47  refactor: use QUALITY_GREEN constant instead of hardcoded colour in detail loop
da128bb  feat: add quality detail section to Index sheet
b5714d0  feat: add quality score column E to Index sheet
b82de46  refactor: move override_engine import to top-level; import ALL_KEY_ROWS in tests
ee19aa1  feat: add _compute_quality helper for Index sheet quality check
```

---

### 2026-05-03（Session 17）

**日期區間反轉驗證**

- **`main.py`**：
  - `_run_single()`：解析 `start_year` / `end_year` 後新增 guard，`start_year > end_year` 時彈錯誤訊息並中止，訊息帶實際數值（如「起始年份（2025）不可大於結束年份（2020）」）
  - `_run_batch()`：同上

**全套測試：236/236 PASSED**

**Commits**

```
385a67c  fix: guard against start_year > end_year in Tab 1 and Tab 2
```

---

### 2026-04-29（Session 16）

**UI 日期區間 + 報表類型 + Sheet 預覽 + FY Month Fix**

設計文件：`docs/superpowers/specs/2026-04-29-ui-date-range-design.md`

**後端（`fetcher_gaap.py`）**

- **`_filter_filings_by_year()`**（新 helper）：依起訖年份過濾 filing 列表，支援 `date` 物件與 ISO 字串兩種格式
- **`fetch_gaap_statements()` 新增 5 個參數**：
  - `start_year` / `end_year`：只抓指定年份區間的 filings（`None` = 全部，現有行為不變）
  - `fetch_quarterly` / `fetch_annual`：控制是否抓 10-Q / 10-K（預設兩者皆抓）
  - `excluded_sheets`：跳過指定 sheet 名稱，不寫入 Excel
- **`preview_sheets(ticker, identity)`**（新公開函式）：只抓最新一份 10-Q，回傳預期 sheet 名稱清單（~5–15 秒），用於執行前快速預覽
- **FY month fix**：當 `fetch_annual=False` 時，後端自動偷抓 1 份 10-K 僅用於偵測財年結束月份，不產生 `Data_Financials(Y)`；確保 AAPL（9 月）等非 12 月財年公司的季報欄位標籤正確

**後端（`fetcher_nongaap.py`）**

- **`_filter_nongaap_by_year()`**（新 helper）：依 label 字串中的年份（如 `FY2021Q2`）過濾
- **`fetch_nongaap_statements()` 新增 `start_year` / `end_year` 參數**

**UI（`main.py`）**

- **Tab 1（單一公司）**：
  - 「快速掃描 ▶」按鈕（Row 0）：在執行前偵測該 ticker 有哪些 `Data_Seg_*` sheets；固定 sheet（Financials/Meta）不可取消
  - 「▶ 進階設定」收折區（預設隱藏）：展開後可個別勾選季報（10-Q）/ 年報（10-K）
  - 日期區間 Spinbox（起 / 迄年份，留空 = 全部）
- **Tab 2（批量更新）**：
  - 「▶ 進階設定」收折區：同上，可選報表類型
  - 日期區間 Spinbox

**測試（`tests/`）**

- `test_fetcher_gaap.py`：新增 14 個測試（filter、fetch_gaap 新參數、preview_sheets、FY probe）→ 共 **110 tests**
- `test_fetcher_nongaap.py`：新增 4 個 filter 測試 → 共 **41 tests**
- 全套 **151/151 PASSED**

**Commits（2026-04-29）**

```
555193d  feat: move report type to inline adv settings; fix FY month probe for Q-only mode
adf2336  fix: guard scan against running fetch; hide panel on empty sheet list
5dd7cdd  feat: add quick scan button and sheet selection panel to Tab 1
63cd654  fix: align Tab 2 row sticky to ew matching Tab 1 pattern
bbb7d43  feat: add report type and date range to Tab 2 batch UI
d15f1d9  feat: add report type (10-Q/10-K) and date range to Tab 1 UI
ab9d697  feat: add start_year/end_year filter to fetch_nongaap_statements
a9caeec  feat: add preview_sheets() for quick segment detection
016845b  feat: add start_year/end_year/fetch_quarterly/fetch_annual/excluded_sheets to fetch_gaap_statements
9260596  feat: add _filter_filings_by_year helper with year-range support
```

---

### 2026-04-28（Session 15）

**GAAP Fetcher 精度修復（Tasks 1–8）+ Batch Smoke Test**

設計文件：`docs/superpowers/plans/2026-04-26-fetcher-gaap-fixes.md`

- **Task 1：pre-XBRL early exit** — IS/BS/CF/Seg 四個 filing loop 在 `filing_date < 2008-01-01` 時 `break`，AMD 等老公司抓取速度大幅改善
- **Task 2：Dividends Paid bug** — CF_TEMPLATE Dividends Paid 的 `std_concept` 從 `DistributionsToMinorityInterests`（錯誤）改為 `None`，fallback regex 正確命中
- **Task 3：Net Income fallback 順序** — IS 先嘗試 `NetIncomeLossAttributableToParent`，再 `ProfitLoss`，避免少數股東損益被誤計入母公司 Net Income
- **Task 4：Revenue fallback 擴展** — 新增 `Revenues`、`SalesRevenueNet`、`SalesRevenueGoodsNet`，涵蓋更多公司的 XBRL 命名
- **Task 5：Total Non-op DERIVED guard** — 當 discontinued operations 行存在時跳過衍生計算，避免雙重計入
- **Task 6：Investment Proceeds 多概念加總** — 新增 `_sum_matching_rows()` helper，將 `ProceedsFromSaleOfInvestments`、`ProceedsFromMaturitiesOfInvestments` 等多個概念加總，而非只取第一筆
- **Task 7：Debt Proceeds/Repayments 多概念加總** — 同上，LT 債 + ST 債分別加總，不再只抓第一個命中的概念
- **Task 8：FY Label 對齊公司財年** — 新增 `_detect_fy_end_month(filings_k)` 從 10-K 推算財年結束月份；`_col_to_quarter_label()` 根據 `fy_end_month` 調整 Q 序號，確保 AAPL（Sep）、MSFT（Jun）、NVDA/WMT（Jan）的季度標籤正確

**`tests/test_fetcher_gaap.py`**：新增對應 Tasks 1–8 的 unit tests；**95/95 全數通過**

---

**Batch Live Smoke Test（新腳本）**

- **`scripts/smoke_test_10.py`（新檔）**：
  - 對 10 間公司（AAPL/MSFT/TSLA/AMD/NVDA/GOOGL/META/WMT/COHR/AMZN）呼叫真實 EDGAR API
  - 自動檢查 7 個 key rows：Revenue / Gross Profit / Operating Income / Net Income / Operating Cash Flow / Capex / Free Cash Flow
  - 輸出每間公司的 OK/NONE 狀態 + 末尾彙總表，列出有問題的 ticker
  - Windows cp950 encoding 修復：`sys.stdout = io.TextIOWrapper(..., encoding="utf-8")`

- **`scripts/README.md`（新檔）**：腳本索引，符合 CLAUDE.md 規則（新增腳本必須更新此表）

**實機測試結果（2026-04-28）：9/10 OK**
- AAPL/MSFT/TSLA/AMD/NVDA/GOOGL/META/WMT/AMZN：7 個 key rows 全部 ✅
- COHR：Revenue=NONE、Operating Income=NONE — 已知限制（2022 三方合併導致非標準 XBRL）
- FY label 驗證：AAPL FY2026Q1 ✅ / MSFT FY2026Q2 ✅ / NVDA FY2026Q3 ✅ / WMT FY2026Q3 ✅

**所有 commits 已 push 至 GitHub remote**（`0508dc6`）

---

### 2026-04-26（Session 14）

**CF YTD Overflow 修復 + 測試套件（COHR/LITE/AAPL/NVDA/GOOGL）**

- **`fetcher_gaap.py`**：
  - `_build_cf_table`：移除 `if not is_ytd: _collect_overflow(...)` 限制
  - 新增 `overflow_per_filing: dict[str, dict]` — 對所有 filing（含 YTD）收集原始 overflow 值
  - Filing loop 結束後，與模板行相同的跨 filing 減法計算出 Q2/Q3 standalone overflow
  - Q2 overflow standalone = Q2_YTD_overflow − Q1_overflow；Q3 = Q3_YTD − Q2_YTD
  - 若前一季無對應 concept，Q2/Q3 overflow 保持 None（保守策略，與模板行一致）

- **`tests/test_fetcher_gaap.py`**：
  - 新增 4 個 CF overflow YTD unit tests：
    - `test_cf_overflow_q1_standalone_captured` — Q1 overflow 正常收集
    - `test_cf_overflow_q2_ytd_subtracted` — Q2 overflow = Q2_YTD − Q1
    - `test_cf_overflow_q3_ytd_subtracted` — Q3 overflow = Q3_YTD − Q2_YTD
    - `test_cf_overflow_q2_without_q1_is_none` — 前一季缺失時結果為 None
  - **73/73 unit tests 全數通過**

- **`tests/test_live_snapshots.py`**：
  - 新增 `CF_OVERFLOW_TICKERS = ["COHR", "LITE", "AAPL", "NVDA", "GOOGL"]`
  - 新增 `cf_overflow_tables` fixture（module scope，同 `all_tables`）
  - 3 個 `@pytest.mark.cf_overflow` live tests：
    - `test_cf_overflow_rows_exist` — COHR/LITE 至少有 1 個 CF overflow row
    - `test_cf_overflow_multi_quarter_coverage` — 至少 1 個 overflow row 有 ≥2 季數據（驗證 YTD 減法有效）
    - `test_cf_overflow_no_all_none_rows` — 無全 None 的 overflow row
  - 執行：`pytest -m "slow and cf_overflow"` 實測約 5 分鐘

- **`conftest.py`**：新增 `cf_overflow` marker 定義

**Live 驗證結果（2026-04-26）：15/15 PASSED**
- COHR ✅ LITE ✅ AAPL ✅ NVDA ✅ GOOGL ✅（5:05 分鐘）
- 三項驗證全過：overflow rows 存在、≥2 季有數據、無全 None rows

---

**Non-GAAP 指標名稱正規化（NVDA 稀疏表格修復）**

問題：NVDA press release 的表格包含多個比較期間，AI 回傳如：
- `"Non-GAAP Q4 FY26 Gross margin"` / `"Non-GAAP Q3 FY26 Gross margin"` / `"Non-GAAP Q4 FY25 Gross margin"` — 同一指標三個版本
- `"Non-GAAP FY2026 Gross margin"` — 全年版（重複）
- `"Expected Non-GAAP Gross margin (Q1 FY27)"` — 展望指標（不應存）
- 修前 Q4 FY26 filing 回傳 34 個 metrics；修後 6 個

- **`fetcher_nongaap.py`**：
  - 新增 `_clean_metric_name(name)` — 用 `_PERIOD_TOKEN_RE` 移除所有期間 token（`Q4 FY26`、`FY2026`），再移除空括號和噪音標籤
  - 新增 `_normalize_nongaap_metrics(raw)` — 兩段式去重：quarterly token（或無 token）優先，FY-only token 補缺；同類別內第一次出現的值優先（press release 以最新期為首，比較期自動被丟棄）
  - `_call_ai()` 回傳前呼叫 `_normalize_nongaap_metrics(result)`

- **`tests/test_fetcher_nongaap.py`**：新增 12 個正規化 unit tests（37/37 通過）

**Unit tests：110/110 通過**

---

**金融股警告 + Tab 2 Non-GAAP 批量支援**

- **`main.py`**：
  - 新增 `_FINANCIAL_SECTOR_TICKERS = frozenset({"GS", "JPM", "BAC", "C", "WFC", "MS", "BLK", "BX", "KKR"})`
  - `_worker_single` / `_worker_batch`：fetch 完成後若 ticker 在金融股集合內，log 警告「BS/IS 部分欄位可能為空」
  - Tab 2 新增「同時抓取 Non-GAAP」Checkbutton + API Key 警告 label
  - `_on_batch_nongaap_toggle()`：切換 checkbox 時顯示/隱藏 API Key 警告
  - `_run_batch()`：讀取 Non-GAAP checkbox 狀態，未設 API Key 時 error dialog
  - `_worker_batch(fetch_nongaap: bool)`：若啟用，對每個 ticker 呼叫 `fetch_nongaap_statements` 並合併 tables

**Data_EPS_Recon 說明**：edgartools `eps_reconciliation` 對 NVDA/AAPL/MSFT 均回傳 None（非 XBRL tagged），無法從 edgartools 取得。此功能保留為已知限制，待 edgartools 改善或改用 AI 解析方案。
- COHR ✅ LITE ✅ AAPL ✅ NVDA ✅ GOOGL ✅（5:05 分鐘）
- 三項驗證全過：overflow rows 存在、≥2 季有數據、無全 None rows

---

### 2026-04-25（Session 13）

**B1 Overflow Rows — 確保所有 XBRL 財務數字都不遺漏**

設計文件：`docs/superpowers/specs/2026-04-25-overflow-rows-design.md`

核心概念：每次 filing 建表時追蹤模板已消耗的 XBRL row indices（`consumed: set[int]`），未被消耗的 rows 以 overflow 形式追加在模板行之後。

- **`fetcher_gaap.py`**：
  - 新增 `_NONGAAP_KEYWORDS: frozenset` + `_is_nongaap_label(label)` — keyword-based 分類，label 含 "adjusted"/"non-gaap"/"excluding"/"excl."/"ex-" 的視為 Non-GAAP
  - 新增 `_collect_overflow(df, consumed, data_col, quarter_label, gaap_out, ng_out)` — 從未消耗的 XBRL rows 中收集數值，分流至 GAAP / NG 兩個 output dict；跳過 abstract / breakdown / dimension rows（沿用 `_consolidated_mask`）
  - `_build_is_table` → 回傳 `tuple[StatementTable, StatementTable]`：
    - IS df 的 consumed 追蹤；CF-source rows 只追蹤 IS df consumed（CF overflow 由 `_build_cf_table` 負責）
    - `gaap_tbl`（`Data_IS`）= 22 個模板行 + GAAP overflow 行；`ng_tbl`（`Data_IS_NG`）= NG overflow 行
  - `_build_bs_table` → 回傳 `tuple[StatementTable, StatementTable]`（`Data_BS` / `Data_BS_NG`）
  - `_build_cf_table` → 回傳 `tuple[StatementTable, StatementTable]`（`Data_CF` / `Data_CF_NG`）；只對非 YTD filings（Q1/FY）收集 overflow，避免需要跨 filing 減法
  - `fetch_gaap_statements` → 所有 6 次 build 呼叫改用 tuple unpacking；若任一段有 NG overflow rows，則建 `Data_Financials_NG(Q)` / `Data_Financials_NG(Y)` 並加入輸出清單

- **`tests/test_fetcher_gaap.py`**：
  - 新增 16 個 helper tests（`_is_nongaap_label` × 8，`_collect_overflow` × 8）
  - 新增 3 個 smoke tests（`_build_is/bs/cf_table` 空 filings 回傳 tuple）
  - 所有現有 `_build_is_table` / `_build_cf_table` tests 改為 `gaap_tbl, _ = ...` 解包

已知限制（待辦，不在本次範圍）：CF Q2/Q3 YTD overflow 需要跨 filing 減法，目前跳過；地雷十二仍存在。

**69/69 unit tests 全數通過**

---

### 2026-04-24（Session 11）

**Auto-repair Override Engine**

新增 `override_engine.py`，fetch 完成後自動偵測並修復關鍵欄位缺失，無需人工介入。

設計文件：`docs/superpowers/specs/2026-04-23-auto-repair-design.md`

- **`override_engine.py`（新檔）**：
  - `check_key_rows()`：檢查 9 個 key rows（Revenue / Operating Income / Net Income / EPS Diluted / Total Assets / Total Liabilities / Total Equity / OCF / Capex）最近 4 期是否全為 None
  - `e1_fuzzy_match()`：Rule-based，用 `SYNONYM_MAP` 對 EDGAR DataFrame 做子字串比對（免 API 費用）
  - `e2_llm_diagnose()`：E1 失敗時呼叫現有 AI API，提供概念清單給 LLM 識別正確 std_concept 或確認為 structural_absence
  - `load_overrides()` / `save_overrides()`：永久記錄診斷結果至 `%APPDATA%/SEC_Financial_Tools/ticker_overrides.json`，下次同 ticker 不重診斷
  - `run_diagnosis()`：串接 E1 → E2，診斷完自動存檔

- **`fetcher_gaap.py` 整合**：
  - `_apply_row_override()`：新 helper，按 override 的 fix_type（concept_override / structural_absence）從 DataFrame 取值
  - `_build_is_table` / `_build_bs_table` / `_build_cf_table`：各加 `*_overrides` 參數，filing loop 開頭先套用 override，不再呼叫 `_match_is_row`
  - `fetch_gaap_statements()`：加 `ai_config` 參數；三表建完後自動 check_key_rows → run_diagnosis；有新 override 時重建三表（當次 fetch 直接出正確結果）

- **`main.py`**：兩處 `fetch_gaap_statements` 呼叫改傳 `ai_config=self.cfg.get("ai", {})`

**Bug 修復：E2 LLM response 解析**

- 舊邏輯 `response.upper() == "ABSENT"` 只能抓完全等於 "ABSENT" 的回應
- 新邏輯：`re.search(r'\bABSENT\b', ..., re.IGNORECASE)` 處理 ABSENT 嵌在句子中的情況
- 新增垃圾回應防護：含空格或長度 >100 → 回傳 None，不存入 override

新增 **28 個 unit tests**（`tests/test_override_engine.py`）+ 4 個 override 整合測試（`tests/test_fetcher_gaap.py`）；**總計 151 tests，全數通過**

---

### 2026-04-24（Session 12）

**Live Snapshot Tests（自動化實機驗證）**

新增 `tests/test_live_snapshots.py`，對真實 EDGAR API 抓資料，斷言 key rows 不全為 None。

- **`tests/test_live_snapshots.py`（新檔）**：
  - 24 個 `@pytest.mark.slow` tests（8 tickers × IS/BS/CF）
  - Tickers：MSFT、AMZN、META、GOOGL、NVDA、JPM、GS、JNJ
  - `all_tables` fixture（module-scoped）：8 間公司只抓一次（max_filings=8）
  - 金融股（GS/JPM）IS 允許 Operating Income 缺失；JPM CF 允許 Capex 缺失
  - 實跑結果：**24/24 PASSED**（總耗時約 36 分鐘）

- **`conftest.py`（新檔）**：註冊 `slow` marker，避免 PytestUnknownMarkWarning

- **`override_engine.py` bug 修正**：
  - KEY_ROWS 三個命名錯誤（與 StatementTable concept 名稱不符）
  - `"EPS Diluted"` → `"Diluted EPS"`；`"Total Equity"` → `"Total Equity — Parent"`；`"Capital Expenditures"` → `"Capex"`
  - 同步更新 SYNONYM_MAP key 名稱
  - 影響：原本 3/9 key rows 的 override 自動診斷為無效 no-op，現已修復

- **`tests/test_override_engine.py`**：更新 IS_CONCEPTS mock 及 e1_fuzzy_match 測試名稱

- **`PITFALLS.md` 地雷十六**：JPM CF Capex structural absence（銀行類公司使用不同 XBRL 概念名）

執行方式：`pytest -m slow`（~36 分鐘）、`pytest -m "not slow"`（~9 秒）

**總計 175 tests（151 unit + 24 live），全數通過**

---

### 2026-04-23（Session 10）

**CF YTD→季度換算**
- `fetcher_gaap.py`：新增 `_ytd_col()` — 偵測 `(YTD)` 欄位（edgartools 對 Q2/Q3 CF 的標記格式）
- `fetcher_gaap.py`：新增 `_prev_quarter_label()` — 將 "FY2025Q2" 對應到 "FY2025Q1"
- `fetcher_gaap.py`：改寫 `_build_cf_table()` — 不再呼叫 generic `_build_template_table`；改為：
  - Q1/FY：直接讀 standalone 欄位（行為不變）
  - Q2/Q3：讀 YTD 欄位 + 借 IS 取 quarter label（同 BS 做法）；儲存原始 YTD 值，最後做 `Q2 = Q2_YTD − Q1`、`Q3 = Q3_YTD − Q2_YTD` 換算
  - 無前期資料時保留 YTD 原值（best-effort）
- `fetcher_gaap.py`：在 `_match_is_row` 上方加 TODO — 提醒驗證 fallback 涵蓋 ≥10 間公司（MSFT/AMZN/META/GOOGL/NVDA/JPM/GS/JNJ 待測）
- 新增 13 個 unit/integration tests（總計 119 tests，全數通過）

---

### 2026-04-23（Session 10）

**核心修復：`_match_is_row` 串接優先級邏輯（Cascading Priority）**
- 舊邏輯：`label_hint` 找不到符合行時，仍從未過濾的候選行取 first/last（可能拿到錯誤項目）
- 新邏輯：`label_hint` 不符合 → 傳回 None → 呼叫端跳至下一優先級（std_concept → fallback_suffix → label_fallback）
- 避免「hint 失效後偷跑到錯誤行」的問題

**IS 修復**
- `Cost of Revenue`：新增 `label_hint="cost"`，避免 XOM "Sales-based taxes" 誤匹配（其 std_concept 也是 CostOfRevenue）
- `Operating Income`：fallback_suffix 改為 `"OperatingIncomeLoss"`，避免比對到 `us-gaap_OtherOperatingIncomeExpenseNet`（舊 "OperatingIncome" 是後者的子字串）；同時移除已破損的 `label_hint="operation"`（"Operating income" 含 "operat" 不含 "operation"）
- `Gross Profit` 新增 DERIVED fallback：若無直接匹配，從 `Revenue − COGS` 衍生（修復 COHR）

**CF 修復**
- `Net Income`：fallback_suffix 改為 `"NetIncomeLoss|ProfitLoss"`（regex OR），修復 XOM CF Net Income = None（XOM 用 ProfitLoss 而非 NetIncome）
- `Change in Receivables`：新增 `label_hint="receivable"`，搭配 cascading 修復 COHR（ChangeInReceivables std_concept 誤貼在 "Income taxes" 行）
- `Cash Taxes Paid`：label_hint 從 `"income tax"` 改為 `"paid"`，排除 "Deferred income taxes"（不含 "paid"）
- `Cash Interest Paid`：label_hint 從 `"interest paid"` 改為 `"paid"`，修復 COHR "Cash paid for interest"（詞序不符）
- `Operating/Investing/Financing Cash Flow`：label_hint 從 `"net cash"` 改為 `"^net cash|^cash"`（正則 starts-with），修復 AAPL "Cash generated by operating activities" 同時繼續排除 XOM "Noncash right of use assets..."

**實機測試結果（AAPL / TSLA / XOM / NVDA / COHR）**
- AAPL：IS 正確 ✅；CF OCF/ICF/FCF Q1 正確（FY2025Q1 OCF=$53.9B）✅
- TSLA：IS/CF 正常 ✅
- XOM：Revenue 正確；Gross Profit / Operating Income = None（XOM 不報此兩行，預期行為）✅；CF OCF FY2025Q1=$13.0B ✅
- NVDA：IS/CF 正常 ✅
- COHR：Gross Profit DERIVED 正確 ✅；Operating Income = None（COHR 無此獨立行，正確）✅

---

### 2026-04-18（Session 9）

**Bug 修復（實機測試發現）**
- `fetcher_gaap.py`：CF 三大彙總行（Operating/Investing/Financing Cash Flow）新增 `label_hint="net cash"`，避免 `match="last"` 因相同 `standard_concept` 拿到 trailing noncash 項目（如 XOM 的 ROU lease 調整項）
- `fetcher_gaap.py`：FCF 計算改為 `OCF − abs(Capex)`，修正 XOM 等以負數回報 Capex 的公司 FCF 加倍的問題

**實機測試結果**
- TSLA、BA：三表輸出正常，IS/BS 數值與公開財報吻合 ✅
- XOM：OCF 從 $6M（誤）修正為 $12,953M，FCF 從 $5,904M（偶然正確）修正為 $7,055M ✅
- 已知限制：Q2/Q3 CF 全為 None（XOM/TSLA/BA 均以 YTD 格式回報，`_current_q_col` 跳過）

---

### 2026-04-17（Session 8）

**Bug 修復**
- `fetcher_nongaap.py`：`fetch_nongaap_statements` 新增 `max_filings` 參數（預設 80），首次抓取不再無限往回到 2004 年
- `main.py`：呼叫 `fetch_nongaap_statements` 時傳入 `max_filings`，與 GAAP 上限保持一致

**程式碼文件化**
- `main.py`：補上 SECFetcherApp class docstring（說明 thread/queue 架構）及 27 個 method docstring（涵蓋所有原本空白的方法）
- 重點標注非顯而易見的行為：`_wl_draft` 暫存模式、cache-first 查詢邏輯、Tkinter thread 安全機制（msg_queue + _poll_queue）、double-run 防護等
- 其他五個檔案（config.py、excel_formatter.py、excel_writer.py、fetcher_gaap.py、fetcher_nongaap.py）原本已有完整 docstring，無需異動

---

### 2026-04-17（Session 7）

**Bug 修復（程式碼審查後）**
- `config.py`：修正 config 值為非 dict 時 `dict.update()` 會 crash 的問題（如舊 config 格式不相容）
- `main.py`：修正 `_wl_add()` 中重複呼叫 `wl_group_var.get()` 的冗餘程式碼
- `fetcher_nongaap.py`：修正 `_call_ai()` 中 AI 回傳非數字字串時 `float()` 整批失敗的問題；改為逐項 try/except

**UI 改善**
- 移除「確認公司」按鈕，Enter 鍵觸發查詢即可
- 輸出設定改為可收合（▼/▶ toggle），預設展開
- 視窗支援縮放（resizable），Log 區域隨視窗高度延伸
- Advanced Settings 新增「預設模板 / 自訂模板」Radio 選擇

**Excel 自訂模板**
- `excel_writer.py` 新增 `template_path` 參數
- 自訂模板模式：保留所有 cell 格式，只寫入數值；超出模板欄數時複製最後一欄格式
- 新增 `_write_sheet_template()` 及 `_copy_cell_format()` helper
- 兩種模式（預設自動著色 / 自訂模板）可在 Advanced Settings 切換

**Watchlist 群組管理**
- 支援股票分群（群組 CRUD：新增、重命名、刪除）
- 刪除群組時自動把 ticker 移至「未分類」
- 群組依字母排序，「未分類」固定最後
- Watchlist popup 改為「儲存關閉 / 放棄關閉」（Escape = 放棄），操作前先建立 deep copy draft

---

### 2026-04-17（Session 6）

**BS 抓取修復**
- `_build_bs_table` 改為獨立實作；BS 欄位為 instant（bare date），`_current_q_col` 無法識別導致全空白
- 修法：從同一 filing 的 IS 借用 quarter label，BS 取第一個非 meta 欄位讀值

**Watchlist popup 改善**
- 「目前 Watchlist」區加入捲軸（固定高度 160px）
- Ticker 輸入框：Enter → 查詢，查到後自動加入（單次 Enter 完成）

**Tab 2 批量更新改善**
- Watchlist 改為捲軸顯示（固定高度 150px），只顯示 ticker 代號，3 個一列

---

### 2026-04-17（Session 5）

**Config 搬家 + Watchlist 路徑管理**
- config.json 移到 `%APPDATA%\SEC Financial Tools\config.json`，啟動時自動 migrate 舊檔
- Watchlist 管理介面每行新增 📁 按鈕，可為每間公司設定獨立輸出資料夾
- 路徑存於 watchlist item `output_dir` 欄位，優先順序：watchlist `output_dir` → `ticker_paths`（向後相容）→ 全域 `output_dir`

---

### 2026-04-17（Session 4）

**Excel 自動美化（Phase 3）**
- 新增 `excel_formatter.py`：format_workbook() 在存檔前自動套用所有格式
- 欄寬修正（A=22, B=24, 資料欄=13）：解決科學記號顯示問題
- 深藍色 header 列（Row 1/2）、藍色 section header、灰色分隔列、交替底色、subtotal 粗體
- 財務數字自動 ÷1M，套用 `#,##0.0` 千分位格式；EPS 保留原值用 2 位小數；Shares ÷1M 整數
- Index sheet 自動插入第一頁：列出所有 Data_* sheet 用途、最早/最新期間
- 凍結窗格 C3（Rows 1–2 + Cols A–B 固定）
- 新增 Data_Financials(Y)（年報 10-K），原 Data_Financials 更名為 Data_Financials(Q）
- 新增 72 個 unit tests，全數通過（總計 106 tests）

---

### 2026-04-17（Session 3）

**Per-Ticker Output Path Memory**
- config.json 新增 ticker_paths 欄位
- 確認公司後自動帶出已記憶路徑
- 瀏覽選資料夾後自動儲存至 ticker_paths

**Non-GAAP Fetching（Phase 2）**
- fetcher_nongaap.py 完整實作
- 8-K Item 2.02 篩選，EPS reconciliation（edgartools 原生）
- AI 從 EX-99.1 press release 提取 Non-GAAP 指標（Google / OpenAI / Anthropic）
- nongaap_cache.json 增量快取，只對新季度呼叫 AI
- 輸出：Data_EPS_Recon + Data_NonGAAP sheet

---

### 2026-04-15（Session 2）

**萬能模板實作**
- IS_TEMPLATE 從 18 行擴展至 22 行（新增 D&A、SBC、Minority Interest、Total Non-op）
- 新增 BS_TEMPLATE（41 行）、CF_TEMPLATE（25 行 + FCF 衍生）
- 模板 tuple 從 4-tuple 升級為 6-tuple，加入 `match` 和 `label_hint` 欄位
- `_match_is_row` 新增第三層 label fallback（解決 TSLA D&A nan 問題）
- 三表合一：IS + BS + CF 合併輸出為單一 `Data_Financials` sheet，section header 分隔
- `StatementTable` 新增 `labels: list[str]` 欄位（B 欄 Original Item）
- `excel_writer.py` 改為 A=Std Name / B=Original Item / C+=季度數據
- Post-processing fallbacks：ProfitLoss（Net Income）、DERIVED（Total Non-op）、label "depreciation"（D&A）
- GOOGL encoding fix：`unicodedata.normalize("NFKC")` 處理非 ASCII 標籤
- 新增 53 個 unit tests，全數通過

**GUI 設定**
- Advanced Settings 加入 max_filings 調整（Spinbox，from=4 to=320，預設 80）

### 2026-04-13（Session 1）

- 完成 Phase 1：GAAP fetcher + Excel writer + 完整 Tkinter GUI
- 初始化專案
