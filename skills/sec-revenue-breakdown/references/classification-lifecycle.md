# 分類生命週期

每套分類 registry 記錄 dimension、version_id、effective_from、effective_to、status、nodes（node_id、original_label、parent_id）、來源、變更原因與可比性。版本生效日期以公司披露為準；抽取到的最早值不一定等於分類生效日。

狀態及數據分開：registry 的 active/inactive 是目前顯示規則；record 的 verified/provisional/not_disclosed 是抽取核對狀態；presentation 的 as_reported/recast 是披露口徑，三者不能混用。

## 跨公司分類的建立順序

1. 先盤點公司原文的拆分維度與定義，分開建立 product、market_platform、reportable_segment、geography、revenue_type 或其他有來源支持的維度。不要求每家公司都有這些維度，不套用 NVDA 節點。公司只披露一層時保留平面分類，不為了套版補父子層級。
2. 每個分類以 company + dimension + classification_version + node_id 識別；原文名稱與顯示名稱另存。ID 在同一版本內唯一且穩定，不用 Excel 列號、列順序或中文翻譯當識別，也不按同名去重。既有案例 ID 可沿用，不為此重編舊 records。
3. 節點定義需保留官方來源定位；父子關係也需來源支持。縮排、字體、名稱相似或加總剛好相等只能作查證線索。沒有父項依據就保留獨立列、parent_id=null，並在 checks.unresolved／Coverage 記錄「階層未確認」，不能把它當已驗證的平面互斥拆分。父項須屬同一維度／版本，不能跨維度掛接或形成循環。
4. 同維度可存在多套官方拆分，例如地域按客戶所在地或帳單地址；定義不同就分開版本／區塊。不合併成一套可相加的清單。Original label 保持原文；翻譯或縮寫只改 display_label，不改來源身分。
5. 每個版本記錄來源支持的定義、變更原因與可比性。未知原因明寫「未確認」；effective_from／effective_to 未獲來源支持時留 null，first_observed／last_observed 只代表已取得值的範圍，不冒充正式生效／停用日。既有 registry 若只有 observed 欄位，更新時補說明，不猜日期。

## 遇到變更或缺漏時的判定

| 情況 | 處理 |
| --- | --- |
| 名稱改變，官方支持範圍不變 | 保留原文名稱歷史與證據，可沿用節點；不能只靠翻譯、名稱近似判定 |
| 名稱相同，範圍、計量定義或分類歸屬改變 | 另存分類版本；舊值不接到新定義，新版歷史只用官方同口徑重編值 |
| 分類拆分、合併或組織重整 | 新版本、新節點／關係，保留舊版與官方對照證據；沒有官方數值對照就不分攤或回推 |
| 單季沒找到、來源沒下載、表格未解析或只有年度披露 | 是取得／披露頻率問題，不直接改為 inactive，也不能僅據此判定「公司未披露」；已核對相關官方文件的未披露情況依資料契約另記 |
| 最新文件只是列出舊期比較分類 | 不因此認定該分類仍 active；需核對當期採用的披露方式 |
| 官方說明已改用／停止舊分類，或完整當期表及定義能確認替代關係 | 舊版放 Inactive Segments，保留判定來源；只代表披露分類停用 |
| 是否停用仍不明 | 不強行改 active/inactive；保留既有顯示狀態並標「狀態待確認」，新發現且無已確認狀態的分類作 provisional 獨立區塊，Coverage 說明，不冒稱最新或已停用 |
| 官方披露 Other／Unallocated／Eliminations | 按原文保存及其計量口徑；不能將其與 AI 計算的差額混為同一分類 |

「狀態待確認」是未解問題說明，不要求舊 registry 新增一個 status 枚舉。新抽取的待確認節點／版本可保存在 registry，status 留 null，checks.unresolved 記原因；目前 NVDA 專用 views builder 未支援此一般情境，不能直接重用來決定其他公司的狀態。

## 加總與顯示前核對

- 先確認同期間、同版本、同單位／幣別、同計量口徑，以及分類是否互斥且涵蓋完整。只有確定是完整互斥拆分時，才以葉節點核對該官方總額；父項與其子項不重複計入。
- 公司只披露部分產品、主要國家或大客戶時，不硬要求加總到公司營收。不同維度不可相加；交叉或不完整的分類保留原值與定義，不計算虛構 Other。
- 無適用加總關係時寫 not_applicable；缺來源或互斥／完整性尚未確認時，核對標 provisional 並寫入 unresolved，不以 not_applicable 掩蓋未查證。加總相等不能單獨證明階層、名稱或定義正確。
- 同一分類／期間有不同揭露值時，保留各自 record ID、申報與 as_reported／recast 身分。官方修訂或重編按其證據選顯示值；其他未解衝突不按「最大值」或「最新下載」挑選，主表不放未解的唯一數字，Coverage 列出候選與缺口。
- 輸出前核對 ID 唯一、父節點存在且無循環、每個顯示值對應 record ID、版本及來源；Excel 列名和縮排只供閱讀，不能取代 registry。更新分類可移動顯示列，但不改 record 身分、不刪舊版本。

本節是分類與抽取規範，尚未由跨公司通用分類驗證器強制執行；現有四頁格式檢查不證明上述核對已通過。新公司的 C 分析映射仍需按它自己的維度／節點設定，不沿用 NVDA 的列號與分類名稱。

## 表格規則

1. 最新財務季度採用的分類放上方，維持原文支持的父子階層。父項不和子項一起加總。
2. 已停止披露的分類置於下方 Inactive Segments，按最新停用版本至最早版本排列。保留版本、生效/停用期間與原因；名稱重複也不能消掉不同定義的版本。
3. 原始歷史值選最早發布的有效揭露；已正式修訂的財報另標版本與理由。官方重編前期值保存為額外 records，最新分類顯示使用其最新可比值，不改寫原 records。
4. 同名持續分類（例如 NVDA 的 Data Center 父項）僅有來源支持範圍不變時可跨版本連續顯示；子項換維度不代表父項也必須停用。禁止以名稱相似推斷可比。
5. 無官方重編的舊期間在新列留空；停用列的新期間留空。缺資料、未披露、分類尚未存在以說明區分，不推定 0。
6. 各互斥分類版本獨立核對總額。不同市場維度、報告部門、地域及 active/inactive 不互相相加。
7. QoQ/YoY 只用同版本的比較數。Q4 只從同年度、同分類/單位的全年和九個月累計推導，缺少任一輸入時留空。

## NVDA 回歸範例

最新市場：Data Center 下有 Hyperscale 與 AI Clouds, Industrial, & Enterprise；另有 Edge Computing。旧 Gaming、Professional Visualization、Automotive、OEM and Other，以及 Data Center 的 Compute/Networking 拆分放 Inactive Segments。

FY2027Q2 Hyperscale 48,710 ÷ 該文件重編 Q1 43,050 − 1，不能除以 Q1 原报 37,869。ACIE 同理使用重編 Q1 32,196。兩個版本的 Q1 Data Center 加總均為 75,246。不能將 Q1 版的 FY2026Q1 或 Q4 比較值冒充 Q2 版重編值。

報告部門 Compute & Networking / Graphics 是另一套 active 分類，不因市場改為 Hyperscale/ACIE 而放入 inactive，也不能把其 OP 套用到市場子項。

全歷史模式需列出實際下載數、抽取期間、數值核對、尚未解析的資料。總額核對通過不證明列名、期間與分類涵義全部正確；保留表格定位與原文以便抽查。

歷史抽取教訓：金額括號可能拆成不同 HTML cells，也可能含空格，例如 `( 1,714 )`；必須保留負號，不可因 parser 讀不到就丟掉整列 OP。早期財年結束日可能不同，或有一個月過渡期；以當時原文為準，不用最新財年規則反推全部年份。部門表格中的 Total operating income 可能是部門總額，不得自動標為合併 GAAP OP。
