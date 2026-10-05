# 分類生命週期

每套分類 registry 記錄 dimension、version_id、effective_from、effective_to、status、nodes（node_id、original_label、parent_id）、來源、變更原因與可比性。版本生效日期以公司披露為準；抽取到的最早值不一定等於分類生效日。

狀態及數據分開：registry 的 active/inactive 是目前顯示規則；record 的 verified/provisional/not_disclosed 是抽取核對狀態；presentation 的 as_reported/recast 是披露口徑，三者不能混用。

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
