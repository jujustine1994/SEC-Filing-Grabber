# 跨公司 Excel 固定格式（版本 1）

使用者指定的共用輸出規範。所有公司的營收拆分 Excel 固定只有下列四張工作表，名稱（含空格及大小寫）與順序完全一致；不得新增 Recent、公司名稱工作表或改成翻譯名稱。公司名稱、ticker、分類與期間放在表內。若使用者另行明確指定版面，以其要求為準。

| 順序 | Sheet | 固定用途 |
| --- | --- | --- |
| 1 | Quarterly History | 完整單季營收拆分主表，最新分類在上、Inactive Segments 在下；原始營收、披露的部門 OP、外部／部門間營收。僅保留簡單的 OP%。 |
| 2 | Annual History | 官方年度拆分；最新／停用分類分組。過渡期間獨立標示，不能冒充全年或將累計值當單季。 |
| 3 | Coverage | 公司及資料期間、分類／重編規則、數值核對與具體缺口；逐季列出各拆分維度的取得狀態。 |
| 4 | Sources | 原始官方來源清單及逐筆揭露紀錄，保留原標籤、期間、分類版本、原報／重編、數值／單位、計量口徑、來源定位與原文註解；配套 JSON 可追溯顯示儲存格。 |

表內的產品、市場、報告部門、地域與其他維度依公司實際披露動態建立，不套用 NVDA 的分類或財年規則。沒有披露的維度／OP 留空並在 Coverage 說明；即使完全沒有年度拆分，Annual History 仍保留，明確標示缺口。Inactive Segments 是表內區塊，不是第五张工作表。

季度與年度主表固定 B2 公司／主題標題、B3 單位與口徑說明、B5 分類欄名、C5 起按時間遞增的期間；B6 公司總營收，B8 起為分類區塊。凍結前五列及前兩欄。分類版本與父子層級須可辨識，不同維度不可相加。營收、占公司總營收、季度 QoQ、季度／年度 YoY 為標準列；預測與比例回推不放 B。營收展示以 record ID 的 INDEX／MATCH 讀 Sources，不以固定列號識別來源；比重、成長、利益率與有證據的 Q4 推導由程式產生 Excel 公式，不把計算結果寫死。Sources 原始值仍是值。官方直接披露的成長率／毛利率需明標 reported，不冒充自行計算。

原始資訊優先：保留 Revenue、官方 consolidated total、reported segment OP，及公司披露的 external revenue、intersegment revenue、eliminations / unallocated items。公司另有毛利、銷售量、ASP、客戶／產品細分類時，只有確實相關且直接披露才增加指標列，標示原單位；不自行計算 ASP、EBITDA 等。地域需保留按客戶所在地、帳單地址或最終需求等定義。區分 GAAP、non-GAAP 與 segment measure，不能跨口徑代用。

Sources 上方 A1:F1 固定來源清單標頭：Source ID、Form、Document type、Filed、SEC official URL、SHA-256。清單結束後留三行，再放 Raw disclosures 表；空白模板此表在第 6 列。每筆揭露保留 record ID、period start/end、quarter/annual/YTD/transition、metric、原標籤／parent／dimension／classification version、原報／比較／官方重編身分、原始值及單位、標準化值及單位、currency、basis、filed/accession、source ID/locator/URL、direct/derived、推導輸入 IDs、verification、原文註解與顯示格對照。尚未保存的原始字串／原文註解留空，Coverage 標示未擷取，不把標準化數字冒稱原字串；原 SEC 文件保持可用。原始揭露可包含累計值，主季度表不可把 YTD 混成單季。

新版入口為 `assets/templates/catalog.json` 的三份 formula-v1 版型，由 `scripts/build_b_workbook.py` 選版、填來源與公式。尖括號佔位文字只用於空白模板，不得交付為公司研究結果。舊 `assets/Revenue_Breakdown_Template.xlsx` 與 `build_workbook_template.mjs` 保留作歷史版；不要拿它覆蓋新模板。

## 多份 Excel 版型與共同介面

允許依官方拆分結構維護多份 Excel 版型，不要求所有公司使用相同內容排列。適合的版型方向包括平面拆分、父子階層、報告部門含 OP／調節；同一公司有多個維度時可在四張表內增加獨立區塊，不把不同維度拆成可相加的一棵樹。選版依實際已確認的披露結構，不因公司產業或 ticker 猜類型。

各版型仍共用本文件的四張表、B2／B3／B5／C5 起期間位置、Sources 契約與分類版本／來源對照；差異在內容區塊、父子展示及可用指標，不改成另一套匯入介面。公司特例確實需要新布局時，先核對 C 與 builder 是否支援，不任意新增表或移動固定欄位。

版型可重用於披露結構相似的公司，不預設每家公司新增一份模板。使用時記錄採用的模板檔名與版本；不得把模板示範分類當成公司已披露資料。無適合版型时依原始披露調整並記錄差異，分類不確定仍按 classification-lifecycle.md 處理，不能由模板替它決定。

三份空白模板由同一版面引擎生成，catalog 保存檔名、ID、版本與 SHA-256；已修改模板須先驗證並更新 catalog，不能讓 builder 靜默使用未知格式。新版共用 builder 明確選模板，資料與分類來自 breakdown／views，查核仍來自既有 checks 及原文。NVDA 專用 builder 含公司特定分類、加總及 Coverage 文字，保留歷史用途，不作通用入口。

利用後續申報中的官方前期重編值延長最新分類歷史，保留原報版本及公布來源；只有可以唯一計算且口徑一致時才推導，不能按新分類比例回推舊季。

產出後執行 `scripts/validate_workbook_format.py <xlsx> --plan <workbook-plan.json>`，比對計畫內全部儲存格、公式／來源值、固定位置、可見性、字體與百分比格式，拒絕未規劃內容。沒有 --plan 時保留歷史四頁名稱／順序檢查；它不能當成新版完整驗收。另用原生 Excel 計算、存檔再查分類精確文字及公式結果。核對通過不表示資料完整或所有分類涵義已審計。

## 公式與可比性

- 占公司總營收＝分類營收／同期間同單位公司總營收；子分類如另加占父項比重，須獨立命名，不替換標準公司比重。
- QoQ 比對明確上一財季，YoY 比對上一財年同季／上一完整年度，不以相鄰 Excel 欄位替代缺期。分類版本、節點、單位、幣別、口徑需相同；原報／重編混合時需同份官方申報的可比支持。過渡期不計成長。
- 前期／分母≤0或缺值時留空；官方零營收保留零，其占比可為0，但下一期不能以零為分母算成長。負利益可保留負利益率，只要相同口徑營收為正。
- 營業利益／淨利只有官方披露才列；利益率用同節點、期間、分類版本及口徑的利益／營收，不用部門 OP 推估淨利率。
- Q4 僅用同財年／版本／口徑的全年減九個月累計，來源 IDs 必須齊全，已有期間日期需相容；缺少歷史起日會保留限制，不推定日期已核對。
- 正黑體最小12、原生格線可見、不直接隱藏列欄頁；未指定區域無填滿／框線。分類子項与指標以明確縮排區別，XML 保留前導空白，不能讓 Excel 開存後修剪分類身分。
