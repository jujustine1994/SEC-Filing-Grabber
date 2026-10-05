# 跨公司 Excel 固定格式（版本 1）

使用者指定的共用輸出規範。所有公司的營收拆分 Excel 固定只有下列四張工作表，名稱（含空格及大小寫）與順序完全一致；不得新增 Recent、公司名稱工作表或改成翻譯名稱。公司名稱、ticker、分類與期間放在表內。若使用者另行明確指定版面，以其要求為準。

| 順序 | Sheet | 固定用途 |
| --- | --- | --- |
| 1 | Quarterly History | 完整單季營收拆分主表，最新分類在上、Inactive Segments 在下；原始營收、披露的部門 OP、外部／部門間營收。僅保留簡單的 OP%。 |
| 2 | Annual History | 官方年度拆分；最新／停用分類分組。過渡期間獨立標示，不能冒充全年或將累計值當單季。 |
| 3 | Coverage | 公司及資料期間、分類／重編規則、數值核對與具體缺口；逐季列出各拆分維度的取得狀態。 |
| 4 | Sources | 原始官方來源清單及逐筆揭露紀錄，保留原標籤、期間、分類版本、原報／重編、數值／單位、計量口徑、來源定位與原文註解；配套 JSON 可追溯顯示儲存格。 |

表內的產品、市場、報告部門、地域與其他維度依公司實際披露動態建立，不套用 NVDA 的分類或財年規則。沒有披露的維度／OP 留空並在 Coverage 說明；即使完全沒有年度拆分，Annual History 仍保留，明確標示缺口。Inactive Segments 是表內區塊，不是第五张工作表。

季度與年度主表固定 B2 公司／主題標題、B3 單位與口徑說明、B5 分類欄名、C5 起按時間遞增的期間；B6 起為內容。凍結前五列及前兩欄。分類版本與父子層級須可辨識，不同維度不可相加。預設不加占比、QoQ、YoY、預測及比例回推；OP% 僅為同口徑 OP / Revenue，公式必須保留，缺值或零分母留空。官方直接披露的成長率／毛利率可當原始指標另列，清楚標示 reported，不當成自行計算。

原始資訊優先：保留 Revenue、官方 consolidated total、reported segment OP，及公司披露的 external revenue、intersegment revenue、eliminations / unallocated items。公司另有毛利、銷售量、ASP、客戶／產品細分類時，只有確實相關且直接披露才增加指標列，標示原單位；不自行計算 ASP、EBITDA 等。地域需保留按客戶所在地、帳單地址或最終需求等定義。區分 GAAP、non-GAAP 與 segment measure，不能跨口徑代用。

Sources 上方 A1:F1 固定來源清單標頭：Source ID、Form、Document type、Filed、SEC official URL、SHA-256。清單結束後留三行，再放 Raw disclosures 表；空白模板此表在第 6 列。每筆揭露保留 record ID、period start/end、quarter/annual/YTD/transition、metric、原標籤／parent／dimension／classification version、原報／比較／官方重編身分、原始值及單位、標準化值及單位、currency、basis、filed/accession、source ID/locator/URL、direct/derived、推導輸入 IDs、verification、原文註解與顯示格對照。尚未保存的原始字串／原文註解留空，Coverage 標示未擷取，不把標準化數字冒稱原字串；原 SEC 文件保持可用。原始揭露可包含累計值，主季度表不可把 YTD 混成單季。

先複製本 skill 的 `assets/Revenue_Breakdown_Template.xlsx` 再填資料，替換尖括號佔位文字，按公司披露增加分類／期間；不得把佔位文字交付為研究結果。Template 是空白骨架，不含虛構公司數值。需要重建時，依試算表 skill 執行 `scripts/build_workbook_template.mjs <輸出資料夾>`。

利用後續申報中的官方前期重編值延長最新分類歷史，保留原報版本及公布來源；只有可以唯一計算且口徑一致時才推導，不能按新分類比例回推舊季。

產出後執行 `scripts/validate_workbook_format.py <xlsx>`，以標準函式庫檢查四個 sheet 名稱與順序；另核對數值、公式、來源和版面。此檢查不表示內容已完整或分類已正確。
