# Revenue 總計選值驗證（2026-10-06）

## 改動與邊界

本輪修正模板選值及 Excel 分頁命名，不重寫正式財報資料庫。保存的 DataFrame 原始概念與金額仍在；先前標準化將總計和構成項都映射成 Revenue，模板又選第一列，才使管理費／產品收入冒充營收。詳見 [架構邊界](../ARCHITECTURE.md) 及 [原調查](revenue-selection-investigation-2026-10-06.md)。

新規則先選無維度原始總計，不取最大金額；calculation parent 能證明子項時排除子項。同層總計衝突留空、報 `AmbiguousRevenueTotal`，所有未使用來源仍在 overflow。已辨識總計缺值不借構成項，Revenue override／diagnosis rebuild 不能繞過總計。沒有已辨識總計的自訂概念維持舊規則，並未在本輪認證。

驗收另外抓到 `_merge_financials` 以顯示名稱判斷固定列，讓同名 Revenue／Cash 的來源構成列插入模板，推移 BS／CF。已改成依來源表的模板 slots 身份分類，NG 使用零模板 slots。舊基線有 71 張財務表列位置不符，中間 Revenue 候選仍有 68 張；數值按名字／日期比對不足以找出這種風險，因此新增完整固定列順序檢查，並重跑全庫，沒有直接沿用中間版驗收。

## 原 SEC 來源驗證

七份原 accession XBRL instance 獨立下載在 `output/`，fixture 保存 URL、SHA-256、原始概念、完整 start/end、context 與 USD 值。`verify_original_revenue_instances.py` 不使用模板／edgartools 標準化，也不存取資料庫；解析原 XML 的 unit namespace、無維度 context 與 exact duration，再由 `verify_revenue_source_examples.py` 比較實際輸出。

| 公司 | 期間 | 原始概念 | Revenue（USD） |
|---|---|---|---:|
| MAR | 2009-01-03～2010-01-01 年度 | Revenues | 10,908,000,000 |
| HLT | 2017-01-01～2017-03-31 單季 | Revenues | 2,161,000,000 |
| AMZN | 2013 年度 | SalesRevenueNet | 74,452,000,000 |
| MSFT | 2015-07-01～2016-06-30 年度 | SalesRevenueNet | 85,320,000,000 |
| AFL | 2026-04-01～2026-06-30 單季 | Revenues | 4,117,000,000 |
| C | 2026-04-01～2026-06-30 單季 | Revenues | 24,766,000,000 |
| COST | 2018-02-19～2018-05-13 單季 | Revenues | 32,361,000,000 |

MAR 原本選管理費 530M；HLT 原本選 83M。COST 的 SalesRevenueNet 31.624B 是 calculation 子项，總計為 32.361B；不是看到較大金額就採用。本輪銀行／保險來源抽樣為 C、AFL，不能推廣成整個產業所有數字已認證。

## Excel 與回歸

- 非 slow 測試：1,749 passed、65 deselected；五個 warning 為既有相依套件棄用及 docstring 跳脫字元警告，沒有測試失敗。
- 七家完整工作簿含比率，以專屬隱藏 Excel instance 實際開啟、重算、儲存，另修改 Index 財年月份後讀取表頭，最後不保存月份覆寫。
- 211,776 格已核對、0 discrepancy，包含 27,199 格比率與 7,180 格表頭檢查。這證明匯出與模板／比率計算輸入一致，不替全部比率的經濟定義背書。
- MSFT `2011-12`、AFL `2010-06` 兩個只有年月的舊期間，各三格月份覆寫表頭共六格列為 INCONCLUSIVE，**不包含在上述通過格數**。Excel 會產生標籤，但缺完整期末日不可宣稱來源期間已認證；原保存值未改。
- 實際開檔找到 MAR 兩個 31 字分頁只差大小寫，openpyxl 自动加號成 32 字。writer 改為事前分配大小寫唯一、預留後綴長度的名稱；來源 `StatementTable` 不變，Index 使用合法輸出名。一般與 template 分支回歸均通過。

## 全庫與資料庫驗收

- 最新固定程式 `67e0a46` 從正式庫重新執行全部 215 家，正式 listing 完整、215 對 215，缺公司／失敗 sidecar／非預期本地解析例外皆 0；487 張財務表的固定列順序全部通過。沿用金融行為等同 master `77e7805` 的上一輪正式全庫基線，兩側使用相同凍結官方 listing、正式讀取條件及 80 季報／20 年報限制。
- 新比較器 `compare_revenue_audits.py` 以 section＋固定 slot 或原來源 key 識別列，保留同日期全部非空值及重複次數。104 家有數值差異；固定模板差異只有 Revenue 4,341 筆及連動 Gross Profit 2,052 筆，BS／CF 及其他固定指標沒有新數值差異。這些是列／期末日的比較記錄，不是逐筆原 SEC 真值認證。
- 主 Revenue 的 4,341 筆：4,224 筆由原選值改為新值、77 筆補出總計、40 筆保守留空（含拆季連動）。剩餘 `AmbiguousRevenueTotal` 29 筆 ledger 記錄：COP 24、OXY 4、TGT 1；既有 `MissingStandalonePeriod` 230、`MissingCurrentPeriod` 86。不能把全部 data 缺口當成此次新增錯誤。
- Q／Y 金融表的期間標籤與期間表頭差異 0；NG 頁因來源構成列移入／移出而有 86 筆標籤、258 筆表頭及 13 筆值差異，保留來源，不代表 NG 分類已認證。
- 初次只用全表同名 occurrence 比對，將 IS 的原始 Net Income 構成列與 CF 固定 Net Income 錯配，產生 541 筆假回歸；已先用三個 RED 重現、修比較器再重比全部 215 家，這個舊統計撤回。金融 pipeline 不需為量測修正重跑。reviewer 額外驗證重複 source key 及同日期新增非空值未被 overwrite，但不任意配對 duplicate filings。
- 七本 Excel 的原始輸入表與最新全庫七家公司表完全相同；原 Revenue 七例及前輪 CDNS／JNJ 的 15 格來源衍生數字全部通過。
- 正式庫全部非 `.git` 檔案 29,277 個、3,204,105,955 bytes，前後逐檔 SHA-256 完全相同：新增／移除／變更皆 0；資料庫 Git 工作區乾淨。模板修正未改 filing JSON、metadata 或 history。

精簡機器證據：[revenue-verification-summary.json](evidence/revenue-verification-summary.json)。完整輸出為 `output/revenue-accepted-all`、`revenue-stable-comparison`、`revenue-workbooks-accepted`；中間候選與誤判比較另存，未覆寫或用來代替最終驗收。

## 決策與剩餘工作

1. 不補寫或覆蓋原來源金額。一般更新仍可能下載／保存新解析結果，與本輪模板唯讀驗收不同。
2. 同層總計衝突、包含 Other Income 的口徑、沒有認得的自訂總計，接續逐 accession/context 查原申報；不能以最大值或多數決解決。
3. KR 缺當期 facts、KHC predecessor/successor 併購口徑、無可信年報錨點、transition 身份與 GUI 月份預覽仍是獨立待辦。G13 未全解。
4. YYYY-MM 期間的 Excel 月份覆寫需另立規格與來源恢復工作，不能將不完整日期硬補成真實期末日。
5. reviewer 提出的 synthetic calculation cycle＋duplicate metadata edge 為 Minor，庫內未發現實例；列為後續防禦測試，不冒稱已觀察到來源錯誤。
6. NG overflow 的既有 `excluding` 字串分類也會匹配稅額排除：UNP、MPC 有來源 GAAP 構成列移到 NG 頁的實例。原值保留，但分頁名稱不證明 Non-GAAP 口徑；本輪不改此分類器，另列 TODO，不將 NG 新增表頭當成財季映射改動或來源丟失。

## Checkpoints 與重現

`b76b97f` RED；`867d1a0` 原始總計優先；`bb28678` override／rebuild 防護；`d682ac6` parent／缺值；`8645585` 七份來源例；`c741f9a` Excel RED；`f4afd8b` 分頁修正；`d0e4d72` Excel 驗證工具；`42c67bf` 原 XML 獨立驗證；`575662c` 固定列 RED；`67e0a46` 來源模板身份；`416ea15`／`f38a591` 全庫純讀固定列驗證；`d8a9adf` 比較器 RED；`8eaf7e8` section／source 身份比較。

最終全庫金融程式固定 `67e0a46`，輸出 `revenue-accepted-all`；後續驗證腳本與文件不改金融計算。早期 `revenue-candidate-all`／`revenue-final-all` 都是被取代的中間版本，不能當成最終驗收。

```powershell
# 在專案根目錄執行，使用專案外 venv。
$revenuePython = Join-Path $env:USERPROFILE 'venvs/SEC Financial Tools/Scripts/python.exe'
$env:PYTHONIOENCODING = 'utf-8'
$env:PYTHONDONTWRITEBYTECODE = '1'
$env:EDGAR_LOCAL_DATA_DIR = Join-Path (Get-Location) 'output/database-verification-edgar'
& $revenuePython scripts/audit_fiscal_pipeline.py --all --require-official-metadata --output output/revenue-accepted-all
& $revenuePython scripts/compare_revenue_audits.py output/fiscal-official-final-all output/revenue-accepted-all --output output/revenue-stable-comparison
& $revenuePython scripts/verify_original_revenue_instances.py output
& $revenuePython scripts/verify_revenue_source_examples.py output/revenue-accepted-all
& $revenuePython scripts/verify_fiscal_source_examples.py output/revenue-accepted-all
& $revenuePython scripts/verify_fixed_financial_rows.py output/revenue-accepted-all
& $revenuePython scripts/generate_revenue_verification_workbooks.py --code . --source output/revenue-accepted-all --output output/new-owned-revenue-workbooks
./scripts/recalculate_verification_workbooks.ps1 -WorkbookFolder output/new-owned-revenue-workbooks
& $revenuePython scripts/verify_revenue_workbooks.py output/new-owned-revenue-workbooks
& $revenuePython -m pytest -m "not slow" -q -p no:cacheprovider
```

全庫 audit 需要相同凍結的官方 listing。腳本每十家公司重開 process，禁用保存 hook、AI diagnosis、companyfacts 股數與 gap 網路探測；結果不代表最新 SEC 申報完整性。Excel 重算只供自己新建的驗證資料夾使用。大型原 XML／稽核／工作簿在 gitignored `output/`；可攜證據與驗證規則保存在此文件與 fixture。

本機 edgartools 有舊 runtime cache 的 locale／ACL 清理警告，部分 PowerShell 包裝會因 stderr 回傳 1；未依提示刪除快取，也未刪正式庫。驗收需檢查全部公司 JSON、metadata mode、失敗 sidecar、ledger exception 與比較器結果，不能只用包裝程式 exit code 判斷 215 家是否成功。
