# 三表期間與現金流驗證紀錄（2026-10-06）

本工作使用外部 venv 與 UUID 指向的外部財報資料庫。所有全庫量測只讀原始 JSON，不下載更新、不寫入資料庫、不執行 AI 診斷；輸出另存專案 `output/`。結果代表既有資料，不代表最新 SEC 申報完整性。

## 已保存的步驟

1. `edb6e5d`：新增裸日期、標籤碰撞、重複期末日與季度順序的量測；使用實際 fetch pipeline，80 季報／20 年報，十家公司換一個 process。
2. `3273eec`：加入文件當期 DEI 與完整連續年報／季報日期的期間識別，以及 Excel 預設財季標籤保留。此為中間 checkpoint，尚待全庫比較及來源矛盾處理，不能視為驗收通過。
3. CF 安全步驟：缺少拆季基準時，累計流量不再當成單季；保留期末餘額、原始累計診斷與 accession，透過既有缺口帳本報告 `data`。原始資料庫不變。原行為測試先失敗、修改後通過；額外驗證餘額、FCF 與缺口分類。

## 決策

- Ruling：優先完成量測、期間與 CF 正確性，再處理警告 UX 和 G13(a) 原始 facts 恢復。舊洞數 4,176 是分類器上限，不是已確認真缺資料。
- Ruling：不把 companyfacts 後續重報值當作原 accession 的真值。基準與候選都排除 companyfacts 股數及 AI 診斷，正式驗收須另註明這個範圍限制。
- Ruling：當期 cover focus 不套到比較期；缺一個季度就不以排序補季度；年度錨點衝突時不猜測。季度 DEI 與完整、連續、命名一致的年報及三季日期鏈矛盾時，以該日期鏈建立季度識別。AEP、ACN 的中間版本回歸促成此防護，不能把 cover DEI 無條件當真。
- Ruling：SEC DEI 也可能錯。KR 年報 `0001558370-24-004603` 和 `0001558370-25-004267` 的 DEI 年份分別為 2024、2025，但文件正文明確定義 2024-02-03 為 fiscal 2023、2025-02-01 為 fiscal 2024。需保存來源證據並精確限制任何修正，不能以多數決覆寫整家公司。

全庫結果、完整回歸與最終審查尚在執行，完成後追加數字及結論。

## 審查修正與範圍裁決

Fresh-context reviewer 審查 `cc22a43..674b5a9`，指出 COHR／CRM 年度映射碰撞及 ARM 6-K 跨年相減的 P1，audit 未套 CIK／核心表閘口及 comparer 忽略失敗／空輸入的 P2。這些已在 `d608c19`／`2c17a76` 修正；全部先重現 RED 再驗證 GREEN。實際 COHR／CRM 重跑後財務／標籤均 0 差異；ARM 標籤 0 差異，僅缺基準的舊 CF 改空白。早期兩個候選均停止驗收用途，另以固定 `2c17a76` checkout 全庫重新量測。

Ruling：年度 DEI 鏈衝突就撤銷該鏈的直接身份，而不是留下衝突鍵；6-K 與 10-Q 同等保護，缺完整鏈不猜季度排序，只拒絕與可信已結束年度衝突的 focus。人工查證更正僅限兩筆 KR，不追加未查原文的 COHR／CRM 特例。

Ruling：既有 CF 單季 Q2 冒充 YTD 的漏洞也納入安全防護。沒有直接可靠累計欄時不自行重建累計；流量留空、餘額保留、記缺口。399 個 period／fetcher／ledger 測試通過。重試 ContextVar 外漏及跨公司診斷殘留另在 `e3d9e62` 修正，390 個 fetcher／ledger 測試通過；這個 scope 修正不改無網路重試的 cached 數值流程。

Ruling：NOC 同申報同日期的分裂欄及 DELL 同日期不同申報不能以字典覆写，也不能跳过整家公司。比較器保留來源分組，對同概念同日期的全部非空值做保留重複次數的無序比較；不指定哪個來源為真。這是數值保留比較，不是替原表洗掉期間錯誤。`d6d1abe` 的 9 個比較器測試通過。

Ruling：沒有真正新手上下文就不冒稱已做 P2 黑箱；GUI 預覽仍是月份模型，未由這輪白箱核對證明一致。G13(a)、最新資料完整性、companyfacts 股數與 AI 診斷不在此輪 cached 驗收範圍。缺 authority 的 KR 最新部分年度仍走舊退路，不能宣稱 G13(b) 全解。

## 實際 Excel 驗證

在專屬新檔路徑建立 AAPL／CDNS／JNJ／KR 基準及候選活頁簿，使用自己建立的隱藏 Excel instance 重算並關閉，只修改測試檔。候選核對 **75,588 格資料矩陣及 row 1 預設標籤**，0 value／label mismatch；逐本修改 B4 月份的表頭結果也一致。基準 CDNS／JNJ／KR 的預設標籤與 pipeline 分別有 10／16／64 個差異。此驗證先按元→百萬等既有單位換算再比較；沒有把 raw 元與 Excel 百萬直接相比。

這個計數跳過 Fiscal Quarter／Calendar Quarter 兩列，也未包含另產生的 Data_Ratios；它證明上述輸出格與 pipeline／月份規格一致，**不證明所有來源期間和全部公式都正確**。後續更完整驗證若執行，另列結果。

## 中間驗證

- CF checkpoint `2ca3409`：385 個 fetcher／ledger 測試通過。
- 來源／Q4 checkpoint `ecb16ce`：18 個來源及 Q4 測試通過。
- Excel checkpoint `5e84de6`：90 個 fiscal input 測試通過。
- 檢查器 checkpoint `59b231e`：4 個判準測試通過；穩定月結公司也須驗證、失敗不回報成功、驗證只讀。
- 比較器 `99ea671`／`b21265e`：依實際期末日逐格比較，分開列出標籤、表頭及 metadata 差異；同一申報可供多個期間，不以發布日合併。5 個測試通過。
- 2026-10-06 第一輪完整非 slow 回歸：1,686 passed，65 deselected；後續新增測試另行驗證，最終會重跑完整套件。
- AEP、ACN 直接採季度 DEI 的中間版本有不合理回歸，故停止該候選量測，保留輸出但不作驗收依據。收緊日期鏈判準後，ACN 全部資料 0 差異；AEP 財季標籤 0 差異、僅舊邊界缺基準的 CF 流量改為空白。
- KR 修正後裸日期／重複期末日／錯序消失，但 CF 仍有 17 期缺基準。查快取發現第一季表格實際缺當期欄位（例如 2025-06-27 申報 CF 最新欄為 2025-02-01）。因此舊交接「只修 G13(b) 就會消除全部 KR CF 問題」不成立；G13(a) 恢復原 accession facts 仍是另一個必要工作，不能靠改標籤補數字。

## KR 來源更正證據

兩份文件 Item 1 皆有「All references to ... are to the fiscal years ended ...」的明確定義。2026 年年報也再次列出同一對照。更正登錄只匹配 accession、期末日、原 DEI 年份及 FY，不能套用到其他申報或其他原值；不修改 raw JSON。

| Accession | DEI → 正文財年 | SEC 原始文件 |
|---|---|---|
| 0001558370-24-004603 | 2024 → 2023；2024-02-03 | [KR 2023 10-K](https://www.sec.gov/Archives/edgar/data/56873/000155837024004603/kr-20240203x10k.htm) |
| 0001558370-25-004267 | 2025 → 2024；2025-02-01 | [KR 2024 10-K](https://www.sec.gov/Archives/edgar/data/56873/000155837025004267/kr-20250201x10k.htm) |

Ruling：採用上述逐筆人工查證更正，拒絕通用多數決修正；代價是其他來源錯誤仍需新證據、新登錄及測試。
# 2026-10-06 補充：稽核必須保留正式 listing metadata

HD 剩餘 2 格中，期末現金 overflow 原值 -388,000,000 是把兩期餘額相減；快取中的原始期末餘額是 2,719,000,000。相同標準概念重複列有一列被模板消耗、另一列落在 overflow，既有模板保護沒有涵蓋它。新增兩種確定為時點的標準現金概念白名單，overflow 保留原始餘額，不作 YTD 相減或因缺基期清空。自訂概念不猜測；Q4 overflow 仍沿用空白政策。兩個新測試確認有／無基期都正確，相關 383 項測試通過。

Ruling: 修正實際樣本揭露的餘額語義漏洞，保留所有來源列與原始資料，不藉刪除重複列隱藏問題。全庫驗收會包含此修正。

全庫候選比對曾顯示 EXC 984 個、HD 135 個財務儲存格改變。追查發現早期快取缺 cover 日期，稽核以 cover 代替正式 SEC reportDate，遺漏年度錨點；這份結果不能代表正式路徑。補上同一份官方 listing 後，EXC 降為 54 個變更且標籤零變更，HD 降為 2 個數值變更，期間資料不再被去重遺失。剩餘變更仍須分類。

Ruling: 全庫驗收改用 `--require-official-metadata`，改版前後使用完全相同的官方 accession/reportDate 清單；缺清單、缺 accession、無效非 6-K reportDate 一律失敗。cover 替代模式只保留為有限的離線診斷，不作正式全庫驗收。成本是必須重新跑兩側，不能沿用先前 215 家結論。

新增防護先觀察失敗，再通過 5 項稽核輸入測試；全套前一版為 1,706 passed、65 deselected。驗收仍進行中。

增量 reviewer 又重現兩個稽核問題：硬編碼 2009-06-15 與正式 `_XBRL_CUTOFF`（2008-01-01）不同；解析 `ValueError` 的例外分類可能觸發 SEC HEAD 探測。已改為使用所選 pipeline 的 cutoff，並注入 `FetchLedger(probe=lambda: True)`，本地解析失敗歸資料問題，禁止這個網路側路。新增三個 RED→GREEN 案例，37 項相關測試通過。這兩項改動沒有改變金融計算；驗收要檢查全部已用輸入的官方 metadata 覆蓋，以及例外路徑。

官方 metadata 覆蓋已獨立檢查全部 215 家、14,417 個 canonical 快取：以正式起始門檻及正式 cache gates 核對，0 個缺失／無效日期／不相容輸入。最新金融程式 `7f12a68` 正與基線 `edb6e5d` 用同一 adapter 重跑；輸出金融計算相同於目前 HEAD。若有本地解析例外，須以已禁止探測的新 adapter 重跑該公司；沒有例外的資料沒有進入探測路徑。

SEC 原始 instance 的 3 個年度／9 個月配對已保存於 [來源證據](evidence/fiscal-source-examples.json)。用相同年度起日的無維度 USD facts，核對 Q4 Revenue、OCF、Capex 支出金額、FCF、Ending Cash 共 15 格，早期候選全部符合。Capex 比支出金額絕對值；不能把 XBRL 正支出與 parser 呈現負號當成數字錯誤。最終全庫輸出仍需再次執行 `scripts/verify_fiscal_source_examples.py`。

擴充的自有 Excel 測試檔先用前一版保存輸入驗證 checker：95,156 格，含 16,048 格比率輸出及 3,279 格預設／覆寫標頭，0 discrepancy。這是輸出一致性，不是所有比率經濟定義的獨立來源驗證。最終候選四份 workbook 待完整輸入產出後重建、重算、再驗。

## 可重跑方式

Ruling: 正式 SEC 清單可用時才啟用新期間身份映射；季度或年度任一清單退回離線時，整次抓取保留原有期間命名。實際 `_OfflineFiling` 沒有 reportDate，會重現先前稽核的 EXC／HD 混用鍵風險；因此不能只修稽核卻放過產品離線路徑。代價是離線輸出不保證取得新的財年命名修正；現有 OFFLINE FALLBACK 提示繼續保留，資料庫不補寫猜測日期。兩種混合線上／離線情境皆有回歸測試，相關 412 項測試通過。完整 215 家官方清單比對不走此退回路徑，數值結果不受這個分支變更影響。

離線實際樣本：同一 adapter 模擬 listing 網路失敗，兩側重跑 EXC／HD 的全部保存輸入；兩家標籤與 header 差異都是 0。119 個數值差異中，55 格為有記錄缺基期的 CF 流量留空，64 格為標準期末現金 overflow 恢復餘額，未分類差異 0。最新完整測試為 **1,716 passed、65 deselected、4 warnings**；排除的是標記 slow 的網路測試。

使用外部 venv，先以 `scripts/fetch_fiscal_audit_metadata.py --all --output output/<新證據資料夾>` 下載官方清單；清單不覆寫，財報資料庫不寫入。之後 `scripts/audit_fiscal_pipeline.py --all --require-official-metadata --output output/<新證據資料夾>/candidate`。若要同介面跑歷史版本，設定 `SEC_AUDIT_PIPELINE_ROOT` 為基線 worktree 絕對路徑，輸出到同一證據資料夾的 `baseline`，完成後清掉此環境變數。

`scripts/compare_fiscal_audits.py <baseline> <candidate> --output <新比對目錄>` 會把失敗、缺公司、空資料列為非零 exit。`scripts/verify_fiscal_source_examples.py <candidate>` 額外核對原始 SEC facts 範例。
