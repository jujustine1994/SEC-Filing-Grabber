---
name: sec-revenue-breakdown
description: 當使用者需要建立、更新或核對美國上市公司的營收拆分 B 報表、分類歷史或部門利益率時使用；來源包含 SEC 財報及公司官方法說資料。
---

# 官方營收拆分 B

依公司實際披露整理分類，保留原報、重編與原文定位；不預設每家公司有相同產品或部門。依要求交付來源包、研究 JSON 或 B Excel，C 同步另按公司任務處理。

## 1. 確認公司與既有成果

先讀公司 README，確認 ticker、CIK、期間與已有 B／證據。讀 [專案整合](references/project-integration.md)，沿用公司路徑與資料庫；已有來源先核對重用，缺文件、新申報或官方修訂才補抓。CIK 從 SEC 公司搜尋或 ticker 對照確認，不能猜。

## 2. 取得必要官方來源

| 來源 | 取得方式 |
|---|---|
| 本地解析財報 | `scripts/local_context.py` 離線探索營收／OP／分類候選，仍需核對原文 |
| SEC 10-Q、10-K、相關 8-K 與 EX-99 | `scripts/collect_sources.py` 取得正文及附件，細節見專案整合 |
| 公司 IR 財報新聞稿、法說投影片、CFO commentary | 用可用搜尋／瀏覽工具查官方 IR 財報頁，既有文件先讀，必要時下載原檔 |

下載器只接受 SEC 網址，不會自動巡查 IR。投影片若列為所選 SEC 申報的 EX-99 附件會一併取得；未附入 SEC、或分类／定義仍不足時，補查同一期官方 IR 資料。附件編號不保證內容，逐份確認；新聞稿、CFO commentary、segment 附註互補。同一文件的 SEC／IR 版本不算兩份獨立佐證。逐字稿按相關問題需要讀取，官方優先，第三方明示身分。

依共用規則保存、去重，登記官方 URL、來源層級、發布日、頁／表位置及雜湊。`--supplement` 只登記已保存檔案，不下載、不自動認證。IR 來源納入 builder 的 manifest 時，原檔也須落在其根目錄可驗證的範圍。下載或核對失敗明示缺口，不能因解析器没抽到就寫「未披露」。歷史實績與 guidance／預測分開。

## 3. 整理分類與可追溯數字

開始抽取讀 [資料契約](references/data-contract.md)；新增分類或處理重分類時讀 [分類生命週期](references/classification-lifecycle.md)。保存 breakdown、checks；需要 B 時再準備 views。每個數字有 record ID、期間、單位、口徑及原文定位。

產品、部門、地域各自整理，不相加。父子關係、版本、active／inactive 須有官方依據，不按名稱、縮排或加總猜。最新分類在上方，已確認停用分類放下方 Inactive Segments，保留歷史。缺一季不等於停用，最早取得值不等於生效日。缺值留 null，不補差額 Other；只有同口徑官方重編或可驗證推導才能延伸新分類歷史。資料庫或其他 AI 摘要不是獨立原文證據。

## 4. 需要 Excel 時建置與驗證

讀 [B 建置入口](references/build-b-workbook.md)，依披露結構選 flat／hierarchy／segment-op 模板，用共用 builder 另存新檔。精確版面與公式要求見 [工作表規範](references/workbook-format.md)。

固定 Quarterly History、Annual History、Coverage、Sources 四頁。Sources 保留原值，主表按 record ID 取值；比重、QoQ／YoY及已披露利益率由公式計算。CIK 缺少／非法、公司與來源包不符、引用的 SEC 網址 CIK 不符，builder 會拒絕。舊模板與 NVDA 專用 builder 是歷史入口，不用於新通用報表。

交付前核對來源雜湊、分類／口徑及數值，執行 `validate_workbook_format.py --plan`，再用原生 Excel 計算／保存驗證。未完成標 provisional 並列原因，格式 PASS 不代替財務核對。保留原成果與工具版本／時間，不自動覆寫正式 B／C。沒有新來源或分類變更時，核對後重用已驗證結果。

## 公司專用案例

重跑 NVDA 專用歷史抽取時才讀 [NVDA 案例](references/nvda-history.md)；公司特定分類與陷阱不是其他公司的通用分類器。
