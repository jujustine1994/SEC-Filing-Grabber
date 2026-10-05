---
name: sec-revenue-breakdown
description: 從 SEC 官方 10-Q、10-K 與 8-K 附件整理美國上市公司的多季營收拆分及營業利益率，保留來源、分類版本與核對結果；適用於重複執行或交接其他 AI。
---

# SEC 營收拆分

以公司實際揭露的細節建立分類，不預設所有公司有相同產品或部門。先取得可追溯來源，再做分類；下載工具不呼叫任何付費 AI API。

## 取得來源

先閱讀 [本專案資料整合](references/project-integration.md)。如果 SEC Financial Tools 專案可用，先執行 `scripts/local_context.py` 離線讀取本地營收/OP 與 XBRL 分類維度，再核對已下載的 SEC 來源包；缺少的附件或新申報才補抓。本地快取是衍生資料，不能當作完整原始附件或宣稱已驗證分類階層。可合併其他研究資料，但必須保留獨立來源與衝突紀錄。

用本 skill 的 `scripts/collect_sources.py`，只有 Python 標準函式庫依賴。指定公司 CIK、申報日期範圍及輸出資料夾。CIK 從 SEC 公司搜尋或 ticker 對照確認，不能猜。

歷年模式納入 `10-K405`／修訂版。2004-08-23 之前的 8-K 項目編號不一致（包含 Item 9、12、空白），全部保留為候選後確認是否財報，不僅依 Item 2.02 篩選，也不把所有舊 8-K 視為財報。SEC 清單的 primaryDocument 可能空白或已失效，此時取得 accession 的 SEC 完整申報 `.txt`，不能猜現代文件名。Plain-text / SGML 表格需另外解析，不可因 HTML extractor 找不到表格就認定未披露。歷史清單與實際已抽取、已核對的期間覆蓋分開回報。

```powershell
python <skill目錄>/scripts/collect_sources.py --cik 1045810 --start 2024-11-01 --end 2026-10-05 --out output/NVDA_sec_sources
```

User-Agent 從 `SEC_IDENTITY` 環境變數或 SEC Financial Tools 使用者設定讀取，不能把姓名、email 或 API key 寫進報告。工具保留申報索引、10-Q/10-K 主文件、8-K 主文件及所有 EX-99 附件；manifest 含 SEC URL、accession、日期、雜湊。固定日期、使用相同 manifest 可重跑；`--refresh` 明確更新申報清單，普通重跑沿用清單並驗證檔案完整性。SEC 阻擋或來源不完整時報告缺口，不能宣稱成功。

申報日期範圍不是財務季度範圍。取得後，依文件內的財報期間選出要求的季度，確定最新 10-Q/10-K 及其 earnings 8-K。範圍不足時擴大，不以湊足八個檔案代替八季。年報 Q4 可能要用全年減前三季累計。

## 整理與交接

開始抽取前閱讀 [資料與核對規範](references/data-contract.md)。先讀 8-K 的 EX-99.1、EX-99.2（新聞稿與 CFO commentary）；10-Q/10-K 補充 segment footnote。附件編號不保證內容，逐份確認。只把相關表格、分類變更段落與必要註解交給 AI，原始文件保留，不每次重讀整份報表。

每次產出 `breakdown.json`、`checks.json` 和簡短說明，使用引用到原始表格的資料契約。依使用者需求另產出 Excel；產出時使用可用的試算表 skill。不得把不同維度的拆分相加，缺值使用 null。驗證未過的結果標示 provisional，附具體缺口。允許只交付来源包，不捏造分類。

產出任何公司的 Excel 時，先複製 `assets/Revenue_Breakdown_Template.xlsx`，遵循 [跨公司固定工作表規範](references/workbook-format.md)：依序固定為 `Quarterly History`、`Annual History`、`Coverage`、`Sources`，不產出 Recent。內容優先保留原始揭露與逐筆來源；預設不算占比、QoQ、YoY，僅容許同口徑 OP% 與明確標示的簡單推導。分類依公司披露調整，不套用 NVDA 類別；用 `scripts/validate_workbook_format.py` 核對表名與順序。

已完成的結果重跑先驗證來源 SHA-256、契約與 checks；沒有新文件或分類變更時重用結果。更新時保留舊版本，原始季度口徑與後續重分類分開儲存，不無聲覆盖。結果必須記錄抽取/分類工具版本與時間；下載、檔案產生日期不是財報更新日期。

## 最新分類與 Inactive Segments

遵循使用者要求：每套營收拆分的最新分類在上方，停用的分類放下方 **Inactive Segments**，依分類版本分組。閱讀 [分類生命週期與呈現](references/classification-lifecycle.md) 後建立分類 registry。市場、報告部門、地域各自處理，不共用一個可相加的區塊。

Inactive 是分類停止披露，不是業務停止。未披露留空，不填 0；官方有重編值才回填到最新分類。舊分類原始值仍保留。只改名且有證據定義不變可沿用同列；相同名稱但範圍改變仍需另存版本。不能按最新文件出現的所有舊比較列判定它們仍 active，也不能把 active 與 inactive 合計。

## NVDA 已驗證的陷阱

需要重跑 NVDA 全歷史案例時，讀 [NVDA 歷史案例與命令](references/nvda-history.md)。該案例提供專用抽取器、最新／inactive 版面準備及 Excel builder；它們不是跨公司的通用分類器。

- market platform 和 reportable segment 是兩套拆分。Compute & Networking 的部門 OP 不可套到 Data Center、Compute 或 Networking。
- FY2027Q1 開始出現 Hyperscale、AI Clouds / Industrial / Enterprise、Edge Computing；不能假設能與 Gaming 等舊分類直接對接。
- FY2027Q2 重分類一家公司並重編比較期；QoQ 用同份當期文件的可比前期值，另外保留原始發布值。
- FY2027Q1 的 Hyperscale/ACIE 原報分類與 Q2 客戶重編分類需保留兩個版本，即使名稱相同。Q1 原報 Hyperscale 37,869、ACIE 37,377；Q2 對 Q1 的重編值是 43,050、32,196（USD millions）。前者在 Inactive Segments / original presentation，後者放最新分類。
- 部門 OP 可能排除未分攤費用；兩部門 OP 相加不必等於合併 GAAP OP。non-GAAP OP 必須另列。
- 公司 IR PDF 是官方公司來源但不是 SEC。缺少 SEC 對應附件時明確標示來源層級，不能稱其為 SEC 文件。

SEC 官方規格：[Submissions API](https://www.sec.gov/search-filings/edgar-application-programming-interfaces)、[Developer Resources / Fair Access](https://www.sec.gov/about/developer-resources)。
