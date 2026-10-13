# SEC Financial Tools 與 skill 的分工

## SEC 下載細節

用本 skill 的 `scripts/collect_sources.py`，只有 Python 標準函式庫依賴。指定公司 CIK、申報日期範圍及輸出資料夾。CIK 從 SEC 公司搜尋或 ticker 對照確認，不能猜。

歷年模式納入 `10-K405`／修訂版。2004-08-23 之前的 8-K 項目編號不一致（包含 Item 9、12、空白），全部保留為候選後確認是否財報，不僅依 Item 2.02 篩選，也不把所有舊 8-K 視為財報。SEC 清單的 primaryDocument 可能空白或已失效，此時取得 accession 的 SEC 完整申報 `.txt`，不能猜現代文件名。Plain-text / SGML 表格需另外解析，不可因 HTML extractor 找不到表格就認定未披露。歷史清單與實際已抽取、已核對的期間覆蓋分開回報。

```powershell
python <skill目錄>/scripts/collect_sources.py --cik 1045810 --start 2024-11-01 --end 2026-10-05 --out output/NVDA_sec_sources
```

User-Agent 從 `SEC_IDENTITY` 環境變數或 SEC Financial Tools 使用者設定讀取，不能把姓名、email 或 API key 寫進報告。工具保留申報索引、10-Q/10-K 主文件、8-K 主文件及所有 EX-99 附件；manifest 含 SEC URL、accession、日期、雜湊。固定日期、使用相同 manifest 可重跑；`--refresh` 明確更新申報清單，普通重跑沿用清單並驗證檔案完整性。SEC 阻擋或來源不完整時報告缺口，不能宣稱成功。

申報日期範圍不是財務季度範圍。取得後，依文件內的財報期間選出要求的季度，確定最新 10-Q/10-K 及其 earnings 8-K。範圍不足時擴大，不以湊足八個檔案代替八季。年報 Q4 可能要用全年減前三季累計。


本機專案位置是 `C:/Users/CTH/Documents/Code/SEC Financial Tools`。其他機器需重新定位 repo，不假設固定使用者路徑。本機 Python 是 `C:/Users/CTH/venvs/SEC Financial Tools/Scripts/python.exe`，工具也可用其他 Python 3 執行。

公司任務的資料夾命名、B 證據／SEC 原文／法說資料／逐字稿保存與去重，以 `C:/Users/CTH/.claude/project-rules/financial-assistant/financial_paths.json` 的 PUBLIC_WORKBOOK「公司作業資料夾的預設配置」及「B 來源蒐集與保存」為準，按需讀取；已有明確公司路徑直接沿用，不另建公司目錄。明確要求 B 包含所需官方資料的取得與分析，先核對既有來源、缺少才補抓。公司 IR／法說直接披露可支持分類；實績與 guidance 分開，來源未核對不升為 verified。以下三層資料與補充入口保留，不另抄一份存放規則。

## 三層資料

1. 獨立資料庫 `filings/<TICKER>/<accession>.json`：既有解析財報。透過 repo/src 的集中連接介面找資料庫，不拼接 repo/local_db。income_statement 含合併營收、OP 和部分分類 rows；保留 concept、dimension_axis、dimension_member、所有期間欄位。資料不是每個維度的完整披露，通常沒有 CFO commentary 的細節。單位常是美元而不是百萬美元，需要核對。cached_at 不是財報期間或更新日期。schema/parser version 是抽取版本，不證明內容完整。
2. `output/<TICKER>_sec_sources/manifest.json` 與 raw：完整 SEC 原文與附件，是新增分類和矛盾核對的證據。local_context 會驗證完整狀態、CIK、路徑及 SHA-256。
3. 使用者資料、公司 IR、AI 額外搜尋結果：用 `--supplement` 明確登記檔案，保留來源 URL/日期。AI 自己知道的數字不能作為已驗證數據；找到原始來源後才納入。supplement 是未驗證輔助資料，不自動覆蓋前兩層。

`Data_Segments` 是同一批快取產生的視圖，不是第二個獨立佐證；可閱讀既有 xlsx 幫助探索，但避免再讀全本試算表浪費上下文。若需要重建財務表，沿用專案既有 GAAP/segments builder 與 CLI，不另寫一套 Q4 財務算法。

## 離線匯出範例

```powershell
python <skill目錄>/scripts/local_context.py --repo "C:/Users/CTH/Documents/Code/SEC Financial Tools" --ticker NVDA --start 2024-11-01 --end 2026-10-05 --source-pack "output/NVDA_sec_sources" --supplement "output/nvda_revenue_pilot/dataset.json" --out "output/NVDA_local_context.json"
```

相對路徑以 shell 工作目錄為準；不在 repo 時使用絕對路徑。--repo 指程式位置，資料庫位置由連接設定決定；失聯會報錯，先用 cli.py db-connect 修復，不能當成公司未披露。不具備 source pack 時先省略該參數匯出本地資料，再補來源。連接有效但尚無該公司財報時清單才為空。重新解析歷史在庫內 history，單一 accession 的 current 與歷史不得當成兩個獨立披露佐證。操作與備份見專案 docs/DATABASE.md。

讀取 context 後先選目標季度的當期欄位，再保留比較期/年度/YTD 供推導及重編核對；不能對多份 filings 的所有欄位盲目 concat。以 CIK、accession、period、concept、dimension、presentation 定位資料，重分類的比較數另存版本。每筆使用的 local row 引用 source_id 加 concept/axis/member/期間欄位；原文驗證另掛 SEC source_id，不能把同源資料算成兩個獨立證據。

衝突記錄包含本地值、外部值、單位、口徑、來源、處理原因。公司 IR 更細分類可以補充，不可將市場營收和報告部門 OP 混用。結果仍須依 data-contract.md 核對加總和 OP%。
