# SEC Financial Tools 與 skill 的分工

本機專案位置是 `C:/Users/CTH/Documents/Code/SEC Financial Tools`。其他機器需重新定位 repo，不假設固定使用者路徑。本機 Python 是 `C:/Users/CTH/venvs/SEC Financial Tools/Scripts/python.exe`，工具也可用其他 Python 3 執行。

## 三層資料

1. `local_db/filing_cache/<TICKER>/<accession>.json`：既有解析財報。income_statement 含合併營收、OP 和部分分類 rows；保留 concept、dimension_axis、dimension_member、所有期間欄位。資料不是每個維度的完整披露，通常沒有 CFO commentary 的細節。單位常是美元而不是百萬美元，需要核對。cached_at 不是財報期間或更新日期。schema/parser version 是抽取版本，不證明內容完整。
2. `output/<TICKER>_sec_sources/manifest.json` 與 raw：完整 SEC 原文與附件，是新增分類和矛盾核對的證據。local_context 會驗證完整狀態、CIK、路徑及 SHA-256。
3. 使用者資料、公司 IR、AI 額外搜尋結果：用 `--supplement` 明確登記檔案，保留來源 URL/日期。AI 自己知道的數字不能作為已驗證數據；找到原始來源後才納入。supplement 是未驗證輔助資料，不自動覆蓋前兩層。

`Data_Segments` 是同一批快取產生的視圖，不是第二個獨立佐證；可閱讀既有 xlsx 幫助探索，但避免再讀全本試算表浪費上下文。若需要重建財務表，沿用專案既有 GAAP/segments builder 與 CLI，不另寫一套 Q4 財務算法。

## 離線匯出範例

```powershell
python <skill目錄>/scripts/local_context.py --repo "C:/Users/CTH/Documents/Code/SEC Financial Tools" --ticker NVDA --start 2024-11-01 --end 2026-10-05 --source-pack "output/NVDA_sec_sources" --supplement "output/nvda_revenue_pilot/dataset.json" --out "output/NVDA_local_context.json"
```

相對路徑以 shell 工作目錄為準；不在 repo 時使用絕對路徑。不具備 source pack 時先省略該參數匯出本地資料，再補來源。沒有本地 cache 時結果清單為空，改走官方來源流程，不表示公司未披露。

讀取 context 後先選目標季度的當期欄位，再保留比較期/年度/YTD 供推導及重編核對；不能對多份 filings 的所有欄位盲目 concat。以 CIK、accession、period、concept、dimension、presentation 定位資料，重分類的比較數另存版本。每筆使用的 local row 引用 source_id 加 concept/axis/member/期間欄位；原文驗證另掛 SEC source_id，不能把同源資料算成兩個獨立證據。

衝突記錄包含本地值、外部值、單位、口徑、來源、處理原因。公司 IR 更細分類可以補充，不可將市場營收和報告部門 OP 混用。結果仍須依 data-contract.md 核對加總和 OP%。
