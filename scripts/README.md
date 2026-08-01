# Scripts

開發與維護用的獨立腳本。不屬於主程式流程，可單獨執行。

## Index

| 腳本 | 說明 | 狀態 |
|------|------|------|
| `smoke_test_10.py` | 批次 live smoke test：10 間公司各抓 GAAP，自動檢查 Revenue/Gross Profit/Operating Income/Net Income/OCF/Capex/FCF 是否有值，輸出彙總表 | 啟用 |
| `survey_nongaap_metrics.py` | 調查 32 家（大中小型跨產業）8-K 新聞稿實際使用的 Non-GAAP 指標，統計跨公司覆蓋率，決定 `Data_NonGAAP` 固定模板要收哪些行。**不呼叫 AI**（純文字比對，不吃配額）。原文會存到快取目錄，調整比對規則後可重跑分析不必重新下載 | 啟用 |

---

## 測試方案對比：smoke_test_10.py vs tests/test_live_snapshots.py

| | `scripts/smoke_test_10.py` | `tests/test_live_snapshots.py` |
|---|---|---|
| **執行方式** | `python scripts/smoke_test_10.py` | `python -m pytest -m slow` |
| **用途** | 人工快查：最新季數值有沒有抓到 | 自動化迴歸：程式行為是否符合預期 |
| **公司** | AAPL/MSFT/TSLA/AMD/NVDA/GOOGL/META/WMT/COHR/AMZN（10 間） | MSFT/AMZN/META/GOOGL/NVDA/JPM/GS/JNJ（8 間）+ CF overflow 組（COHR/LITE/AAPL/NVDA/GOOGL） |
| **抓取筆數** | `max_filings=80`（完整，偵測長期資料） | `max_filings=8`（只抓最新 8 季，省時） |
| **輸出** | 彩色 terminal 表格，直接顯示數值 | pytest PASS/FAIL，失敗才顯示原因 |
| **判斷標準** | 最新季 7 個 key rows（Revenue/Gross Profit/Operating Income/Net Income/OCF/Capex/FCF）全非 None | key rows 近 4 季有 ≥1 非 None + B1 overflow 結構完整 + CF YTD 減法正確 |
| **耗時** | 較長（完整抓取） | 約 12 分鐘（8 季/ticker） |
| **適合場景** | 開發後手動驗收、懷疑某公司資料有問題時 | 改動 fetcher 後確認沒有迴歸 |
