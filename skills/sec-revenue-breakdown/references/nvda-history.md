# NVDA 全歷史案例

專用工具放在本 skill 的 scripts/：`build_nvda_history.py`、`prepare_nvda_segment_views.py`、`build_nvda_segment_history.mjs`。下載仍用共用 `collect_sources.py`。Python 抽取器需要 BeautifulSoup（bs4），本專案 Python 已具備；Excel builder 使用試算表 skill 的 artifact-tool runtime，不使用 repo 的套件。

先確定 repo 與 Python，指定同一個 repo 作為輸出位置：

```powershell
python <skill目錄>/scripts/collect_sources.py --cik 1045810 --start 1998-01-01 --end <指定申報截止日> --out <repo>/output/NVDA_history_sources
python <skill目錄>/scripts/build_nvda_history.py --repo <repo> --cache-repo <既有資料庫所在repo>
python <skill目錄>/scripts/prepare_nvda_segment_views.py --repo <repo>
```

已有 pinned selection 改截止日需另一資料夾或明確 `--refresh`。目前專用 extractor 讀固定 `output/NVDA_history_sources`；要新增季度，先更新此來源包。版面最新 taxonomy 目前固定到 FY2027Q2，更新公司新分類時需檢查/更新 prepare 工具，不可自動假設永遠相同。

Excel builder 在執行前遵循可用 spreadsheets skill：讀 API 與樣式規範、標記 authoring operation、在可寫的暫存目錄放 builder 與 artifact-tool node_modules junction，再執行 `node build_nvda_segment_history.mjs --repo <repo>`。不直接在共用 skill 裡裝 node_modules，也不修改 runtime。輸出 `NVDA_Revenue_Segment_History.xlsx`；遵循 workbook-format.md，檢查 Quarterly History、Annual History、Coverage、Sources，核對來源、空白與公式。

結果位於 `<repo>/output/NVDA_segment_history`：breakdown.json（來源、原報/比較數、分類registry）、checks.json、views.json、cell_sources.json、candidate_tables.json、Excel。來源包與這些檔案一起保存。原始證據不覆蓋，更新前保留上一版分類結果。

這次發現：

- 早期 Selected Financial Data 含 Product / Royalty，能追溯至 1994 年；這是 revenue_type，與 market_platform、reportable_segment、geography 分開。
- 1996/1997 財年止於 12 月，1998 有一個月過渡期間；不可套用現代 NVDA 財年規則。
- SEC 官方 CFO HTML 可直接抽取市場表及比較列，比 IR PDF 更適合 SEC 優先的來源鏈。
- 本地 income_statement 可能漏 OEM 等維度成員。加總不符標為 incomplete，保留差額，從 CFO 附件補找；不能自行生成 Other 補平。
- Verified 代表指定來源完整性/數值核對通過，不等於已找到所有公開細節。全歷史 Coverage 必須列出未取得可用拆分，仍需補查敘述型數字與未解析版型。

回歸核對：最新版 FY2027Q1 Hyperscale 43,050 / ACIE 32,196；Inactive Q1 原報 37,869 / 37,377。FY2027Q2 Hyperscale 48,710 / ACIE 40,313。Data Center + Edge 分別等於季度總額；市場產品 OP 未披露時留空。
