# Revenue 衝突修復與全庫驗收（2026-10-07）

**最終金融規則版本 `0ea4ec8`。** 原有 85 筆 `AmbiguousRevenueTotal` 已處理；最終 215 家輸出此類缺口為 **0**。先前完整名稱改版造成的 880 筆營收留空比較紀錄全部恢復。這些是模板輸出驗收，不代表所有儲存格都已逐一對照原始 XML，也不代表 G13 已全解。

機器統計、版本、來源範圍與 Excel 結果見 [驗收摘要](evidence/revenue-resolved-verification-summary.json)。本輪只修改程式與模板結果，正式資料庫 29,277 檔、3,204,105,955 bytes 的 SHA-256 前後完全相同，新增／刪除／變更均為 0。

## 決策及修復原因

| 案例 | 固定规则與口徑 |
| --- | --- |
| CMG | CamelCase 拆成完整單字；`AvailableforsaleSecurities` 不因跨字串含 sales 而與營收競爭。 |
| HCA | 辨識報告扣除呆帳 provision 後的淨營收及歷史 services concept，不把 gross／provision／net 三列任選一列。 |
| MS／BK／WFC／AXP／GS | 精確淨營收總計優先；銀行沒有總計時，可用非利息收入＋淨利息收入，採信用損失提列前口徑。 |
| NEM／DUK／APH／FIS | 加入有限產業／歷史原始 concept。goods/services 子項仍要查未覆蓋獨立營收，不因 concept 已知就認定為總額。 |
| TGT | 原始 Total revenues 含信用卡營業收入，優先於僅含銷售的 SalesRevenueNet；不依最大值猜測。 |
| COP／OXY／MPC／CVX／GE／PSX | Other Income aggregate 不取代營業銷售。相同父項下已識別、權重 +1 的營收子項才可構成完整群組；CVX 可含關係人收入，GE 可含商品、服務及金融服務。未知子項或缺必要值不補零。 |
| COP／CVX 新舊申報 | 一般總計及群組共用 concept registry，避免只修歷史類型卻漏掉現代 contract revenue。含／不含 assessed tax 是互斥選值，不能相加。 |
| JNJ | 不分大小寫排除精確 PercentToSales 後綴，0.003 比例不與 20.830B 營收金額競爭。 |
| XOM | 自訂 SalesAndOtherOperatingRevenueIncludingSalesBasedTaxes，在 Other Income 父項與 +1 計算關係下，以完整原始 `Sales and other operating revenue(s)` label 辨識，取申報含銷售稅原值，不取包含其他收益的合計。 |

同 concept 重複列相同值只算一次；同層不同值仍留空並報衝突。所有日期皆無值且無計算權重的空 presentation placeholder 可忽略，真實缺當期的組成仍保留。所有選值、衍生及缺口辨識均由固定程式執行；Revenue 不套歷史 override、不進 E1/E2，不呼叫 LLM。規則及資料庫分層見 [ARCHITECTURE](../ARCHITECTURE.md)。

## 驗收範圍

1. 固定 `e8239bd`、`68a8653` 各完成一次全 215 家重建，分別約 23.0／18.4 分鐘；全部使用同一份凍結官方 accession/reportDate，80 份季報／20 份年報限制，唯讀正式庫，停用 diagnosis、companyfacts、保存 hook 及 gap 網路探測。
2. 後續修復以 AST 驗證外層選值、維度、值過濾、名稱與 subtotal 覆蓋規則未變，掃描全部 14,417 份申報中的 13,155 張損益表。452 張相關表的 **全部日期欄**共 1,554 欄逐期比較。第一次僅 COP／CVX／JNJ／MPC／OXY／PSX／XOM 受影響並重建；最後完整標籤修復再掃全部輸入，只影響 XOM 並重建。因此最終結果包含 208 家經輸入影響證明不變的既有結果與七家更新結果，並非宣稱全部 215 家都在最後 commit 新啟程序。
3. 215 對 215、缺公司／失敗檔／比較例外為 0；489 張財務分頁固定列順序通過。原 SEC 七例、新增 25 例及 15 格 fiscal／CF 來源衍生結果全通過。新增來源案例獨立核對 XML SHA、CIK、原始 QName、無維度 duration、USD unit 與 72 條 role-qualified 計算關係；不是以輸出自證輸出。
4. 非 slow suite：**1,825 passed、65 deselected、5 warnings**，51.34 秒。Live slow 測試不作離線驗收；測試使用隔離臨時資料庫。
5. 新建 BK／CMG／COP／HCA／MS／NEM／TGT／MPC／CVX／GE／JNJ／XOM 共 12 份 Excel，以獨立自有 Excel instance 真正開啟、重算及保存；另測月份覆寫後關閉不存覆寫。401,234 格核對、含 47,849 格比率及 13,822 格表頭，錯誤為 0。NEM 一個來源 `2010-06` 缺日，覆寫來源身份列 INCONCLUSIVE，沒有改原日期。未批次產生全部 215 家 Excel。
6. 最終再逐檔雜湊正式庫，與本輪最初快照完全一致。完整 XML 僅下載到專案 `output/` 作抽樣证據；不補寫正式庫、不提高 parser version，也不失效既存解析資料。

## 全庫差異

數字是穩定 section／模板 slot／來源 key／期間身份的**比較紀錄數**，保留同期間非空值及重複次數，不以任意 occurrence 配對。

| 基準 | 有數值變更公司 | Revenue 季／年 | Gross Profit 季／年 | 其他固定指標 |
| --- | ---: | ---: | ---: | ---: |
| 上次完整名稱版 `bdc930d` | 33 | 1,048／266 | 326／83 | 0 |
| 完整名稱之前 `67e0a46` | 22 | 759／190 | 212／52 | 0 |

先前 880 筆由有值變空的 Revenue 比較紀錄已全部恢復。主季／年表的期間標籤與表頭無差異。MPC 保留原 contract revenue 構成列後新增兩張 NG overflow 分頁，帶出 76 個期間標籤與 228 個表頭比較差異；全部是新增分頁，並非改動主表財季。它的「excluding consumer excise taxes」被既有 NG 分類誤判，仍列 [TODO](../TODO.md)，未把分頁位置視為 Non-GAAP 來源證據。

`MissingStandalonePeriod=230`、`MissingCurrentPeriod=86` 與前版相同；這些是其他期間／資料缺口，不能因 Revenue 衝突變 0 就稱全部財報完整。KR 原 accession 當期 facts、缺可信年度錨點、transition duration、KHC predecessor/successor 及 GUI 一致性仍為 G13 待辦。

## 重現

使用專案外 venv `C:/Users/CTH/venvs/SEC Financial Tools/Scripts/python.exe`，不要用舊 `.venv`。正式資料庫為 `C:/Users/CTH/Documents/SEC財報資料庫`。完整 audit 需要本機 `output/TICKER-sec-listing.json` 的凍結官方清單；缺少時先取得正式 metadata，不能改用 cover date 偷換量測輸入。

```text
python -m pytest tests -q -p no:cacheprovider -m "not slow"
python scripts/download_revenue_resolution_sources.py output
python scripts/verify_revenue_resolution_sources.py output
python scripts/audit_fiscal_pipeline.py --all --require-official-metadata --output output/new-revenue-audit
python scripts/verify_revenue_resolution_audits.py output/new-revenue-audit
python scripts/verify_revenue_source_examples.py output/new-revenue-audit
python scripts/verify_fiscal_source_examples.py output/new-revenue-audit
python scripts/verify_fixed_financial_rows.py output/new-revenue-audit
python scripts/compare_revenue_audits.py output/revenue-exact-final-all output/new-revenue-audit --output output/new-revenue-comparison
```

完整輸出、SHA manifest、掃描明細、workbook 與 logs 在忽略版控的 `output/revenue-resolved-*`；可攜摘要與 25 例来源 fixture 已 commit。分步 commits 包含來源失敗例、各規則修復、影響掃描及文件，保留回溯點。
