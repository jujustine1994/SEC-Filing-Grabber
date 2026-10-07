# 共通規則、GUI 與 Excel 使用流程驗收（2026-10-07）

本輪依使用者優先序，處理跨公司使用規則與 GUI，未提高 KR／KHC 等個別例外的優先級。金融計算版本固定 `2147332`；最後 GUI 掃描鎖修復版本 `3f3615a`。後者僅改按鈕同步，不改財務計算。

**全庫重建、原值比較、GUI／Excel 驗收及來源雜湊均已完成。** 機器結果見 [驗收摘要](evidence/gui-common-rules-verification-summary.json)。本輪只改程式、模板呈現與文件，未寫回正式來源庫，未呼叫 LLM。

## 已完成修復

- 快速掃描與正式輸出共用 `_period_records`、`build_period_map`：包括原 cover focus、來源更正、年度錨點、衝突與期間身份碰撞；FPI 使用文件期末日。預覽只讀相容保存資料，不為身份下載歷史申報；缺可信身份時明示估計。
- 背景掃描結果帶 ticker；切換公司會清除舊分頁排除項，過期成功／失敗結果不套到新公司。掃描完成使用共通按鈕狀態，正式抓取仍在執行時不誤解除鎖。
- 已識別原始 GAAP revenue concept 的完整 assessed／excise／sales tax 排除語句不再誤列 NG；額外 adjusted／SBC／終止營業排除仍保留，未知自訂概念不能只憑稅字獲得 GAAP 身份。
- 缺口提示按連線、解析缺當期欄、無可靠單季基準及營收總計衝突分組；混合問題不全部稱為網路失敗，也不保證重抓能修復。四語言 GUI 與 Excel 共用摘要。
- Excel 起始月驗證為 1–12 整數；公式另防護貼上繞過 validation 的情況。只改主季／年表的顯示，不重新拆季、改財務值或改原 JSON；日曆季獨立。原日期不完整的欄位保留原標籤。

## 已完成驗收

- 全部離線測試 **1,845 passed / 65 deselected / 5 warnings**，266.75 秒；slow live 測試未跑。
- 215 家保存申報＋凍結官方 metadata 的最新預覽與正式輸出比對：**206 家可信身份一致、9 家估計顯示亦一致、0 不一致**。九家為 COHR、CRM、FICO、HON、ICE、LHX、LULU、ONTO、ORCL；不把一致的估計稱為原始來源認證。此驗收只涵蓋保存資料集合，不含尚未保存的新 SEC 申報，亦不驗 segment 掃描。使用年份篩選或僅輸出季／年報時，掃描「最新申報」與實際輸出範圍可以不同。
- 真正隱藏自有 Tk 視窗、隔離臨時庫、無 SEC I/O：按鈕啟動、鎖定、背景執行緒、queue、ticker 清除、過期結果拒絕、估計可見及完成恢復，七項檢查全部通過。此項驗證事件與元件狀態，不宣稱視覺排版截圖驗收。
- 真正自有 Excel instance：12 個合法月份＋0／13／−1／1.5／文字／空白共六種非法輸入全部通過；另核對預設申報年度身份、日曆季、財務值與不完整日期標籤保持。fixture 為獨立規格測試，不宣稱公司來源認證。
- 全 215 家以保存申報與凍結官方 metadata、`cached-read-only` 模式重建，22 批次全部 exit 0，含前後雜湊共 2,896.72 秒（約 48 分鐘）。沒有抓取新 SEC 申報；不代表日後 live 下載耗時。
- 與上一輪 Revenue 完成版比較：三表固定指標變更 **0**；主季／年表的期間標籤及表頭變更 **0**。EXC、FICO、MPC、UNP 僅有稅額排除營收的 GAAP／NG overflow 分類移動，NG 分頁因此有預期的刪減；不能把移動後的列號差異當原值改變。
- 215 家、403,499 個原始科目／季年／期間值袋跨分類核對，原值與重複次數差異 **0**。483 張財務表固定列全部通過。既有 25 個 Revenue 來源案例、舊七例及 15 格 fiscal／CF 來源驗證全部通過；不是全部儲存格逐一 XML 認證。
- 營收總計衝突 **0**；`MissingStandalonePeriod` 230 筆、`MissingCurrentPeriod` 86 筆，與上一版相同。數量是缺口紀錄，不是 316 個完整期間全部缺失。
- 11 份實際公司新 Excel：AAOI、AAPL、ARM、CMG、COP、COST、MPC、UNP、NEM、JNJ、LHX，以獨立隱藏 Excel COM 重算，核對 **270,522 格**（含比率 **39,530**、表頭 **9,074**），錯誤 **0**。UNP 一欄、NEM 一欄、LHX 十欄日期不完整，月份覆寫共 **12 欄 INCONCLUSIVE**；其原標籤保留，不視為來源日期認證。只操作本輪自有驗收工作簿。
- 正式庫 **29,277 檔、3,204,105,955 bytes**，前後 SHA-256 完全相同，新增／刪除／修改均 **0**；重跑前狀態亦與上一輪最終雜湊相同。

## 驗收工具與範圍

`scripts/probe_preview_gui.py` 建立隔離臨時庫與隱藏 Tk，實際跑 worker／queue，結束只關閉自己的視窗。`scripts/verify_saved_preview_periods.py` 以相同保存申報與正式 metadata 比對預覽；禁止來源寫入及 SEC 網路探測。

`scripts/verify_overflow_reclassification.py` 跨 GAAP／NG 分頁，以原 source concept、季／年、原期末日及非空值袋（保留重複次數、包含零）核對分類前後，不用顯示列號或任意 occurrence 配對。主表固定列、來源案例及全庫比較另用既有驗收工具。

完整 logs、輸出、Excel、SHA manifest 在忽略版控的 `output/gui-rules-*`，可攜式摘要已收錄於上述 evidence JSON。程式與文件均分步 commit；本輪未更新正式資料庫，未批次產生所有公司的 Excel。

重現時使用專案外 venv 與現行資料庫設定：先以 `audit_fiscal_pipeline.py --all --require-official-metadata` 在新驗收輸出目錄重建；用 `compare_revenue_audits.py` 比較舊／新結果，再以 `verify_overflow_reclassification.py OLD NEW --output REPORT --require-complete` 檢查跨分頁原值。預覽另用 `verify_saved_preview_periods.py`，GUI 事件另用 `probe_preview_gui.py`。來源全庫 SHA 仍應前後核對；不要用模板輸出相同替代來源庫雜湊，也不要把只讀保存資料重跑稱為更新所有最新申報。

## 保留界線

共通期間規則沿用既有拒絕衝突及不可靠拆季的保護；沒有因此認證全部原始日期與 context。九家來源不確定性、缺當期欄、缺可信年度錨點、transition duration 及併購口徑仍是來源／模型待辦，個別案例非最高優先。月份覆寫不能修好這些問題。GAAP／NG 分頁分類仍是呈現規則，不是每個自訂科目的會計口徑判讀。
