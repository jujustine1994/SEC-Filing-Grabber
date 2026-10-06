# 三表營收選值機械化研究（2026-10-06）

初始研究於 2026-10-06；後續已實作完整名稱規則與成本候選排除，最新全庫驗收見 [完整名稱驗證](revenue-exact-label-verification-2026-10-07.md)。原始 SEC 資料與獨立財報資料庫不修改。

## 後續實作：完整 label 匹配

使用者確認後，已實作有限範圍的精確名稱辨識：已知 GAAP concept 完整匹配保持原順序；未知 concept 接受原始 `Revenue(s)`、`Total Revenue(s)`、`Total Net Revenue(s)` 完整 label，先統一 NFKC、大小寫、標點與空白。裸 Revenue 有其他營收候選時保留衝突；匹配列不同金額也保留衝突，僅用當期有值列。沒有精確名稱不再退回標準化 Revenue 或 concept 子字串匹配。

Revenue 不再套用歷史 override，也不建立 E1/E2 自動修補。舊 override 檔案保留，沒有刪除；其他指標的 override 行為保持不變。COP/OXY/TGT 的已知 concept 衝突尚未在這一步解決。這次不新增來源下載或資料庫寫入，修改的是模板選值與診斷入口。

新增測試涵蓋裸 Revenue、完整 total 名稱、名稱正規化、分項／Other Income 不誤中、競爭候選與舊 override 不可繞過規則，以及即使 E2 旗標開啟也不得對 Revenue 發出 LLM 呼叫。原測試的縮略合約 concept fixture 改成完整 GAAP concept；其他欄位與 Q4 測試目標保持不變。

驗證：完整非 slow suite **1,763 passed / 65 deselected / 5 warnings**，48.39 秒；log 在忽略的 `output/revenue-exact-label-suite.log`。這一步尚未重跑 215 公司離線重建，先前大量樣本數值驗收不能當作新 label fallback 的驗收結果。

## 初始研究時的程式行為（後續實作以上方更新為準）

- `src/override_engine.py` 的 `E2_LLM_ENABLED = False`：GAAP 缺列診斷預設不呼叫 LLM，即使傳入 API key。`tests/test_override_engine.py` 已有不呼叫 `_llm_call` 的測試。
- E2 實作仍存在，修改旗標可啟用；模組開頭的 pipeline 說明也尚未明確反映預設停用。
- E1 是固定字串模糊比對，不花 AI token，但 Revenue 的廣泛同義字可能混淆總額、分項與不同營收口徑。零 token 並不保證選值正確。
- `load_overrides` 沒有依來源限制歷史 override；是否存在歷史 AI 產生的有效規則，仍需依實際儲存格式完整追溯，不能把關閉新呼叫當成歷史資料已排除。
- 目前 Revenue 選值使用原始 concept、非維度列及部分 calculation parent；不同候選值仍衝突時保留缺口。
- 儲存的 DataFrame 有 `parent_concept`，但不是完整且帶 statement role 的 XBRL calculation graph。安裝的 edgartools 會優先找該 statement role，再搜尋其他 roles；不能把單一 parent 欄位視為完整原始證據。

## 29 筆衝突與適用规则

以下是既有解析快取的觀察，尚未逐筆完成原始 SEC linkbase 複核。

| 公司 | 筆數 | 衝突 | 決定性處理方向 |
| --- | ---: | --- | --- |
| COP | 24 | Sales and other operating revenues 與 Total Revenues and Other Income | 明定模板營收口徑；以原始計算關係與組成項驗證，將非營業 other income 分開 |
| OXY | 4 | Net sales 與 Total revenues and other income | 同上；不得僅靠 larger value 或 label 子字串選值 |
| TGT | 1 | Sales 與 Total revenues，後者包含信用卡營收；快取沒有 parent | 查原始 calculation linkbase；信用卡營收可能是營業收入，不能一律排除 interest/fee income |

「營業收入」的口徑是產品政策，需要明文版本化。計算圖能證明加總關係，不能單獨決定模板應採用哪個口徑。不能對所有產業硬套 SalesRevenueNet 優先，也不能把所有 interest 都當非營業收入。

## 建議固定流程

1. 限定同 accession、報表 role、entity、期間、單位與維度範圍的候選。避免用其他季度、分部或另一張報表的關係補證。
2. 精確 concept 對照先行；保留原始 concept 與 label，不讓標準化 Revenue 名稱消除分項身份。
3. 利用原始 calculation arcs、權重與組成項，驗證總額與分項。依 decimals/rounding 處理容差與重複 facts，不能用任意百分比容差掩蓋差異。
4. 用明文的營收口徑與有限、經來源驗證的分類規則選值。規則適用範圍需包含 taxonomy/報表/期間條件；證據不足時使用狹義、有來源的 declarative exception，避免 ticker 全年代硬編碼。
5. 找不到唯一可驗證答案時，回報 `AmbiguousRevenueTotal` 與候選理由。禁止選最大值、第一列、模糊名稱或呼叫 AI 猜答案。僅有數字剛好相加不構成分類證明。
6. 每次結果記錄 rule id/version、候選、選中 concept、來源 accession、排除理由。相同來源與規則版本必須得到相同結果。

原始 linkbase 若需補抓，另存衍生證據區，不覆寫既有解析資料庫；來源修復與模板修復必須分開。

## 實作順序與驗收

1. 先確立 GAAP 選值的零 LLM 邊界：停用或移除 E2 生產入口、補全流程測試，讓 `_llm_call` 被呼叫就測試失敗；檢查歷史 override 來源，未知來源保留待審，不直接刪除。
2. 對 Revenue 禁止 E1 模糊匹配直接產生永久數值選擇，改為同一固定規則選值器。
3. 先取得 COP/OXY/TGT 原始計算與上下文證據，再建立固定規則及正反例；不能宣稱 29 筆已解決。
4. 重跑既有 215 公司離線重建與固定列稽核，核對非預期 IS/BS/CF 差異；確認正式資料庫檔案 hash 完全不變。保存前後缺口、命中規則與來源核對結果，逐步 commit。

## 官方資料依據

- [SEC：Calculation relationships 要求與缺漏提醒](https://www.sec.gov/rules-regulations/staff-guidance/disclosure-guidance/divisionscorpfinguidancexbrl-calculation)：SEC 接受申報不代表 calculation relationships 完整，所以關係缺失必須有明確 fallback。
- [XBRL Calculations 1.1](https://specifications.xbrl.org/work-product-index-calculations-2-calculations-1-1.html)：正式規範特別處理 rounded 與 duplicate facts，不能用簡單浮點相等當作全部驗證。

結論：這是固定規則、來源證據及模板口徑問題，無需執行期 AI。開發時研究規則與核對異常可以由人或 AI 協助，但正式三表選值應是可重現、可追溯且零 LLM token 的純程式流程。
