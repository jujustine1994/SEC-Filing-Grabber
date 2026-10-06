# 營收衝突：實際選值來源追蹤

> **本診斷的待修項目已完成。** 以下是修復前檢查點與原因；最新 215 家衝突為 0、880 筆留空恢復、25 例原 XML 與 Excel 驗證見 [完成驗證](revenue-resolved-verification-2026-10-07.md)。

本次是診斷檢查點，**營收修復尚未完成**。上一輪把安全拒絕選值、測試通過與完成營收校正混為一談，停在 `bdc930d` 太早。1770 個測試與 215 家能執行，不代表新增的營收空值都是正確的。

## 證據範圍

`scripts/trace_revenue_selection.py` 在唯讀 audit adapter 上追蹤實際送進 `_match_revenue_row` 的 accession、期間、原始 concept、label、計算父項和數值；只標記暫存 DataFrame 副本，不更新正式資料庫。

29 家追蹤 audit 與 `output/revenue-exact-final-all` 比對：財務值、期間標題、標籤與 gaps 沒有變化；唯一 metadata 差異是 Fetched Date 從 2026-10-06 到 2026-10-07。追蹤與詳細結果在忽略版控的 `output/revenue-open-traces`、`output/revenue-open-traces-audits`。

以下是實際執行使用的已保存解析輸入所顯示的原因，尚未宣稱每項已獨立對照原始 SEC XML。不能改用後年申報的比較欄來證明當年選值。

## 85 筆已列帳衝突

| 公司 | 筆數 | 實際原因與代表案例 |
| --- | ---: | --- |
| CMG | 7 | 競爭候選用不分大小寫的 `revenue\|sales` 子字串搜尋。`AvailableforsaleSecurities` 的 `sale` 加 `Securities` 恰好包含 `sales`，把其他綜合損益當成營收。2016 FY，accession `0001058090-17-000009`，Revenue 3,904,384,000 被與 OCI 1,402,000 誤列競爭。這是程式錯誤。 |
| HCA | 26 | 未識別扣除呆帳後的淨營收，讓 gross、provision、net 互相競爭。2017 FY，`0001193125-18-056057`：47,653,000,000 − 4,039,000,000 = 43,614,000,000；net concept 是 `HealthCareOrganizationPatientServiceRevenueLessProvisionForBadDebts`。2014 FY 的歷史 concept 是 `SalesRevenueServicesNet`。 |
| MS | 12 | `RevenuesNetOfInterestExpense` / Net revenues 未列入已識別總額；費用收入等子項卻參與候選。2017 FY，`0001193125-18-060831`：NoninterestIncome 34,645,000,000 與 net interest 3,300,000,000 的上層總額是 37,945,000,000。 |
| NEM | 2 | `RevenueMineralSales` / Total revenues, net 未識別為礦業總營收。2009 Q2，`0000950123-09-024448`：黃金 1,373,000,000 + 銅 229,000,000 = 1,602,000,000；子項的 label Revenue 不能取代總額。 |
| BK | 9 | Clearing fees 被標準化成 Revenue，實際仍只是手續費子項。必須核對銀行的 noninterest income 與 net interest income、是否存在報表總額，以及貸損提列前後口徑。不能直接取 gross interest，也不能只因加總吻合就擅自衍生總額。 |
| COP | 24 | `SalesRevenueNet` 與包含 Other Income 的 `Revenues` 範圍不同；程式刻意未處理此父子關係。2017 FY，`0001193125-18-049729`：29,106,000,000 與 32,584,000,000。需要確立模板 Revenue 的營業營收口徑。 |
| OXY | 4 | 同樣是 net sales 與 Total Revenues and Other Income 的口徑不同。2016 Q3，`0000797468-16-000039`：2,648,000,000 與 2,733,000,000。 |
| TGT | 1 | 2010-01-30 FY，`0001047469-10-002121`：Sales 63,435,000,000 + Credit card revenues 1,922,000,000 = Total revenues 65,357,000,000。解析輸入缺少父項，現有同層總額規則仍判衝突；需以原始完整總額標籤與 SEC 結構確認。 |

這 85 筆是 `AmbiguousRevenueTotal` ledger 數量，**不是全部空白 Revenue 儲存格數量**。上一輪相對舊版有 880 筆 Revenue 被留空，其中 `Net sales` 等未接受名稱可能是規則漏接；不得將全部留空宣稱為正確修復。

## 修復要求與完成判準

1. 修正競爭候選的詞彙辨識，避免跨單字子字串命中；不能繼續逐一補費用排除名單。
2. 用原始完整 concept、完整 label、合併範圍與計算父子關係辨認總額／淨額／子項。銀行利息不能套用一般公司的收入排除規則。
3. 對照實際 accession 的原始 SEC 資料，確認產業總額與 Other Income 政策；所有執行期規則必須固定、零 LLM 呼叫。
4. 同時檢查 85 筆衝突與 880 筆新增空值；保留真正缺 facts／缺当期欄的資料限制，禁止捏造值。
5. 修復後重新檢查 215 家、原始來源樣本、其他三表列不受影響，以及正式資料庫雜湊不變；每個完成階段 commit，更新交接文件。

不能以「測試通過、沒碰原始資料、模糊案例留空」單獨作為完成判準。需要證明應可辨認的營收已正確辨認。
