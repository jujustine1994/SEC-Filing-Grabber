# 2026-10-13｜A 財務品質修正與 AMD 候選交接

本次負責修產生 A 的工具。README／windows-tool 規則、Git 狀態及共用 SOURCE_HANDOFF 已先讀；原有 `skills/sec-revenue-breakdown` 的未提交修改保留。本次不移動資料庫、不改連接、不覆寫正式 AMD A／B／C、不刪 C 官方補正。工具來源只讀現有資料庫與已下載原文，未補抓、未呼叫 AI API。

## 已確認根因與修改

| 問題 | 實際根因 | 共用修正 |
|---|---|---|
| Q4 負股數、EPS 不符 | `_synthesize_q4` 把 IS 的 EPS／weighted shares 一起當可加總流量；AMD 實際 cached pipeline FY2024Q4 重現 basic −3,235m、diluted −3,275m，EPS0.29／0.30與官方0.30／0.29不符 | 四個非加總鍵禁止相減；無官方值留空明示。完整 EPS local name 排除繼續營業子項；不由 NI/EPS 回推股數 |
| Total Liabilities 含權益 | std match 不成立後，`Liabilities` 子字串配合 last 選中 `LiabilitiesAndStockholdersEquity`；fuzzy／facts 也含錯誤候選 | 原始完整總負債概念／名稱；不受舊 override 繞過；缺總額時採同期間完整權益範圍的差額，有 NCI 缺完整權益則留空 |
| 營業成本漏攤銷 | excluding-D&A 子項與完整總成本同被 standard_concept 正規化為 CostOfGoodsAndServicesSold；first 取到子項 | 原始總額優先，計算父子排除，不任意對所有公司加攤銷。總額衝突／缺值不回退子項 |
| 稅前至淨利缺橋接概念 | source 保有權益法、繼續營業及停業稅後概念，但 IS 固定模板缺鍵；停業也可能被 overflow 分類到 NG | IS由24增27列，加入三個鍵與四語顯示。保留原文與源值，缺值不當零，Pre-tax Income 仍維持公司原本 before-equity 語意 |

初始交接的通用負股數例子 −4,942／−4,989 與本次 AMD 重現 −3,235／−3,275 分開記錄，不把兩者硬說成同一份 AMD 輸出。

## 官方單季補值

`--quarter-source-pack` 只讀既有 SEC source pack。驗證 issuer／URL accession／SHA256／路徑；HTML colspan 對齊 Three Months Ended 日期欄，支持共同 millions／thousands 單位，拒絕不支持的跨度、混合股數單位、Non-GAAP、EPS 子區、橋接子項與歸屬子範圍。只有六個已存在損益錨點全部與同日官方表一致才能填空格，不覆蓋報告值；正值 equity loss 依完整損失標籤轉為負收益，原始正值與標籤另存追溯。

選最早**可解析**的官方表，不能承諾一定是首次披露。AMD FY2020Q2／Q3 的所選 parser 表來自後續官方比較欄；與交接選定的原始披露值逐格一致，兩邊 accession 均記於核對 JSON。每格補值的实际來源保存在 Data_QuarterSources，含 Comparative；不將 C 補正 JSON 數字直接匯入候選 A。

## 驗證結果

- 最終完整離線測試：**1,880 passed，65 slow deselected，0 failures**。另有失敗後轉綠的34項本次專項回歸；包含獨立 reviewer 找出的期間、單位、EPS 分區、橋接子項、權益損失正負號及父項回退案例。
- AMD：官方 **32季／389格**與候選 pipeline 及輸出 Excel 一致；來源包40個可解析單季保留定位。八個 FY2018～FY2025 Q4 的股數／EPS均使用官方單季，不做全年減前三季。
- 原生 Excel：**81個期間表頭，公式錯誤0**，保存後再讀389格仍一致。七頁候選使用既有正式 writer；Data_QuarterSources 是新增追溯頁。
- 其他公司：AAPL、NVDA、MSFT、MRVL 全歷史保存輸出前後核對，期間未丟失；舊格變動分別64／128／74／20。合計212格為非加總 Q4 安全留空（這四家本次未提供官方單季包），另74格為原始同期間 BS 確認的總負債／非流動負債修正；其餘舊財務格不變。MSFT 初版11格成本空白退化已重現、加回歸並修復，最終沒有該退化。
- 本次讀取的五家公司財報、metadata 與 database.json 共**330檔**前後 SHA256 不變；資料庫保存／網路在驗收 runner 中禁止。這是指定輸入範圍，不稱為全庫所有 history／snapshot 雜湊審计。

| AMD 例子 | 修前 | 候選（USD million／million shares，EPS除外） |
|---|---:|---:|
| FY2024Q4 Basic Shares | −3,235 | 1,623 |
| FY2024Q4 Diluted Shares | −3,275 | 1,634 |
| FY2024Q4 Basic／Diluted EPS | 0.29／0.30 | 0.30／0.29 |
| FY2026Q2 Total Liabilities | 84,464 | 17,240 |
| FY2026Q2 Cost of Revenue | 5,073 | 5,333 |
| FY2026Q2 權益法／繼續營業淨利／稅後停業 | 缺固定鍵 | 6／2,305／−8 |

FY2026Q2 原口徑橋為 **2,551−252＋6−8＝2,297**；營收 **11,536−5,333＝毛利6,203**。同季總 Basic／Diluted EPS 為 **1.40／1.38**，不把 continuing EPS1.41／1.39當總EPS。

## 候選與證據

候選：`C:/Users/CTH/Documents/Code/SEC Financial Tools/output/amd-quality/accepted/AMD A candidate.xlsx`

原生 Excel 保存後 SHA256：`7508e7d810da8f8847f8ffa9cbae1a2664b5d06bd9f0bccfdb689660b226adb9`。

固定摘要：[a-financial-quality-2026-10-13.json](evidence/a-financial-quality-2026-10-13.json)。完整工作證據在 `output/amd-quality/accepted/`：official-source-records.json、official-checks.json、peer-checks.json、verification.json、native-excel.json、各公司 pipeline JSON。完整離線測試 log／JUnit 在 `output/amd-quality/final-suite.*`，原始 before 在 `output/amd-quality/before/`。output 為本機產物、不提交 Git。

重現：用 `scripts/verify_a_financial_quality.py --source-pack ... --official-evidence ... --output 新目錄 --baseline output/amd-quality/before --metadata-dir output`。原文 parser 單獨核對可跑 `scripts/verify_quarter_source_pack.py PACK EVIDENCE --cik 2488`；原生 Excel helper 與 finalize 腳本再驗證候選並更新最终雜湊。不得將 output 指向正式 A。

## 剩餘缺口與 C 銜接

保留9筆缺口：FY2010Q2 CF單季拆分、2013-03-31當期解析、FY2010～FY2015六個Q4股數／EPS、FY2017Q1官方表錨點口徑冲突。FY2017 A原報／B ASC606重編差異不強行對齊。本次只認證32季官方數字及四家公司抽樣回歸，不代表全歷史所有科目／全庫審計；GUI／compare沒有自動取得官方Q4原文的新流程，未提供單季來源的非加總值安全留空。

本 session **只交候選，未正式替換與同步 C**。後續由使用者決定採用後：

1. 保留正式原 A與來源／證據歷史，再決定正式替換或將 C SourceAPath 指向候選；不得自動刪除 B或C補正。
2. 同步只匯入 A來源副本，按英文機器鍵更新新增橋接列，執行 ValidateMappings；IS新增三列會使後續固定行號改變。
3. 保留 C現有32季官方補正，逐季比較新A、原文及C分析值。尤其成本已含成本攤銷，不再加一次；Total Liabilities已是負債，不再從該鍵扣權益。TL & Equity鍵仍保有完整合計，兩者語意要分清。Pre-tax before-equity口徑需搭配獨立權益法與停業稅後值，不雙算。
4. 驗證重同步、日期／季度、股數EPS、兩條會計橋、債務及人工預測保留；新增季度需要新增官方驗證，不能沿用32季PASS。只有逐項一致且另行決定後，才可改 C來源選擇／退役已冗餘補正，歷史定位與原文雜湊仍保留。
