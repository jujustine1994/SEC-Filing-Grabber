# B Excel 空白版型

| ID | 檔案 | 適用 |
| --- | --- | --- |
| flat | Revenue_Breakdown_flat_formula_v1.xlsx | 無父子層級的平面披露；有階層時 builder 拒絕此選擇 |
| hierarchy | Revenue_Breakdown_hierarchy_formula_v1.xlsx | 父子分類／多個独立維度；也可處理平面區塊 |
| segment-op | Revenue_Breakdown_segment-op_formula_v1.xlsx | 報告部門及其已披露利益；沒有披露不填數字 |

三份是版型骨架，沒有任何公司數字。示例的來源／比重公式供理解結構，真實來源 IDs、期間与完整成長／利益率公式由 build_b_workbook.py 按研究資料產生；不能只替換公司名字就當完成品。

表名、固定位置與來源介面相同，差異為可用分類結構／指標示例。模板文字／格式可人工維護，但修改後須驗證並更新 catalog 的 SHA-256；builder 拒絕未知模板版本。範圍內格線可見、正黑體12以上、未隱藏列欄頁。驗證與研究缺口見 references/build-b-workbook.md。
