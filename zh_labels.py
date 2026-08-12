"""
zh_labels.py — 三表科目的中文說明（B 欄）。

Excel 版面：
    A 欄  英文標準名（程式內部用這個對映，**不要改**）
    B 欄  中文說明（本檔案，改這裡不影響任何邏輯）
    C 欄  公司原始 XBRL 標籤（看得出這格是從公司的哪個科目抓來的）
    D 欄起 各期數據

中文只是給人看的，程式一律用 A 欄的英文名做比對。所以這張表隨便你改用詞，
改錯也不會弄壞任何計算——最壞的情況只是 B 欄顯示怪怪的。

沒收錄的科目（例如 overflow 區的公司特有科目）B 欄留白。
"""

from __future__ import annotations

ZH_LABELS: dict[str, str] = {
    # ── 表頭 ────────────────────────────────────────────────────────────
    "財季 Fiscal Quarter":      "公司財年基準的季度",
    "日曆季 Calendar Quarter":  "日曆年基準的季度",
    "期末結算日 Period End":     "該期實際結束日（多為 52/53 週制，不是月底）",

    # ── 損益表 ──────────────────────────────────────────────────────────
    "Income Statement":            "損益表",
    "Revenue":                     "營業收入",
    "Cost of Revenue":             "營業成本",
    "Gross Profit":                "毛利",
    "R&D Expense":                 "研發費用",
    "SG&A Expense":                "銷售、管理及行政費用",
    "D&A (CF memo)":               "折舊與攤銷（取自現金流量表，備忘用）",
    "Other Operating Expense":     "其他營業費用",
    "Total Operating Expense":     "營業費用合計",
    "Total Costs and Expenses":    "成本與費用總計（含營業成本）",
    "Operating Income":            "營業利益",
    "Interest Expense":            "利息費用",
    "Interest Income":             "利息收入",
    "Other Non-op Inc/(Exp)":      "其他業外收入／(支出)",
    "Total Non-op Income/(Loss)":  "業外損益合計",
    "Pre-tax Income":              "稅前淨利",
    "Income Tax":                  "所得稅費用",
    "Net Income":                  "稅後淨利（歸屬母公司）",
    "Minority Interest":           "少數股權損益",
    "Net Income incl. NCI":        "稅後淨利（含少數股權，合併數）",
    "SBC":                         "股權獎酬費用（取自現金流量表）",
    "Basic EPS":                   "基本每股盈餘",
    "Diluted EPS":                 "稀釋每股盈餘",
    "Basic Shares":                "基本加權平均股數",
    "Diluted Shares":              "稀釋加權平均股數",

    # ── 資產負債表 ──────────────────────────────────────────────────────
    "Balance Sheet":                  "資產負債表",
    "Cash":                           "現金及約當現金",
    "Short-term Investments":         "短期投資",
    "Accounts Receivable":            "應收帳款",
    "Inventories":                    "存貨",
    "Other Current Assets":           "其他流動資產",
    "Total Current Assets":           "流動資產合計",
    "PP&E, net":                      "不動產廠房設備淨額",
    "Operating Lease ROU Assets":     "營業租賃使用權資產",
    "Long-term Investments":          "長期投資",
    "Goodwill":                       "商譽",
    "Intangible Assets, net":         "無形資產淨額",
    "Deferred Tax Assets":            "遞延所得稅資產",
    "Other Non-current Assets":       "其他非流動資產",
    "Total Non-current Assets":       "非流動資產合計",
    "Total Assets":                   "資產總計",
    "Accounts Payable":               "應付帳款",
    "Short-term Debt":                "短期借款",
    "Current Portion of LT Debt":     "一年內到期長期負債",
    "Op. Lease Liabilities, current": "營業租賃負債（流動）",
    "Accrued Compensation":           "應付薪資及員工福利",
    "Deferred Revenue, current":      "合約負債／遞延收入（流動）",
    "Income Tax Payable":             "應付所得稅",
    "Other Current Liabilities":      "其他流動負債",
    "Total Current Liabilities":      "流動負債合計",
    "Long-term Debt":                 "長期借款",
    "Op. Lease Liabilities, LT":      "營業租賃負債（非流動）",
    "Finance Lease Liabilities, LT":  "融資租賃負債（非流動）",
    "Deferred Revenue, LT":           "合約負債／遞延收入（非流動）",
    "Deferred Tax Liability, LT":     "遞延所得稅負債",
    "Pension & Retirement Oblig.":    "退休金及退職金負債",
    "Other Non-current Liabilities":  "其他非流動負債",
    "Total Non-current Liabilities":  "非流動負債合計",
    "Total Liabilities":              "負債總計",
    "Preferred Stock":                "特別股",
    "Common Stock & APIC":            "普通股股本及資本公積",
    "Additional Paid-in Capital":     "資本公積",
    "Treasury Stock":                 "庫藏股",
    "Retained Earnings":              "保留盈餘",
    "AOCI":                           "其他綜合損益累計額",
    "Total Equity — Parent":          "權益總計（歸屬母公司）",
    "Noncontrolling Interests":       "非控制權益",
    "Total Equity incl. NCI":         "權益總計（含非控制權益）",
    "Total Liabilities & Equity":     "負債及權益總計",
    "Shares Outstanding":             "期末在外流通股數（取自封面頁，日期比財季末晚數週）",

    # ── 現金流量表 ──────────────────────────────────────────────────────
    "Cash Flow":                        "現金流量表",
    "D&A":                              "折舊與攤銷",
    "Amortization of Intangibles":      "無形資產攤銷",
    "Change in Receivables":            "應收帳款變動",
    "Change in Inventories":            "存貨變動",
    "Change in Accounts Payable":       "應付帳款變動",
    "Change in Prepaid & Other Assets": "預付款項及其他資產變動",
    "Change in Other Operating Assets": "其他營業資產變動",
    "Change in Deferred Revenue":       "合約負債／遞延收入變動",
    "Other Working Capital":            "其他營運資金變動",
    "Other Non-cash Items":             "其他非現金項目",
    "Operating Cash Flow":              "營業活動現金流量",
    "Capex":                            "資本支出",
    "Acquisitions":                     "併購支出",
    "Investment Purchases":             "購買投資",
    "Investment Proceeds":              "處分投資所得",
    "Investing Cash Flow":              "投資活動現金流量",
    "Debt Proceeds":                    "舉借債務",
    "Debt Repayments":                  "償還債務",
    "Share Repurchases":                "庫藏股買回",
    "Dividends Paid":                   "支付股利",
    "Financing Cash Flow":              "籌資活動現金流量",
    "FX Effect on Cash":                "匯率對現金的影響",
    "Net Change in Cash":               "現金淨變動",
    "Ending Cash":                      "期末現金餘額",
    "Cash Taxes Paid":                  "實付所得稅",
    "Cash Interest Paid":               "實付利息",
    "Free Cash Flow":                   "自由現金流（營業現金流 − 資本支出）",

    # ── Data_Meta ───────────────────────────────────────────────────────
    "Ticker":                "股票代號",
    "Company Name":          "公司名稱",
    "Fetched Date":          "本檔資料抓取日（SEC 有 restatement，日期很重要）",
    "Quarters Available":    "抓到幾期",
    "Fiscal Year End Month": "財年結束月份（換算日曆季的依據）",
    "Key Rows 完整度":        "9 個關鍵科目中有幾個「最近 4 期至少一期有值」",
    "缺漏的 Key Rows":        "最近 4 期全空的關鍵科目",

    # ── overflow 區 ─────────────────────────────────────────────────────
    "Other (as reported)": "以下為該公司特有、不在固定模板中的科目（依原始標籤呈現）",
}


def zh_label(concept: str) -> str:
    """取中文說明。沒收錄回空字串——overflow 區的公司特有科目本來就沒有。"""
    return ZH_LABELS.get((concept or "").strip(), "")


# ── 維度軸的中文分類（Data_Segments 的 B 欄）─────────────────────────────
#
# XBRL 的分類細項掛在不同的「軸」上。沒有軸就分不出這一列是業務別營收，還是
# 權益變動表的項目別——MSFT 實測會混進 `Retained earnings`（權益項目軸）與
# `Service Life`（耐用年限軸），那些根本不是 segment。
#
# **不過濾、不丟棄**，只把軸標出來讓你自己篩。沒收錄的軸標成「其他維度」。
AXIS_LABELS: dict[str, str] = {
    "us-gaap:StatementBusinessSegmentsAxis":       "業務別",
    "srt:StatementGeographicalAxis":               "地區別",
    "us-gaap:StatementGeographicalAxis":           "地區別",
    "srt:ProductOrServiceAxis":                    "產品／服務別",
    "us-gaap:ProductOrServiceAxis":                "產品／服務別",
    "srt:ConsolidationItemsAxis":                  "合併沖銷別",
    "us-gaap:ConsolidationItemsAxis":              "合併沖銷別",
    "us-gaap:StatementEquityComponentsAxis":       "權益項目別（非 segment）",
    "srt:RangeAxis":                               "區間（非 segment）",
    "us-gaap:PropertyPlantAndEquipmentByTypeAxis": "固定資產類別（非 segment）",
    "srt:MajorCustomersAxis":                      "主要客戶別",
    "us-gaap:ConcentrationRiskByBenchmarkAxis":    "集中度基準別",
}


def axis_label(axis: str) -> str:
    """維度軸 → 中文分類。沒收錄回「其他維度」，不回空字串——
    空白會讓人以為沒有軸，實際上是有軸但我們沒收錄。"""
    axis = (axis or "").strip()
    if not axis:
        return ""
    return AXIS_LABELS.get(axis, "其他維度")
