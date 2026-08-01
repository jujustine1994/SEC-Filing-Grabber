"""
metric_rules.py — Non-GAAP 指標名稱規則表（**唯一可調整處**）

═══════════════════════════════════════════════════════════════════════════════
要調整 Data_NonGAAP 的行為，改這個檔案就好，不用動 fetcher_nongaap.py 或
excel_formatter.py。改完重跑即可生效——這些規則作用在「讀取快取」階段，
不是「寫入快取」階段，所以**不需要重新呼叫 AI、不需要刪 nongaap_cache.json**。
═══════════════════════════════════════════════════════════════════════════════

背景（2026-08-01）：
  原本的 AI prompt 用中文寫，AI 回中文指標名，但下游三條規則（期間 token 剝除、
  guidance 過濾、Excel ÷1M 豁免）全部只認英文，導致 Data_NonGAAP 整張表不可用：
  毛利率 37.5 被除成 3.75e-05、同一指標每季各自成列變成對角線。

  實查 nongaap_cache.json 後發現關鍵事實：**AI 回中還是回英是隨機的**，
  同一個 ticker 內都會混（CRM FY2026Q2 中文、FY2026Q1 英文）。所以採用方案 (c)：
  prompt 改要求英文（減少中文輸入），同時保留中英對照層當防線（AI 不聽話時接住），
  兩邊都做才能真正把同一指標合併成一列。

決策記錄（都可在本檔調整）：
  1. 百分比值存原始數字（37.5）而非 Excel 百分比（0.375）——理由：這張 sheet 是
     餵給下游 skill 的資料落地層，數字與 8-K 原文字面一致最好對帳。若要改成 Excel
     原生百分比，見 excel_formatter.PERCENT_AS_EXCEL_RATIO。
  2. 「服務毛利率」與「訂閱與服務毛利率」視為同一列（ARLO 的 AI 在不同季用了不同
     說法指同一個項目）。若你認為該分開，把 METRIC_ALIASES 裡
     "non-gaap services gross margin" 那行刪掉即可。
  3. 對照表沒收錄的名稱**原樣通過**，絕不丟棄——寧可多一列，不可少一筆資料。
"""

from __future__ import annotations

# ═══════════════════════════════════════════════════════════════════════════
# 1. 期間 token — 要從指標名剝除的字樣
# ═══════════════════════════════════════════════════════════════════════════
#
# 剝除的目的：讓「2024年第四季 Non-GAAP 毛利率」與「Non-GAAP 毛利率」認得出是
# 同一個指標。季度歸「當季桶」、年度歸「年度桶」，年度只在當季缺值時補洞。

# 季度樣式（歸當季桶）
#   2024年第四季 / 2025年第四季度 / 2026財年第三季度 / 第一季 / Q4 FY26 / FY26 Q4
ZH_QUARTER_PATTERNS = [
    r"\d{4}\s*年?\s*(?:財年|會計年度)?\s*第\s*[一二三四1-4]\s*季\s*度?",
    r"第\s*[一二三四1-4]\s*季\s*度?",
]

# 年度樣式（歸年度桶——只補洞，不可蓋掉當季值）
#   2024全年度 / 2025年全年度 / 2024年度 / 2026財年
ZH_ANNUAL_PATTERNS = [
    r"\d{4}\s*年?\s*全\s*年\s*度?",
    r"\d{4}\s*(?:財年|會計年度|年度)",
    r"全\s*年\s*度?",
]

# ═══════════════════════════════════════════════════════════════════════════
# 2. Guidance / 展望 — 含這些詞的整列直接丟棄
# ═══════════════════════════════════════════════════════════════════════════
#
# 這些是公司對「下一季」的預測，不是已實現數字，混進時間序列會造成錯誤結論。
#
# 英文用 startswith（歷史行為，維持不變）；中文用「包含」比對——因為中文的
# guidance 詞常出現在名稱中間，例如「2026財年預期 Non-GAAP 營業利潤率上限」，
# 用 startswith 會整批漏掉。

GUIDANCE_PREFIXES_EN = ("expected", "outlook", "guidance", "anticipated", "projected")

GUIDANCE_SUBSTRINGS_EN = ("outlook",)

GUIDANCE_SUBSTRINGS_ZH = (
    "預期", "預測", "預估", "預計",
    "指引", "展望", "目標",
    "低標", "高標", "上限", "下限",
)

# ═══════════════════════════════════════════════════════════════════════════
# 3. 中文 → 英文 詞彙替換（可組合）
# ═══════════════════════════════════════════════════════════════════════════
#
# 逐詞替換，長詞優先（程式會自動依長度排序，不必手動排）。可組合是重點：
#   「自由現金流」+「利潤率」→ "Free Cash Flow Margin"，不需要為組合另開一條。
#
# 新增詞彙時只要加一行；沒收錄的中文會原樣留在名稱裡（不會消失）。

ZH_TERMS = {
    # 毛利／利潤率類
    "訂閱與服務": "Subscription and Services",
    "毛利率":     "Gross Margin",
    "毛利":       "Gross Profit",
    "營業利潤率": "Operating Margin",
    "營業利益率": "Operating Margin",
    "營業利潤":   "Operating Income",
    "營業利益":   "Operating Income",
    "利潤率":     "Margin",
    "利潤":       "Income",
    "淨利率":     "Net Margin",
    "淨利":       "Net Income",
    "稅率":       "Tax Rate",

    # 每股類
    "稀釋每股收益": "Diluted EPS",
    "稀釋每股盈餘": "Diluted EPS",
    "攤薄每股盈餘": "Diluted EPS",
    "攤薄每股淨利": "Diluted EPS",
    "攤薄每股收益": "Diluted EPS",
    "每股盈餘":     "EPS",
    "每股收益":     "EPS",
    "每股淨利":     "EPS",

    # 現金流類
    "自由現金流": "Free Cash Flow",
    "營運現金流": "Operating Cash Flow",
    "現金流":     "Cash Flow",

    # 其他
    "調整後":     "Adjusted",
    "調整":       "Adjusted",
    "訂閱":       "Subscription",
    "服務":       "Services",
    "營收":       "Revenue",
    "收入":       "Revenue",
    "過去12個月": "LTM",
    "過去十二個月": "LTM",
}

# ═══════════════════════════════════════════════════════════════════════════
# 4. 同義名合併表（key 一律小寫、單一空白）
# ═══════════════════════════════════════════════════════════════════════════
#
# 上一步替換完之後跑這張表。用途有二：
#   (a) 統一大小寫與用詞（AI 的英文本身就不一致：Gross margin / Gross Margin）
#   (b) 把語意相同、說法不同的名稱併成同一列
#
# 沒收錄的名稱原樣通過。新增合併規則就加一行：把左邊（小寫）映到右邊（顯示名）。

METRIC_ALIASES = {
    # 每股盈餘——ARLO 六季用了四種說法指同一列
    "non-gaap eps":                            "Non-GAAP Diluted EPS",
    "non-gaap diluted eps":                    "Non-GAAP Diluted EPS",
    "non-gaap net income per share":           "Non-GAAP Diluted EPS",
    "non-gaap diluted net income per share":   "Non-GAAP Diluted EPS",
    "non-gaap earnings per share":             "Non-GAAP Diluted EPS",

    # 毛利率
    "non-gaap gross margin": "Non-GAAP Gross Margin",

    # 訂閱與服務毛利率——「服務毛利率」視為同一列（見檔頭決策 2）
    "non-gaap services gross margin":                  "Non-GAAP Subscription and Services Gross Margin",
    "non-gaap service gross margin":                   "Non-GAAP Subscription and Services Gross Margin",
    "non-gaap subscription and services gross margin": "Non-GAAP Subscription and Services Gross Margin",
    "non-gaap subscriptions and services gross margin": "Non-GAAP Subscription and Services Gross Margin",

    # EBITDA
    "adjusted ebitda":        "Adjusted EBITDA",
    "adjusted ebitda margin": "Adjusted EBITDA Margin",

    # 現金流
    "free cash flow":        "Free Cash Flow",
    "free cash flow margin": "Free Cash Flow Margin",

    # 損益
    "non-gaap revenue":          "Non-GAAP Revenue",
    "non-gaap operating income": "Non-GAAP Operating Income",
    "non-gaap operating margin": "Non-GAAP Operating Margin",
    "non-gaap net income":       "Non-GAAP Net Income",
}

# ═══════════════════════════════════════════════════════════════════════════
# 5. Excel 數值分類關鍵字（excel_formatter 用）
# ═══════════════════════════════════════════════════════════════════════════
#
# Data_NonGAAP 的值是 AI 直接從新聞稿抓的：金額是絕對數（30400000）、
# 百分比與每股數是原始小數（20.2 / 0.28）。三類的處理不同：
#
#   金額     → ÷1,000,000，顯示千分位（跟 GAAP 三表一致）
#   每股     → 不除，兩位小數
#   百分比   → 不除，顯示 "20.2%"
#
# 判斷靠 A 欄名稱比對這三組關鍵字（大小寫不敏感）。順序：每股 → 百分比 → 股數 → 金額。
# 加新關鍵字就加一行。**注意別讓百分比關鍵字誤傷金額行**：例如 "Adjusted EBITDA"
# 不含 margin 所以安全，但若哪天加了 "率" 以外的寬鬆詞就可能出事。

EPS_KEYWORDS = (
    "eps", "per share", "每股",
)

PERCENT_KEYWORDS = (
    "margin", "rate", "ratio", "yield", "growth %", "percentage",
    "率",           # 毛利率／利潤率／稅率
)

SHARES_KEYWORDS = (
    "shares", "股數",
)
