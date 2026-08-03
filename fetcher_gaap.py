"""
fetcher_gaap.py — Fetch XBRL GAAP financial statements from SEC EDGAR via edgartools.

Fetches quarterly data from up to `max_filings` 10-Q filings (newest first),
and annual data from up to `max_annual_filings` 10-K filings (newest first).

Public API:
    fetch_gaap_statements(ticker, identity, max_filings=80, max_annual_filings=20) -> list[StatementTable]

Sheet outputs:
    Data_Financials(Q) — quarterly IS + BS + CF merged (from 10-Q)
    Data_Financials(Y) — annual IS + BS + CF merged (from 10-K)
    Data_Seg_*         — one sheet per IS concept that has segment breakdowns
    Data_Meta          — ticker / company / date / quarter count

StatementTable layout (A / B / C+):
    Col A  = Std Name (standardised display label)
    Col B  = Original Item (company's XBRL label from edgartools)
    Col C+ = quarterly data, oldest → newest
"""

from __future__ import annotations

import math
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import date, date as _date
from typing import Any

import pandas as pd
from edgar import Company, set_identity as set_identity

from override_engine import load_overrides, run_diagnosis, check_key_rows

# 9 個關鍵科目的清單（品質檢查用）。定義在 excel_formatter，避免兩處各記一份。
_ALL_KEY_ROWS_LAZY = None


# ── Data contract ────────────────────────────────────────────────────────

@dataclass
class StatementTable:
    """One financial statement pre-structured for Excel output."""
    sheet_name:     str
    quarter_labels: list[str]
    filing_dates:   list[str]
    concepts:       list[str]
    values:         list[list[Any]]
    ticker:         str = ""
    labels:         list[str] = field(default_factory=list)   # B-col: original XBRL labels
    # 各期真正的期末結算日（YYYY-MM-DD）。美股多為 52/53 週制，期末日不是月底
    # （ARLO FY2026Q1 結束在 03-29 不是 03-31），從財季標籤反推只能得到月份。
    # 來源是 XBRL 欄名 "2026-03-29 (Q1)"。沒帶到的表留空，下游自行退回反推。
    period_ends:    list[str] = field(default_factory=list)


# ── Constants ─────────────────────────────────────────────────────────────

META_COLS: set[str] = {
    'concept', 'label', 'standard_concept', 'level', 'abstract',
    'dimension', 'is_breakdown', 'dimension_axis', 'dimension_member',
    'dimension_member_label', 'dimension_label', 'unit', 'point_in_time',
    'balance', 'weight', 'preferred_sign', 'parent_concept', 'parent_abstract_concept',
}

# EDGAR began requiring XBRL from ~2009; filings before this date have no XBRL data.
# Filing lists are returned newest-first, so hitting a pre-cutoff filing means we can
# break the loop immediately rather than continuing through decades of empty filings.
_XBRL_CUTOFF: _date = _date(2008, 1, 1)


def _filter_filings_by_year(
    filings: list,
    start_year: int | None,
    end_year: int | None,
) -> list:
    """Filter filings list to only those within [start_year, end_year] (inclusive).

    Handles both date objects and ISO date strings ('YYYY-MM-DD').
    Returns filings unchanged when both bounds are None.
    """
    if start_year is None and end_year is None:
        return filings
    result = []
    for f in filings:
        fd = getattr(f, "filing_date", None)
        if fd is None:
            result.append(f)
            continue
        year = fd.year if isinstance(fd, _date) else int(str(fd)[:4])
        if start_year is not None and year < start_year:
            continue
        if end_year is not None and year > end_year:
            continue
        result.append(f)
    return result

# Tuple: (label, std_concept, fallback_suffix, source, match, label_hint)
#   label         — display name (Col A)
#   std_concept   — primary: standard_concept == value
#   fallback_suffix — secondary: concept contains value
#   source        — "IS" | "CF" | "BS" | "DERIVED"
#   match         — "first" | "last"  (which occurrence when multiple rows match)
#   label_hint    — tertiary filter: prefer rows whose label contains this string
_T = tuple[str, str | None, str, str, str, str | None]

IS_TEMPLATE: list[_T] = [
    ("Revenue",                    "Revenue",                        r"RevenueFromContractWithCustomer|SalesRevenueNet|SalesRevenueGoodsNet|_Revenues$|^Revenues$", "IS", "first", None),
    ("Cost of Revenue",            "CostOfGoodsAndServicesSold",     "CostOfGoodsSold",                                       "IS", "first", "cost"),
    ("Gross Profit",               "GrossProfit",                    "GrossProfit",                                            "IS", "first", None),
    ("R&D Expense",                "ResearchAndDevelopmentExpenses", "ResearchAndDevelopment",                                 "IS", "first", None),
    ("SG&A Expense",               "SellingGeneralAndAdminExpenses", "SellingGeneralAndAdmin",                                 "IS", "first", None),
    ("D&A (CF memo)",              "DepreciationExpense",            "DepreciationDepletionAndAmortization",                   "CF", "first", None),
    ("Other Operating Expense",    "OtherOperatingExpenses",         "OtherOperatingExpense",                                  "IS", "first", None),
    ("Total Operating Expense",    "TotalOperatingExpenses",         "OperatingExpenses",                                      "IS", "first", None),
    ("Total Costs and Expenses",   None,                             "^us-gaap_CostsAndExpenses$",                             "IS", "first", None),
    ("Operating Income",           "OperatingIncomeLoss",            "OperatingIncomeLoss",                                    "IS", "first", None),
    ("Interest Expense",           "InterestExpense",                "InterestExpense",                                        "IS", "first", None),
    ("Interest Income",            "InterestIncome",                 "InterestIncome",                                         "IS", "first", None),
    ("Other Non-op Inc/(Exp)",     None,                             "OtherNonoperatingIncome",                                "IS", "first", None),
    ("Total Non-op Income/(Loss)", "NonoperatingIncomeExpense",      "NonoperatingIncome",                                     "IS", "first", None),
    ("Pre-tax Income",             "PretaxIncomeLoss",               "IncomeLossFromContinuingOperationsBeforeIncomeTax",       "IS", "first", None),
    ("Income Tax",                 "IncomeTaxes",                    "IncomeTaxExpense",                                       "IS", "first", None),
    ("Net Income",                 "NetIncome",                      "NetIncomeLoss|NetIncomeLossAttributableToParent",         "IS", "first", None),
    ("Minority Interest",          None,                             "NetIncomeLossAttributableToNoncontrollingInterest",       "IS", "first", None),
    # 含少數股權的淨利。有 NCI 結構的公司會把「合併淨利」與「歸屬母公司淨利」分開報，
    # 上面的 Net Income 只認 NetIncomeLoss（歸屬母公司），這一列補的是合併數。
    ("Net Income incl. NCI",       None,                             "^us-gaap_ProfitLoss$",                                   "IS", "first", None),
    ("SBC",                        "StockBasedCompensationExpense",  "ShareBasedCompensation",                                 "CF", "first", None),
    ("Basic EPS",                  None,                             "EarningsPerShareBasic",                                  "IS", "first", None),
    ("Diluted EPS",                None,                             "EarningsPerShareDiluted",                                "IS", "first", None),
    ("Basic Shares",               "SharesAverage",                  "WeightedAverageNumberOfSharesOutstandingBasic",          "IS", "first", None),
    ("Diluted Shares",             "SharesFullyDilutedAverage",      "WeightedAverageNumberOfDilutedSharesOutstanding",        "IS", "first", None),
]

BS_TEMPLATE: list[_T] = [
    # ── Assets ──────────────────────────────────────────────────────────
    ("Cash",                           "CashAndMarketableSecurities",             "CashAndCashEquivalents",                                    "BS", "first", "cash and cash equivalents"),
    ("Short-term Investments",         "ShortTermInvestments",                    "ShortTermInvestments",                                      "BS", "first", None),
    ("Accounts Receivable",            "TradeReceivables",                        "AccountsReceivable",                                        "BS", "first", "accounts receivable"),
    ("Inventories",                    "Inventories",                             "Inventories",                                               "BS", "first", None),
    ("Other Current Assets",           "OtherNonOperatingCurrentAssets",          "OtherCurrentAssets",                                        "BS", "first", "other current"),
    ("Total Current Assets",           "CurrentAssetsTotal",                      "AssetsCurrent",                                             "BS", "first", None),
    ("PP&E, net",                      "PlantPropertyEquipmentNet",               "PropertyPlantAndEquipmentNet",                              "BS", "first", None),
    ("Operating Lease ROU Assets",     "OperatingLeaseRightOfUseAsset",           "OperatingLeaseRightOfUseAsset",                             "BS", "first", None),
    ("Long-term Investments",          "LongtermInvestments",                     "LongTermInvestments",                                       "BS", "first", None),
    ("Goodwill",                       "Goodwill",                                "Goodwill",                                                  "BS", "first", None),
    ("Intangible Assets, net",         "IntangibleAssets",                        "IntangibleAssetsNet",                                       "BS", "first", None),
    ("Deferred Tax Assets",            "DeferredTaxNoncurrentAssets",             "DeferredIncomeTaxAssetsNet",                                "BS", "first", None),
    ("Other Non-current Assets",       "OtherNonOperatingNonCurrentAssets",       "OtherAssetsNoncurrent",                                     "BS", "last",  "other"),
    ("Total Assets",                   "Assets",                                  "Assets",                                                    "BS", "last",  None),
    # ── Liabilities ─────────────────────────────────────────────────────
    ("Accounts Payable",               "TradePayables",                           "AccountsPayable",                                           "BS", "first", None),
    ("Short-term Debt",                "ShortTermDebt",                           "ShortTermBorrowings",                                       "BS", "first", None),
    ("Current Portion of LT Debt",     "CurrentPortionOfLongTermDebt",            "LongTermDebtCurrent",                                       "BS", "first", None),
    ("Op. Lease Liabilities, current", "OperatingLeaseCurrentDebtEquivalent",     "OperatingLeaseLiabilityCurrent",                            "BS", "first", None),
    ("Accrued Compensation",           "AccruedCompensation",                     "EmployeeRelatedLiabilitiesCurrent",                         "BS", "first", None),
    ("Deferred Revenue, current",      "OtherOperatingCurrentLiabilities",        "ContractWithCustomerLiabilityCurrent",                      "BS", "first", "unearned revenue"),
    ("Income Tax Payable",             "AccruedIncomeTaxes",                      "AccruedIncomeTaxesCurrent",                                 "BS", "first", None),
    ("Other Current Liabilities",      "OtherNonOperatingCurrentLiabilities",     "OtherLiabilitiesCurrent",                                   "BS", "first", None),
    ("Total Current Liabilities",      "CurrentLiabilitiesTotal",                 "LiabilitiesCurrent",                                        "BS", "first", None),
    ("Long-term Debt",                 "LongTermDebt",                            "LongTermDebt",                                              "BS", "first", "long-term debt"),
    ("Op. Lease Liabilities, LT",      "OperatingLeaseNonCurrentDebtEquivalent",  "OperatingLeaseLiabilityNoncurrent",                         "BS", "first", None),
    ("Finance Lease Liabilities, LT",  None,                                      "FinanceLeaseLiabilityNoncurrent",                           "BS", "first", "finance lease"),
    ("Deferred Revenue, LT",           "ContractLiabilities",                     "ContractWithCustomerLiabilityNoncurrent",                   "BS", "first", None),
    ("Deferred Tax Liability, LT",     "DeferredTaxNonCurrentLiabilities",        "DeferredIncomeTaxLiabilitiesNet",                           "BS", "first", None),
    ("Pension & Retirement Oblig.",    "PensionObligations",                      "PensionAndOtherPostretirementDefinedBenefitPlans",          "BS", "first", None),
    ("Other Non-current Liabilities",  "OtherNonOperatingNonCurrentLiabilities",  "OtherLiabilitiesNoncurrent",                                "BS", "first", None),
    ("Total Liabilities",              "Liabilities",                             "Liabilities",                                               "BS", "last",  None),
    # ── Equity ──────────────────────────────────────────────────────────
    ("Preferred Stock",                "PreferredStock",                          "PreferredStockValue",                                       "BS", "first", None),
    ("Common Stock & APIC",            "CommonEquity",                            "CommonStockValue",                                          "BS", "first", "common stock"),
    ("Additional Paid-in Capital",     "AdditionalPaidInCapital",                 "AdditionalPaidInCapitalCommonStock",                        "BS", "first", None),
    ("Treasury Stock",                 "TreasuryShares",                          "TreasuryStockValue",                                        "BS", "first", None),
    ("Retained Earnings",              "RetainedEarnings",                        "RetainedEarningsAccumulatedDeficit",                        "BS", "first", None),
    ("AOCI",                           "AccumulatedOtherComprehensiveIncome",     "AccumulatedOtherComprehensiveIncomeLossNetOfTax",            "BS", "first", None),
    ("Total Equity — Parent",          "AllEquityBalance",                        "StockholdersEquity",                                        "BS", "first", None),
    ("Noncontrolling Interests",       "MinorityInterestBalance",                 "MinorityInterest",                                          "BS", "first", None),
    ("Total Equity incl. NCI",         "AllEquityBalanceIncludingMinorityInterest","StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", "BS", "first", None),
    ("Total Liabilities & Equity",     "LiabilitiesAndEquity",                    "LiabilitiesAndStockholdersEquity",                          "BS", "first", None),
    # 期末在外流通股數（時點值）。與 IS 的 Basic/Diluted Shares 不同——那兩個是
    # 算 EPS 用的**加權平均**，在有買回或增發的季度會與期末股數差很多。
    ("Shares Outstanding",             "CommonSharesOutstanding",                 "CommonStockSharesOutstanding|EntityCommonStockSharesOutstanding", "BS", "last", None),
]

CF_TEMPLATE: list[_T] = [
    # ── Operating ────────────────────────────────────────────────────────
    ("Net Income",                 "NetIncome",                          "NetIncomeLoss|ProfitLoss",                              "CF", "first", None),
    ("D&A",                        "DepreciationExpense",                "DepreciationDepletionAndAmortization",                  "CF", "first", None),
    ("SBC",                        "StockBasedCompensationExpense",      "ShareBasedCompensation",                                "CF", "first", None),
    ("Amortization of Intangibles","AmortizationOfIntangibles",          "AmortizationOfIntangibleAssets",                        "CF", "first", None),
    ("Change in Receivables",      "ChangeInReceivables",                "IncreaseDecreaseInAccountsReceivable",                  "CF", "first", "receivable"),
    ("Change in Inventories",      None,                                 "IncreaseDecreaseInInventories",                         "CF", "first", "inventories"),
    ("Change in Accounts Payable",     None,  "IncreaseDecreaseInAccountsPayable",                          "CF", "first", None),
    ("Change in Prepaid & Other Assets", None, "IncreaseDecreaseInPrepaidDeferredExpenseAndOtherAssets",     "CF", "first", None),
    ("Change in Other Operating Assets", None, "IncreaseDecreaseInOtherOperatingAssets",                     "CF", "first", None),
    ("Change in Deferred Revenue", "ChangeInDeferredRevenue",            "IncreaseDecreaseInDeferredRevenue",                     "CF", "first", None),
    ("Other Working Capital",      "ChangeInOtherWorkingCapital",        "IncreaseDecreaseInOtherOperatingLiabilities",           "CF", "first", None),
    ("Other Non-cash Items",       "OtherNonCashItemsCF",                "OtherNoncashIncomeExpense",                             "CF", "first", None),
    ("Operating Cash Flow",        "NetCashFromOperatingActivities",     "NetCashProvidedByUsedInOperatingActivities",            "CF", "last",  "^net cash|^cash"),
    # ── Investing ────────────────────────────────────────────────────────
    ("Capex",                      "CapitalExpenses",                    "PaymentsToAcquirePropertyPlantAndEquipment",            "CF", "first", "property"),
    ("Acquisitions",               "AcquisitionsNet",                    "PaymentsToAcquireBusinessesNetOfCashAcquired",          "CF", "first", None),
    ("Investment Purchases",       "InvestmentPurchases",                "PaymentsToAcquireInvestments",                          "CF", "first", None),
    ("Investment Proceeds",        "InvestmentProceeds",                 "ProceedsFromSaleOfInvestments",                         "CF", "first", None),
    ("Investing Cash Flow",        "NetCashFromInvestingActivities",     "NetCashProvidedByUsedInInvestingActivities",            "CF", "last",  "^net cash|^cash"),
    # ── Financing ────────────────────────────────────────────────────────
    ("Debt Proceeds",              "DebtProceeds",                       "ProceedsFromIssuanceOfDebt",                            "CF", "first", None),
    ("Debt Repayments",            "DebtRepayments",                     "RepaymentsOfDebt",                                      "CF", "first", None),
    ("Share Repurchases",          "EquityExpenseIncomeBuybackIssued",   "PaymentsForRepurchaseOfCommonStock",                    "CF", "first", "repurchas"),
    ("Dividends Paid",             None,                                  "PaymentsOfDividends|PaymentsOfDividendsCommonStock|PaymentsOfOrdinaryDividends", "CF", "first", "dividend"),
    ("Financing Cash Flow",        "NetCashFromFinancingActivities",     "NetCashProvidedByUsedInFinancingActivities",            "CF", "last",  "^net cash|^cash"),
    # ── Other ────────────────────────────────────────────────────────────
    ("FX Effect on Cash",          "ForeignExchangeEffectOnCash",        "EffectOfExchangeRateOnCashAndCashEquivalents",          "CF", "first", None),
    ("Net Change in Cash",         "NetChangeInCash",                    "CashAndCashEquivalentsPeriodIncreaseDecrease",          "CF", "first", None),
    ("Ending Cash",                "CashAndCashEquivalents",             "CashAndCashEquivalentsAtCarryingValue",                 "CF", "last",  None),
    ("Cash Taxes Paid",            "IncomeTaxes",                        "IncomeTaxesPaid",                                       "CF", "first", "paid"),
    ("Cash Interest Paid",         "InterestExpense",                    "InterestPaid",                                          "CF", "first", "paid"),
    # ── Derived (computed, not from XBRL) ────────────────────────────────
    ("Free Cash Flow",             None,                                 "",                                                      "DERIVED", "first", None),
]

# ── Index maps for post-processing derived / fallback rows ────────────────

_IS_IDX: dict[str, int] = {row[0]: i for i, row in enumerate(IS_TEMPLATE)}
_NONOP_TOTAL_IDX   = _IS_IDX["Total Non-op Income/(Loss)"]
_OP_INCOME_IDX     = _IS_IDX["Operating Income"]
_PRETAX_IDX        = _IS_IDX["Pre-tax Income"]
_NET_INCOME_IDX    = _IS_IDX["Net Income"]
_DA_CF_IDX         = _IS_IDX["D&A (CF memo)"]
_REVENUE_IDX       = _IS_IDX["Revenue"]
_COGS_IDX          = _IS_IDX["Cost of Revenue"]
_GROSS_PROFIT_IDX  = _IS_IDX["Gross Profit"]

_CF_IDX: dict[str, int] = {row[0]: i for i, row in enumerate(CF_TEMPLATE)}
_CF_NET_INCOME_IDX      = _CF_IDX["Net Income"]
_CF_DA_IDX              = _CF_IDX["D&A"]
_CF_OP_CASH_IDX         = _CF_IDX["Operating Cash Flow"]
_CF_CAPEX_IDX           = _CF_IDX["Capex"]
_CF_FCF_IDX             = _CF_IDX["Free Cash Flow"]

_CF_INV_PURCHASES_IDX  = _CF_IDX["Investment Purchases"]
_CF_INV_PROCEEDS_IDX   = _CF_IDX["Investment Proceeds"]
_CF_DEBT_PROCEEDS_IDX   = _CF_IDX["Debt Proceeds"]
_CF_DEBT_REPAYMENTS_IDX = _CF_IDX["Debt Repayments"]

_INV_PROCEEDS_PATTERNS: list[str] = [
    r"ProceedsFromSaleOfInvestments",
    r"ProceedsFromSaleOfAvailableForSaleSecurities",
    r"ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities",
    r"ProceedsFromSaleAndMaturityOfMarketableSecurities",
    r"ProceedsFromSaleOfShortTermInvestments",
    r"ProceedsFromSaleMaturityAndCollectionOfShorttermInvestments",
]
_DEBT_PROCEEDS_PATTERNS: list[str] = [
    r"ProceedsFromIssuanceOfDebt$",
    r"ProceedsFromIssuanceOfLongTermDebt",
    r"ProceedsFromShortTermBorrowings",
    r"ProceedsFromLinesOfCredit",
    r"ProceedsFromIssuanceOfMediumTermNotes",
    r"ProceedsFromIssuanceOfSeniorLongTermDebt",
]
_DEBT_REPAYMENTS_PATTERNS: list[str] = [
    r"RepaymentsOfDebt$",
    r"RepaymentsOfLongTermDebt",
    r"RepaymentsOfShortTermDebt",
    r"RepaymentsOfLinesOfCredit",
    r"RepaymentsOfMediumTermNotes",
    r"RepaymentsOfSeniorDebt",
]


# ── Helpers ────────────────────────────────────────────────────────────────

def _col_to_quarter_label(col_name: str, fy_end_month: int = 12) -> str:
    """Convert edgartools period column name to FY label.

    fy_end_month: company's fiscal year end month (1-12). Default 12 = calendar year.
    For non-December FY companies, quarterly periods ending after fy_end_month belong
    to the next fiscal year (e.g. AAPL Sep FY: Dec 2023 Q1 → FY2024Q1).
    Annual (FY) labels are never adjusted.

    Examples (default fy_end_month=12):
        "2023-03-31 (Q1)"  -> "FY2023Q1"
        "2024-12-31 (FY)"  -> "FY2024"
    Examples (fy_end_month=9, AAPL):
        "2023-12-30 (Q1)"  -> "FY2024Q1"
        "2024-09-28 (FY)"  -> "FY2024"
    """
    m = re.match(r"(\d{4})-(\d{2})-\d{2}\s+\((\w+)\)", col_name.strip())
    if m:
        year, month, period = int(m.group(1)), int(m.group(2)), m.group(3)
        if period.upper() == "FY":
            return f"FY{year}"
        if fy_end_month < 12 and month > fy_end_month:
            year += 1
        return f"FY{year}{period}"
    return col_name


def _col_to_period_end(col_name: str) -> str:
    """edgartools 欄名 → 期末日。`"2026-03-29 (Q1)"` → `"2026-03-29"`。

    抓不到回空字串——下游會退回從財季標籤反推的年月。
    """
    m = re.match(r"(\d{4}-\d{2}-\d{2})\s+\(\w+\)", (col_name or "").strip())
    return m.group(1) if m else ""


def _detect_fy_end_month(filings_k: list) -> int:
    """Detect company's fiscal year end month from 10-K filings.

    Looks for a column labeled '(FY)' in the IS statement of the first 3 10-K filings.
    Returns the month number (1-12), defaulting to 12 (December) if not detected.
    """
    for filing in filings_k[:3]:
        try:
            tenq = filing.obj()
            is_stmt = tenq.financials.income_statement()
            if is_stmt is None:
                continue
            df = is_stmt.to_dataframe()
            for col in df.columns:
                if col in META_COLS:
                    continue
                mm = re.search(r"\d{4}-(\d{2})-\d{2}\s+\(FY\)", col)
                if mm:
                    return int(mm.group(1))
        except Exception:
            continue
    return 12


def _is_q_col(col_name: str) -> bool:
    """True if column is a quarterly period (Qx or FY), False for YTD."""
    m = re.search(r"\((\w+)\)", col_name)
    if not m:
        return False
    period = m.group(1).upper()
    return bool(re.match(r"Q\d+$", period)) or period == "FY"


def _current_q_col(df) -> str | None:
    """Return the first quarterly (non-YTD) period column from a filing's DataFrame."""
    for col in df.columns:
        if col in META_COLS:
            continue
        if _is_q_col(col):
            return col
    return None


def _ytd_col(df) -> str | None:
    """Return the first YTD period column (labeled '(YTD)'), or None."""
    for col in df.columns:
        if col in META_COLS:
            continue
        m = re.search(r"\((\w+)\)", col)
        if m and m.group(1).upper() == "YTD":
            return col
    return None


def _prev_quarter_label(label: str) -> str | None:
    """Return the label of the previous quarter, or None for Q1 / annual.

    Examples:
        "FY2025Q2" → "FY2025Q1"
        "FY2025Q1" → None
        "FY2025"   → None
    """
    m = re.match(r"(FY\d{4})Q(\d+)$", label)
    if not m:
        return None
    fy, q = m.group(1), int(m.group(2))
    if q <= 1:
        return None
    return f"{fy}Q{q - 1}"


def _consolidated_mask(df):
    """Boolean mask: non-abstract, non-breakdown, no dimension."""
    mask = ~df.get("abstract", False).astype(bool)
    mask &= ~df.get("is_breakdown", False).astype(bool)
    dim_col = df.get("dimension_member_label")
    if dim_col is not None:
        mask &= dim_col.isna() | (dim_col.astype(str) == "nan")
    return mask


# TODO: Verify fallback accuracy against ≥10 tickers.
#       Tested: AAPL, TSLA, BA, XOM (Session 9).
#       Still needed: MSFT, AMZN, META, GOOGL, NVDA, JPM, GS, JNJ.
def _match_is_row(df, std_concept: str | None, fallback_suffix: str,
                   label_fallback: str | None = None,
                   match: str = "first",
                   label_hint: str | None = None) -> int | None:
    """Find the row index in df matching a template entry.

    Priority order (each level only tried when previous level produces no usable result):
        1. standard_concept == std_concept (consolidated rows only)
        2. concept column contains fallback_suffix (case-insensitive, consolidated only)
        3. label column contains label_fallback (case-insensitive, consolidated only)

    label_hint: when candidates are found at a priority level, filter to those whose
        label contains this string.  If the filter leaves no candidates, the entire
        priority level is skipped and the next level is tried — candidates that fail
        label_hint are never returned as a fallback.
    match: "first" → earliest matching row; "last" → latest matching row.

    Returns None if no match found at any priority level.
    """
    mask = _consolidated_mask(df)
    df_c = df[mask]

    def _pick(rows) -> int | None:
        """Apply label_hint filter and return the selected index, or None if filtered out."""
        if rows.empty:
            return None
        if label_hint:
            hinted = rows[rows["label"].astype(str).str.contains(label_hint, case=False, na=False)]
            if hinted.empty:
                return None          # hint not satisfied → skip this priority level
            rows = hinted
        return rows.index[-1] if match == "last" else rows.index[0]

    # Priority 1: standard_concept exact match
    if std_concept:
        rows = df_c[df_c["standard_concept"].astype(str) == std_concept]
        result = _pick(rows)
        if result is not None:
            return result

    # Priority 2: concept contains fallback_suffix (supports regex OR via "|")
    if fallback_suffix:
        rows = df_c[df_c["concept"].astype(str).str.contains(fallback_suffix, case=False, na=False)]
        result = _pick(rows)
        if result is not None:
            return result

    # Priority 3: label contains label_fallback
    if label_fallback:
        rows = df_c[df_c["label"].astype(str).str.contains(label_fallback, case=False, na=False)]
        result = _pick(rows)
        if result is not None:
            return result

    return None


def _apply_row_override(df: pd.DataFrame, col: str, override_entry: dict) -> Any:
    """Look up a value from df using a pre-diagnosed override entry.

    concept_override: re-query df with the override's std_concept.
    structural_absence: return None immediately (confirmed missing in XBRL).
    """
    fix_type = override_entry.get("fix_type")
    if fix_type == "structural_absence":
        return None
    if fix_type == "concept_override":
        sc = override_entry.get("std_concept", "")
        if sc and col in df.columns:
            idx = _match_is_row(df, sc, sc)
            if idx is not None:
                return _to_python_val(df.loc[idx, col])
    return None


def _to_python_val(val) -> Any:
    """Convert pandas NA / float NaN / None to None; leave other values as-is."""
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    return val


def _row_key(df_row) -> str:
    """Unique key for a data row: concept + optional dimension_member_label."""
    concept = str(df_row.get("concept", "") or "")
    dim = str(df_row.get("dimension_member_label", "") or "")
    if dim and dim != "nan":
        return f"{concept}|{dim}"
    return concept


def _seg_sheet_suffix(concept: str, standard_concept: str | None) -> str:
    """Generate a ≤22-char alphanumeric suffix for a segment sheet name."""
    raw = standard_concept if standard_concept and standard_concept != "nan" else concept
    raw = re.sub(r"^[a-z_]+[_:]", "", raw)
    raw = re.sub(r"[^A-Za-z0-9]", "", raw)
    return raw[:22]


# ── Overflow helpers ──────────────────────────────────────────────────────────

# Labels containing any of these substrings are routed to the Non-GAAP overflow
# sheet instead of the GAAP overflow section.  Matching is case-insensitive.
_NONGAAP_KEYWORDS: frozenset[str] = frozenset({
    "non-gaap", "non gaap", "adjusted", "excluding", "excl.", "ex-",
})


def _is_nongaap_label(label: str) -> bool:
    """Return True if label looks like a Non-GAAP / adjusted metric."""
    low = label.lower()
    return any(kw in low for kw in _NONGAAP_KEYWORDS)


def _collect_overflow(
    df: pd.DataFrame,
    consumed: set[int],
    data_col: str,
    quarter_label: str,
    gaap_out: dict,
    ng_out: dict,
) -> None:
    """Collect unmatched XBRL rows from df into gaap_out or ng_out dicts.

    Rows whose index is in `consumed` are skipped (already captured by template).
    Abstract, breakdown, and dimension rows are excluded via _consolidated_mask.
    When a value is None the key is still recorded (so the concept appears in
    output even when it has no data for this specific quarter); periods dict
    only stores non-None values — the caller decides whether to drop all-None rows.
    """
    mask = _consolidated_mask(df)
    df_c = df[mask]
    remaining = df_c[~df_c.index.isin(consumed)]
    for _, row in remaining.iterrows():
        key = str(row.get("concept", "") or "")
        if not key or key == "nan":
            continue
        raw = str(row.get("label", "") or "")
        display = unicodedata.normalize("NFKC", raw)
        out = ng_out if _is_nongaap_label(display) else gaap_out
        if key not in out:
            out[key] = {"label": display, "periods": {}}
        val = _to_python_val(row.get(data_col))
        if val is not None:
            out[key]["periods"][quarter_label] = val


def _sum_matching_rows(
    df: pd.DataFrame,
    col: str,
    patterns: list[str],
    consumed: set[int],
) -> tuple[Any, list[int]]:
    """Sum values from consolidated rows whose concept matches any pattern in patterns.

    Skips rows already in consumed. Returns (total_or_None, list_of_matched_indices).
    """
    mask = _consolidated_mask(df)
    df_c = df[mask]
    total: float | None = None
    indices: list[int] = []
    seen_concepts: set[str] = set()
    for pattern in patterns:
        matches = df_c[df_c["concept"].astype(str).str.contains(pattern, case=False, na=False, regex=True)]
        for idx, row in matches.iterrows():
            if idx in consumed:
                continue
            concept = str(row.get("concept", "") or "")
            if concept in seen_concepts:
                continue
            seen_concepts.add(concept)
            val = _to_python_val(row.get(col))
            if val is not None:
                total = (total or 0.0) + val
                indices.append(idx)
    return total, indices


def _build_template_table(filings, template: list[_T], sheet_name: str,
                           stmt_method: str, max_filings: int,
                           fy_end_month: int = 12) -> StatementTable:
    """Generic fixed-template builder used by IS, BS, and CF."""
    periods: dict[str, tuple[str, dict[int, Any]]] = {}
    row_labels: dict[int, str] = {}   # first available original XBRL label per row

    for filing in filings:
        if len(periods) >= max_filings:
            break
        try:
            tenq = filing.obj()
            stmt = getattr(tenq.financials, stmt_method)()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
        except Exception as exc:
            print(f"[fetcher_gaap] {sheet_name} warning: {exc!r}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        label = _col_to_quarter_label(q_col, fy_end_month)
        if label in periods:
            continue

        row_vals: dict[int, Any] = {}
        for i, (_, std_concept, fallback, source, match, label_hint) in enumerate(template):
            if source == "DERIVED":
                row_vals[i] = None   # filled in post-processing
                continue
            idx = _match_is_row(df, std_concept, fallback,
                                 match=match, label_hint=label_hint)
            val = _to_python_val(df.loc[idx, q_col]) if idx is not None else None
            row_vals[i] = val
            if idx is not None and i not in row_labels:
                raw = str(df.loc[idx, "label"] or "")
                row_labels[i] = unicodedata.normalize("NFKC", raw)

        periods[label] = (str(filing.filing_date), row_vals)

    if not periods:
        return StatementTable(
            sheet_name=sheet_name,
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in template],
            values=[[] for _ in template],
            labels=["" for _ in template],
        )

    sorted_labels = sorted(periods.keys())
    filing_dates  = [periods[lbl][0] for lbl in sorted_labels]

    values: list[list[Any]] = []
    for i in range(len(template)):
        values.append([periods[lbl][1].get(i) for lbl in sorted_labels])

    labels_list = [row_labels.get(i, "") for i in range(len(template))]

    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=[row[0] for row in template],
        values=values,
        labels=labels_list,
    )


# ── IS: template-based fetch ────────────────────────────────────────────────

def _build_is_table(
    filings, max_filings: int, is_overrides: dict | None = None,
    fy_end_month: int = 12,
) -> tuple[StatementTable, StatementTable]:
    """Build IS StatementTables from 10-Q filings using the fixed IS template.

    Returns (gaap_tbl, ng_tbl):
      gaap_tbl — template rows + GAAP overflow rows (unmatched XBRL items)
      ng_tbl   — Non-GAAP overflow rows (labels containing "adjusted", "non-gaap", etc.)
                 Empty table when no Non-GAAP rows are found.
    """
    is_overrides = is_overrides or {}
    period_end_map: dict[str, str] = {}
    periods: dict[str, tuple[str, dict[int, Any]]] = {}
    row_labels: dict[int, str] = {}
    # Overflow dicts accumulate across all filings; key = XBRL concept name
    gaap_overflow: dict[str, dict] = {}
    ng_overflow:   dict[str, dict] = {}

    for filing in filings:
        if len(periods) >= max_filings:
            break
        _fd = getattr(filing, "filing_date", None)
        if isinstance(_fd, _date) and _fd < _XBRL_CUTOFF:
            break   # filings are newest-first; everything older is also pre-XBRL
        try:
            tenq = filing.obj()
            stmt = tenq.financials.income_statement()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
        except Exception as exc:
            print(f"[fetcher_gaap] IS warning: {exc!r}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        label = _col_to_quarter_label(q_col, fy_end_month)
        if label in periods:
            continue
        # 真正的期末結算日（52/53 週制不是月底）。從財季標籤反推只能得到月份，
        # Data_Std 需要精確日期，所以趁還看得到欄名先存下來。
        period_end_map[label] = _col_to_period_end(q_col)

        # Fetch CF statement for D&A / SBC rows
        cf_df: pd.DataFrame | None = None
        cf_q_col: str | None = None
        try:
            cf_stmt = tenq.financials.cashflow_statement()
            if cf_stmt is not None:
                cf_df = cf_stmt.to_dataframe()
                cf_q_col = _current_q_col(cf_df)
        except Exception:
            pass

        # Tracks which IS df indices are consumed by template matching this filing.
        # CF-sourced rows (source == "CF") consume cf_df indices — not tracked here.
        consumed: set[int] = set()

        row_vals: dict[int, Any] = {}
        for i, (row_name, std_concept, fallback, source, match, label_hint) in enumerate(IS_TEMPLATE):
            # Apply override if one exists for this row (concept_override or structural_absence)
            if row_name in is_overrides:
                ov = is_overrides[row_name]
                if ov.get("fix_type") == "structural_absence":
                    row_vals[i] = None
                    continue
                # concept_override: try in IS df; CF-sourced rows are not in KEY_ROWS so no IS overrides expected
                val = _apply_row_override(df, q_col, ov)
                row_vals[i] = val
                continue

            if source == "CF":
                # CF-sourced rows look in cf_df; those indices are NOT tracked in IS consumed set
                # (CF overflow is handled separately by _build_cf_table)
                if cf_df is not None and cf_q_col is not None:
                    idx = _match_is_row(cf_df, std_concept, fallback,
                                        match=match, label_hint=label_hint)
                    val = _to_python_val(cf_df.loc[idx, cf_q_col]) if idx is not None else None
                    if idx is not None and i not in row_labels:
                        raw = str(cf_df.loc[idx, "label"] or "")
                        row_labels[i] = unicodedata.normalize("NFKC", raw)
                else:
                    val = None
            else:
                idx = _match_is_row(df, std_concept, fallback,
                                    match=match, label_hint=label_hint)
                if idx is not None:
                    consumed.add(idx)   # mark as consumed so overflow skips this row
                val = _to_python_val(df.loc[idx, q_col]) if idx is not None else None
                if idx is not None and i not in row_labels:
                    raw = str(df.loc[idx, "label"] or "")
                    row_labels[i] = unicodedata.normalize("NFKC", raw)
            row_vals[i] = val

        # ── Post-processing: fallbacks not expressible in the 6-tuple ──

        # 1. Total Non-op: DERIVED = Pre-tax − Operating Income
        #    Guard: skip if discontinued operations present (would distort the difference)
        if row_vals.get(_NONOP_TOTAL_IDX) is None:
            op_val     = row_vals.get(_OP_INCOME_IDX)
            pretax_val = row_vals.get(_PRETAX_IDX)
            has_discontinued = _match_is_row(df, None, "DiscontinuedOperations") is not None
            if op_val is not None and pretax_val is not None and not has_discontinued:
                row_vals[_NONOP_TOTAL_IDX] = pretax_val - op_val

        # 2. Net Income fallback chain
        if row_vals.get(_NET_INCOME_IDX) is None:
            # 2a. Parent-only net income (more precise than ProfitLoss)
            idx = _match_is_row(df, "NetIncomeLossAttributableToParent",
                                 "NetIncomeLossAttributableToParent")
            if idx is not None:
                consumed.add(idx)
                row_vals[_NET_INCOME_IDX] = _to_python_val(df.loc[idx, q_col])
                if _NET_INCOME_IDX not in row_labels:
                    row_labels[_NET_INCOME_IDX] = unicodedata.normalize(
                        "NFKC", str(df.loc[idx, "label"] or ""))

        if row_vals.get(_NET_INCOME_IDX) is None:
            # 2b. ProfitLoss last resort (includes NCI — use only when parent-only unavailable)
            idx = _match_is_row(df, "ProfitLoss", "ProfitLoss")
            if idx is not None:
                consumed.add(idx)
                row_vals[_NET_INCOME_IDX] = _to_python_val(df.loc[idx, q_col])
                if _NET_INCOME_IDX not in row_labels:
                    row_labels[_NET_INCOME_IDX] = unicodedata.normalize(
                        "NFKC", str(df.loc[idx, "label"] or ""))

        # 3. D&A label fallback: for companies where standard_concept = nan (TSLA)
        #    This searches cf_df, not IS df — no IS consumed tracking needed
        if row_vals.get(_DA_CF_IDX) is None and cf_df is not None and cf_q_col is not None:
            idx = _match_is_row(cf_df, None, "", label_fallback="depreciation")
            if idx is not None:
                row_vals[_DA_CF_IDX] = _to_python_val(cf_df.loc[idx, cf_q_col])
                if _DA_CF_IDX not in row_labels:
                    row_labels[_DA_CF_IDX] = unicodedata.normalize(
                        "NFKC", str(cf_df.loc[idx, "label"] or ""))

        # 4. Gross Profit: DERIVED = Revenue − COGS (companies without explicit GP in XBRL)
        if row_vals.get(_GROSS_PROFIT_IDX) is None:
            rev  = row_vals.get(_REVENUE_IDX)
            cogs = row_vals.get(_COGS_IDX)
            if rev is not None and cogs is not None:
                row_vals[_GROSS_PROFIT_IDX] = rev - cogs

        # Collect unmatched IS df rows into overflow buckets
        _collect_overflow(df, consumed, q_col, label, gaap_overflow, ng_overflow)

        periods[label] = (str(filing.filing_date), row_vals)

    if not periods:
        empty = StatementTable(
            sheet_name="Data_IS",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in IS_TEMPLATE],
            values=[[] for _ in IS_TEMPLATE],
            labels=["" for _ in IS_TEMPLATE],
        )
        empty_ng = StatementTable(
            sheet_name="Data_IS_NG",
            quarter_labels=[], filing_dates=[],
            concepts=[], values=[], labels=[],
        )
        return empty, empty_ng

    sorted_labels = sorted(periods.keys())
    filing_dates  = [periods[lbl][0] for lbl in sorted_labels]

    # ── Build GAAP table: template rows + GAAP overflow ───────────────────
    concepts_g: list[str]       = [row[0] for row in IS_TEMPLATE]
    labels_g:   list[str]       = [row_labels.get(i, "") for i in range(len(IS_TEMPLATE))]
    values_g:   list[list[Any]] = [
        [periods[lbl][1].get(i) for lbl in sorted_labels]
        for i in range(len(IS_TEMPLATE))
    ]
    for key in sorted(gaap_overflow):
        entry = gaap_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue   # skip entirely-empty overflow rows
        concepts_g.append(entry["label"] or key)
        labels_g.append(key)
        values_g.append(row)

    period_ends = [period_end_map.get(lbl, "") for lbl in sorted_labels]

    gaap_tbl = StatementTable(
        sheet_name="Data_IS",
        quarter_labels=sorted_labels,
        period_ends=period_ends,
        filing_dates=filing_dates,
        concepts=concepts_g,
        labels=labels_g,
        values=values_g,
    )

    # ── Build NG table: Non-GAAP overflow rows only (no template rows) ────
    concepts_n: list[str]       = []
    labels_n:   list[str]       = []
    values_n:   list[list[Any]] = []
    for key in sorted(ng_overflow):
        entry = ng_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_n.append(entry["label"] or key)
        labels_n.append(key)
        values_n.append(row)

    ng_tbl = StatementTable(
        sheet_name="Data_IS_NG",
        quarter_labels=sorted_labels,
        period_ends=period_ends,
        filing_dates=filing_dates,
        concepts=concepts_n,
        labels=labels_n,
        values=values_n,
    )

    return gaap_tbl, ng_tbl


# ── BS: template-based fetch ────────────────────────────────────────────────

def _build_bs_table(filings, max_filings: int, bs_overrides: dict | None = None,
                    fy_end_month: int = 12) -> tuple[StatementTable, StatementTable]:
    """Build Data_BS StatementTable using the fixed BS template.

    Balance sheet columns in edgartools are instant (bare date, e.g. "2024-03-31")
    rather than period ("2024-03-31 (Q1)"), so _current_q_col cannot find them.
    We derive the quarter label from the IS statement (same filing) for merge alignment.

    Returns (gaap_tbl, ng_tbl). gaap_tbl = template rows + GAAP overflow;
    ng_tbl = Non-GAAP overflow rows only (sheet "Data_BS_NG").
    """
    bs_overrides = bs_overrides or {}
    periods: dict[str, tuple[str, dict[int, Any]]] = {}
    row_labels: dict[int, str] = {}
    gaap_overflow: dict[str, dict] = {}  # {concept_key: {"label": str, "periods": {q: val}}}
    ng_overflow: dict[str, dict] = {}

    for filing in filings:
        if len(periods) >= max_filings:
            break
        _fd = getattr(filing, "filing_date", None)
        if isinstance(_fd, _date) and _fd < _XBRL_CUTOFF:
            break   # filings are newest-first; everything older is also pre-XBRL
        try:
            tenq = filing.obj()

            # Get quarter label from IS (has "(Q1)"/"(FY)" format)
            is_stmt = tenq.financials.income_statement()
            is_df = is_stmt.to_dataframe() if is_stmt is not None else None
            is_q_col = _current_q_col(is_df) if is_df is not None else None

            bs_stmt = tenq.financials.balance_sheet()
            if bs_stmt is None:
                continue
            df = bs_stmt.to_dataframe()
        except Exception as exc:
            print(f"[fetcher_gaap] BS warning: {exc!r}", file=sys.stderr)
            continue

        # BS columns are bare dates; pick first non-meta column
        bs_col = next((c for c in df.columns if c not in META_COLS), None)
        if bs_col is None:
            continue

        label = _col_to_quarter_label(is_q_col, fy_end_month) if is_q_col else _col_to_quarter_label(bs_col, fy_end_month)
        if label in periods:
            continue

        consumed: set[int] = set()
        row_vals: dict[int, Any] = {}
        for i, (row_name, std_concept, fallback, source, match, label_hint) in enumerate(BS_TEMPLATE):
            if source == "DERIVED":
                row_vals[i] = None
                continue
            if row_name in bs_overrides:
                ov = bs_overrides[row_name]
                if ov.get("fix_type") == "structural_absence":
                    row_vals[i] = None
                    continue
                row_vals[i] = _apply_row_override(df, bs_col, ov)
                continue
            idx = _match_is_row(df, std_concept, fallback, match=match, label_hint=label_hint)
            val = _to_python_val(df.loc[idx, bs_col]) if idx is not None else None
            row_vals[i] = val
            if idx is not None:
                consumed.add(idx)
                if i not in row_labels:
                    raw = str(df.loc[idx, "label"] or "")
                    row_labels[i] = unicodedata.normalize("NFKC", raw)

        _collect_overflow(df, consumed, bs_col, label, gaap_overflow, ng_overflow)
        periods[label] = (str(filing.filing_date), row_vals)

    empty_ng = StatementTable(
        sheet_name="Data_BS_NG", quarter_labels=[], filing_dates=[],
        concepts=[], values=[], labels=[],
    )
    if not periods:
        return StatementTable(
            sheet_name="Data_BS",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in BS_TEMPLATE],
            values=[[] for _ in BS_TEMPLATE],
            labels=["" for _ in BS_TEMPLATE],
        ), empty_ng

    sorted_labels = sorted(periods.keys())
    filing_dates = [periods[lbl][0] for lbl in sorted_labels]

    # ── Build GAAP table: template rows + GAAP overflow ─────────────────────
    concepts_g: list[str]       = [row[0] for row in BS_TEMPLATE]
    labels_g:   list[str]       = [row_labels.get(i, "") for i in range(len(BS_TEMPLATE))]
    values_g:   list[list[Any]] = [
        [periods[lbl][1].get(i) for lbl in sorted_labels]
        for i in range(len(BS_TEMPLATE))
    ]
    for key in sorted(gaap_overflow):
        entry = gaap_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_g.append(entry["label"] or key)
        labels_g.append(key)
        values_g.append(row)

    gaap_tbl = StatementTable(
        sheet_name="Data_BS",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_g,
        labels=labels_g,
        values=values_g,
    )

    # ── Build NG table: Non-GAAP overflow rows only (no template rows) ───────
    concepts_n: list[str]       = []
    labels_n:   list[str]       = []
    values_n:   list[list[Any]] = []
    for key in sorted(ng_overflow):
        entry = ng_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_n.append(entry["label"] or key)
        labels_n.append(key)
        values_n.append(row)

    ng_tbl = StatementTable(
        sheet_name="Data_BS_NG",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_n,
        labels=labels_n,
        values=values_n,
    )
    return gaap_tbl, ng_tbl


# ── CF: template-based fetch ────────────────────────────────────────────────

def _build_cf_table(filings, max_filings: int, cf_overrides: dict | None = None,
                    fy_end_month: int = 12) -> tuple[StatementTable, StatementTable]:
    """Build Data_CF StatementTable using the fixed CF template.

    Q1 and FY filings have standalone period columns (Q1/FY) and are used directly.
    Q2 and Q3 filings have YTD (cumulative) CF columns; standalone quarter values are
    derived by subtracting the prior period's YTD: Q2 = Q2_YTD − Q1, Q3 = Q3_YTD − Q2_YTD.
    Overflow rows use the same subtraction logic: raw values are collected per filing and
    converted to standalone after the filing loop.

    Returns (gaap_tbl, ng_tbl). gaap_tbl = template rows + GAAP overflow;
    ng_tbl = Non-GAAP overflow rows only (sheet "Data_CF_NG").
    """
    cf_overrides = cf_overrides or {}
    # collected: label → (filing_date, {row_i: raw_value}, is_ytd)
    collected: dict[str, tuple[str, dict[int, Any], bool]] = {}
    # ytd_raw: stores raw values (standalone for Q1/FY, cumulative for YTD)
    # used as the subtraction base for the next YTD period
    ytd_raw: dict[str, dict[int, Any]] = {}
    row_labels: dict[int, str] = {}
    gaap_overflow: dict[str, dict] = {}  # {concept_key: {"label": str, "periods": {q: val}}}
    ng_overflow: dict[str, dict] = {}
    # Raw overflow per filing; YTD subtraction applied after loop (mirrors template logic)
    # {q_label: {concept_key: (display_label, is_nongaap, raw_val)}}
    overflow_per_filing: dict[str, dict] = {}

    for filing in filings:
        if len(collected) >= max_filings:
            break
        _fd = getattr(filing, "filing_date", None)
        if isinstance(_fd, _date) and _fd < _XBRL_CUTOFF:
            break   # filings are newest-first; everything older is also pre-XBRL
        try:
            tenq = filing.obj()
            is_stmt = tenq.financials.income_statement()
            is_df = is_stmt.to_dataframe() if is_stmt is not None else None
            is_q_col = _current_q_col(is_df) if is_df is not None else None

            cf_stmt = tenq.financials.cashflow_statement()
            if cf_stmt is None:
                continue
            df = cf_stmt.to_dataframe()
        except Exception as exc:
            print(f"[fetcher_gaap] CF warning: {exc!r}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is not None:
            label = _col_to_quarter_label(q_col, fy_end_month)
            if label in collected:
                continue
            is_ytd = False
            data_col = q_col
        else:
            ytd_col = _ytd_col(df)
            if ytd_col is None or is_q_col is None:
                continue
            label = _col_to_quarter_label(is_q_col, fy_end_month)
            if label in collected:
                continue
            is_ytd = True
            data_col = ytd_col

        consumed: set[int] = set()
        row_vals: dict[int, Any] = {}
        for i, (row_name, std_concept, fallback, source, match, label_hint) in enumerate(CF_TEMPLATE):
            if source == "DERIVED":
                row_vals[i] = None
                continue
            if row_name in cf_overrides:
                ov = cf_overrides[row_name]
                if ov.get("fix_type") == "structural_absence":
                    row_vals[i] = None
                    continue
                row_vals[i] = _apply_row_override(df, data_col, ov)
                continue
            idx = _match_is_row(df, std_concept, fallback, match=match, label_hint=label_hint)
            val = _to_python_val(df.loc[idx, data_col]) if idx is not None else None
            row_vals[i] = val
            if idx is not None:
                consumed.add(idx)
                if i not in row_labels:
                    raw = str(df.loc[idx, "label"] or "")
                    row_labels[i] = unicodedata.normalize("NFKC", raw)

        # Post-processing (BEFORE overflow): Investment Proceeds — sum all relevant rows
        inv_proc_val, inv_proc_indices = _sum_matching_rows(df, data_col, _INV_PROCEEDS_PATTERNS, consumed)
        if inv_proc_val is not None:
            row_vals[_CF_INV_PROCEEDS_IDX] = inv_proc_val
            consumed.update(inv_proc_indices)

        # Post-processing (BEFORE overflow): Debt Proceeds — sum LT + ST + credit lines
        debt_proc_val, debt_proc_indices = _sum_matching_rows(df, data_col, _DEBT_PROCEEDS_PATTERNS, consumed)
        if debt_proc_val is not None:
            row_vals[_CF_DEBT_PROCEEDS_IDX] = debt_proc_val
            consumed.update(debt_proc_indices)

        # Post-processing (BEFORE overflow): Debt Repayments — sum LT + ST + credit lines
        debt_rep_val, debt_rep_indices = _sum_matching_rows(df, data_col, _DEBT_REPAYMENTS_PATTERNS, consumed)
        if debt_rep_val is not None:
            row_vals[_CF_DEBT_REPAYMENTS_IDX] = debt_rep_val
            consumed.update(debt_rep_indices)

        # Collect raw overflow for all filings (incl. YTD); YTD subtraction applied after loop
        df_c = df[_consolidated_mask(df)]
        remaining = df_c[~df_c.index.isin(consumed)]
        filing_ov: dict[str, tuple[str, bool, Any]] = {}
        for _, ov_row in remaining.iterrows():
            ov_key = str(ov_row.get("concept", "") or "")
            if not ov_key or ov_key == "nan":
                continue
            ov_raw = str(ov_row.get("label", "") or "")
            ov_display = unicodedata.normalize("NFKC", ov_raw)
            ov_val = _to_python_val(ov_row.get(data_col))
            filing_ov[ov_key] = (ov_display, _is_nongaap_label(ov_display), ov_val)
        overflow_per_filing[label] = filing_ov

        collected[label] = (str(filing.filing_date), row_vals, is_ytd)
        ytd_raw[label] = row_vals  # Q1 standalone doubles as Q1 YTD base

    empty_ng = StatementTable(
        sheet_name="Data_CF_NG", quarter_labels=[], filing_dates=[],
        concepts=[], values=[], labels=[],
    )
    if not collected:
        return StatementTable(
            sheet_name="Data_CF",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in CF_TEMPLATE],
            values=[[] for _ in CF_TEMPLATE],
            labels=["" for _ in CF_TEMPLATE],
        ), empty_ng

    sorted_labels = sorted(collected.keys())

    # Convert YTD periods to standalone quarters via subtraction
    standalone: dict[str, dict[int, Any]] = {}
    for label in sorted_labels:
        _, row_vals, is_ytd = collected[label]
        if not is_ytd:
            standalone[label] = row_vals
        else:
            prev_label = _prev_quarter_label(label)
            if prev_label and prev_label in ytd_raw:
                prev = ytd_raw[prev_label]
                standalone[label] = {
                    i: (row_vals.get(i) - prev.get(i)
                        if row_vals.get(i) is not None and prev.get(i) is not None
                        else None)
                    for i in range(len(CF_TEMPLATE))
                }
            else:
                standalone[label] = row_vals  # no prior YTD — keep cumulative as best-effort

    # ── Convert YTD overflow to standalone (mirrors template subtraction above) ──
    # Union of all concepts seen in any filing's overflow
    all_overflow_keys: dict[str, tuple[str, bool]] = {}  # concept_key → (label, is_nongaap)
    for filing_ov in overflow_per_filing.values():
        for ov_key, (ov_lbl, ov_is_ng, _) in filing_ov.items():
            if ov_key not in all_overflow_keys:
                all_overflow_keys[ov_key] = (ov_lbl, ov_is_ng)

    for ov_key, (display_label, is_ng) in all_overflow_keys.items():
        out = ng_overflow if is_ng else gaap_overflow
        if ov_key not in out:
            out[ov_key] = {"label": display_label, "periods": {}}
        for q_lbl in sorted_labels:
            _, _, lbl_is_ytd = collected[q_lbl]
            filing_ov = overflow_per_filing.get(q_lbl, {})
            raw_val = filing_ov[ov_key][2] if ov_key in filing_ov else None
            if not lbl_is_ytd:
                if raw_val is not None:
                    out[ov_key]["periods"][q_lbl] = raw_val
            else:
                prev_lbl = _prev_quarter_label(q_lbl)
                prev_ov = overflow_per_filing.get(prev_lbl, {})
                prev_val = prev_ov[ov_key][2] if ov_key in prev_ov else None
                if raw_val is not None and prev_val is not None:
                    out[ov_key]["periods"][q_lbl] = raw_val - prev_val

    filing_dates = [collected[lbl][0] for lbl in sorted_labels]
    values: list[list[Any]] = [
        [standalone[lbl].get(i) for lbl in sorted_labels]
        for i in range(len(CF_TEMPLATE))
    ]

    tbl = StatementTable(
        sheet_name="Data_CF",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=[row[0] for row in CF_TEMPLATE],
        values=values,
        labels=[row_labels.get(i, "") for i in range(len(CF_TEMPLATE))],
    )

    # FCF = OCF − |Capex| (abs normalises sign: some companies report Capex negative)
    for j in range(len(sorted_labels)):
        op_cf = tbl.values[_CF_OP_CASH_IDX][j]
        capex = tbl.values[_CF_CAPEX_IDX][j]
        if op_cf is not None and capex is not None:
            tbl.values[_CF_FCF_IDX][j] = op_cf - abs(capex)

    # ── Build GAAP table: template rows (from tbl) + GAAP overflow ──────────
    concepts_g = tbl.concepts[:]
    labels_g   = tbl.labels[:]
    values_g   = [row[:] for row in tbl.values]
    for key in sorted(gaap_overflow):
        entry = gaap_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_g.append(entry["label"] or key)
        labels_g.append(key)
        values_g.append(row)

    gaap_tbl = StatementTable(
        sheet_name="Data_CF",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_g,
        labels=labels_g,
        values=values_g,
    )

    # ── Build NG table: Non-GAAP overflow rows only (no template rows) ───────
    concepts_n: list[str]       = []
    labels_n:   list[str]       = []
    values_n:   list[list[Any]] = []
    for key in sorted(ng_overflow):
        entry = ng_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_n.append(entry["label"] or key)
        labels_n.append(key)
        values_n.append(row)

    ng_tbl = StatementTable(
        sheet_name="Data_CF_NG",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_n,
        labels=labels_n,
        values=values_n,
    )
    return gaap_tbl, ng_tbl


# ── Three-statement merge ───────────────────────────────────────────────────

_STD_QUARTER_RE = re.compile(r"FY(\d{4})Q([1-4])$")


# ── 期間換算 ────────────────────────────────────────────────────────────────

def _fiscal_period_end(label: str, fy_end_month: int) -> tuple[int, int] | None:
    """財季標籤 → (西元年, 月)，指該財季**結束**的年月。無法解析回 None。

    財年 Y 結束於西元 Y 年的 fy_end_month 月（SEC 慣例）。第 q 財季結束於
    財年結束前 (4-q) 季，也就是 fy_end_month - 3*(4-q) 個月。
    """
    m = _STD_QUARTER_RE.match((label or "").strip())
    if m is None:
        return None
    year, q = int(m.group(1)), int(m.group(2))
    month = fy_end_month - 3 * (4 - q)
    while month <= 0:
        month += 12
        year -= 1
    return year, month


def _calendar_quarter(label: str, fy_end_month: int, period_end: str = "") -> str:
    """財季標籤 → 日曆季標籤（`2026Q1`）。年度標籤或無法解析回空字串。

    有真實期末日就直接用它換算（最準）；沒有才靠結算月反推。
    """
    if period_end and re.match(r"\d{4}-\d{2}", period_end):
        year, month = int(period_end[:4]), int(period_end[5:7])
        return f"{year}Q{(month - 1) // 3 + 1}"
    parsed = _fiscal_period_end(label, fy_end_month)
    if parsed is None:
        return ""
    year, month = parsed
    return f"{year}Q{(month - 1) // 3 + 1}"


def _fiscal_quarter(label: str) -> str:
    """`FY2026Q1` → `FY2026FQ1`。

    財季用 `FQ` 標記、財年用 `FY`，與第 4 列的日曆季（`2026Q1`）在視覺上就分得開。
    這兩列最容易被搞混——非 12 月結算的公司同一欄可能是 FY2026FQ1 但日曆 2025Q4，
    看錯就是整整一季的誤差，所以刻意讓兩種寫法長得不一樣。
    """
    m = _STD_QUARTER_RE.match((label or "").strip())
    return f"FY{m.group(1)}FQ{m.group(2)}" if m else ""


def _period_end(label: str, fy_end_month: int) -> str:
    """財季標籤 → 期末年月（`2026-03`）。無法解析回空字串。"""
    parsed = _fiscal_period_end(label, fy_end_month)
    if parsed is None:
        return ""
    year, month = parsed
    return f"{year}-{month:02d}"


# overflow 區的分隔標題。放公司特有、不在固定模板裡的 XBRL 科目——
# 實測每家中位數 IS 4 / BS 2 / CF 10 個，合計約 16 列。
OVERFLOW_SECTION = "Other (as reported)"

# 三表之間空幾列。捲動時要一眼看出換表了，光靠標題底色不夠。
SECTION_GAP = 5


def _merge_financials(is_tbl: StatementTable,
                       bs_tbl: StatementTable,
                       cf_tbl: StatementTable,
                       sheet_name: str = "Data_Financials(Q)",
                       fy_end_month: int = 12) -> StatementTable:
    """Merge IS + BS + CF into a single StatementTable.

    Quarter union is taken across all three statements; missing values are None.
    Section header rows ("Income Statement", "Balance Sheet", "Cash Flow")
    are inserted as separator rows with all-None values and empty labels.
    """
    all_qs = sorted(
        set(is_tbl.quarter_labels)
        | set(bs_tbl.quarter_labels)
        | set(cf_tbl.quarter_labels)
    )

    # Build date map (IS takes priority over BS over CF)
    date_map: dict[str, str] = {}
    for tbl in [cf_tbl, bs_tbl, is_tbl]:
        for lbl, dt in zip(tbl.quarter_labels, tbl.filing_dates):
            date_map[lbl] = dt
    filing_dates = [date_map.get(q, "") for q in all_qs]

    # 期末日同樣三表取聯集（IS 優先）。缺的留空，Data_Std 會退回反推年月。
    end_map: dict[str, str] = {}
    for tbl in [cf_tbl, bs_tbl, is_tbl]:
        for lbl, end in zip(tbl.quarter_labels, tbl.period_ends or []):
            if end:
                end_map[lbl] = end
    period_ends = [end_map.get(q, "") for q in all_qs]

    concepts:    list[str]        = []
    labels_col:  list[str]        = []
    values:      list[list[Any]]  = []

    def _add_header(title: str) -> None:
        concepts.append(title)
        labels_col.append("")
        values.append([None] * len(all_qs))

    def _add_blank() -> None:
        concepts.append("")
        labels_col.append("")
        values.append([None] * len(all_qs))

    # overflow（公司特有、不在模板裡的 XBRL 科目）全部集中到 sheet 最底部。
    #
    # 原本 overflow 接在每個 section 的模板列之後，IS 多幾行 overflow，BS 整段
    # 就往下推——實測 11 個輸出檔裡 `Cash` 落在第 28~56 列之間，跨檔案公式因此
    # 完全寫不出來。移到底部後模板列號跨公司固定。
    overflow: list[tuple[str, str, list[Any]]] = []

    def _add_rows(tbl: StatementTable, template_names: set[str]) -> None:
        q_idx = {q: j for j, q in enumerate(tbl.quarter_labels)}
        for i, concept in enumerate(tbl.concepts):
            label = tbl.labels[i] if tbl.labels else ""
            row = [_to_python_val(tbl.values[i][q_idx[q]])
                   if q in q_idx else None
                   for q in all_qs]
            if concept in template_names:
                concepts.append(concept)
                labels_col.append(label)
                values.append(row)
            else:
                overflow.append((concept, label, row))

    # 期間標籤三列放最上面。財季用 FY/FQ、日曆季用純數字，兩者視覺上分得開——
    # 非 12 月結算的公司同一欄可能是 FY2026FQ1 但日曆 2025Q4，看錯就是整整一季。
    def _add_label_row(name: str, vals: list[Any]) -> None:
        concepts.append(name)
        labels_col.append("")
        values.append(vals)

    _add_label_row("財季 Fiscal Quarter", [_fiscal_quarter(q) for q in all_qs])
    _add_label_row("日曆季 Calendar Quarter",
                   [_calendar_quarter(q, fy_end_month, period_ends[i] if i < len(period_ends) else "")
                    for i, q in enumerate(all_qs)])
    _add_label_row("期末結算日 Period End",
                   [(period_ends[i] if i < len(period_ends) and period_ends[i]
                     else _period_end(q, fy_end_month)) for i, q in enumerate(all_qs)])
    _add_blank()

    _add_header("Income Statement")
    _add_rows(is_tbl, {r[0] for r in IS_TEMPLATE})
    for _ in range(SECTION_GAP):
        _add_blank()
    _add_header("Balance Sheet")
    _add_rows(bs_tbl, {r[0] for r in BS_TEMPLATE})
    for _ in range(SECTION_GAP):
        _add_blank()
    _add_header("Cash Flow")
    _add_rows(cf_tbl, {r[0] for r in CF_TEMPLATE})

    if overflow:
        for _ in range(SECTION_GAP):
            _add_blank()
        _add_header(OVERFLOW_SECTION)
        for concept, label, row in overflow:
            concepts.append(concept)
            labels_col.append(label)
            values.append(row)

    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=all_qs,
        filing_dates=filing_dates,
        period_ends=period_ends,
        concepts=concepts,
        values=values,
        labels=labels_col,
    )


# ── BS/CF: dynamic row-union fetch (kept for reference / fallback) ──────────

def _build_dynamic_table(filings, stmt_method: str, sheet_name: str,
                          max_filings: int,
                          fy_end_month: int = 12) -> StatementTable | None:
    """Build BS or CF StatementTable using a dynamic row union across all filings."""
    concept_labels: dict[str, str] = {}
    periods: dict[str, tuple[str, dict[str, Any]]] = {}

    for filing in filings:
        if len(periods) >= max_filings:
            break
        try:
            stmt = getattr(filing.obj().financials, stmt_method)()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
        except Exception as exc:
            print(f"[fetcher_gaap] {sheet_name} warning: {exc!r}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        label = _col_to_quarter_label(q_col, fy_end_month)
        if label in periods:
            continue

        mask = _consolidated_mask(df)
        df_c = df[mask].reset_index(drop=True)

        period_vals: dict[str, Any] = {}
        for _, row in df_c.iterrows():
            key = _row_key(row)
            if key not in concept_labels:
                raw = str(row.get("label", "") or key)
                concept_labels[key] = unicodedata.normalize("NFKC", raw)
            period_vals[key] = _to_python_val(row.get(q_col))

        periods[label] = (str(filing.filing_date), period_vals)

    if not periods or not concept_labels:
        return None

    sorted_labels = sorted(periods.keys())
    filing_dates = [periods[lbl][0] for lbl in sorted_labels]
    concepts_ordered = list(concept_labels.keys())

    values: list[list[Any]] = []
    for key in concepts_ordered:
        values.append([periods[lbl][1].get(key) for lbl in sorted_labels])

    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=[concept_labels[k] for k in concepts_ordered],
        values=values,
    )


# ── Segment breakdown sheets ────────────────────────────────────────────────

def _dimension_axis(row) -> str:
    """從 XBRL 列取維度軸名稱，例如 `us-gaap:StatementBusinessSegmentsAxis`。

    edgartools 有 `dimension_axis` 欄；沒有時退回從 `dimension_label`
    （格式為 `"軸: 成員"`）取冒號前那段。都取不到回空字串。
    """
    axis = str(row.get("dimension_axis", "") or "").strip()
    if axis and axis != "nan":
        return axis
    label = str(row.get("dimension_label", "") or "")
    return label.split(":", 2)[0].strip() + ":" + label.split(":", 2)[1].strip()         if label.count(":") >= 2 else (label.split(":")[0].strip() if ":" in label else "")


def _build_segment_tables(filings, max_filings: int, fy_end_month: int = 12) -> list[StatementTable]:
    """Build one StatementTable per IS concept that has segment/dimension rows."""
    seg_data: dict[str, dict] = {}
    periods_seen: set[str] = set()

    for filing in filings:
        if len(periods_seen) >= max_filings:
            break
        _fd = getattr(filing, "filing_date", None)
        if isinstance(_fd, _date) and _fd < _XBRL_CUTOFF:
            break   # filings are newest-first; everything older is also pre-XBRL
        try:
            stmt = filing.obj().financials.income_statement()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
        except Exception as exc:
            print(f"[fetcher_gaap] Seg warning: {exc!r}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        period_label = _col_to_quarter_label(q_col, fy_end_month)
        filing_date = str(filing.filing_date)

        dim_col = df.get("dimension_member_label")
        if dim_col is None:
            continue
        mask_dim = ~(dim_col.isna() | (dim_col.astype(str) == "nan"))
        mask_not_abstract = ~df.get("abstract", False).astype(bool)
        df_dim = df[mask_dim & mask_not_abstract]

        for _, row in df_dim.iterrows():
            concept_xbrl = str(row.get("concept", "") or "")
            if not concept_xbrl:
                continue
            std = str(row.get("standard_concept", "") or "nan")
            member = str(row.get("dimension_member_label", "") or "")

            if concept_xbrl not in seg_data:
                seg_data[concept_xbrl] = {"std": std, "members": {}, "axes": {}, "periods": {}}

            seg_data[concept_xbrl]["members"].setdefault(member, member)
            # 記下這個 member 屬於哪個維度軸。沒有軸就分不出「業務別營收」與
            # 「權益項目別」——MSFT 實測會混進 `Retained earnings`、`Service Life`
            # 這種根本不是 segment 的東西。
            seg_data[concept_xbrl]["axes"].setdefault(member, _dimension_axis(row))

            if period_label not in seg_data[concept_xbrl]["periods"]:
                seg_data[concept_xbrl]["periods"][period_label] = (filing_date, {})

            seg_data[concept_xbrl]["periods"][period_label][1][member] = _to_python_val(row.get(q_col))

        periods_seen.add(period_label)

    tables: list[StatementTable] = []
    used_sheet_names: set[str] = set()

    for concept_xbrl, data in seg_data.items():
        if not data["periods"]:
            continue

        suffix = _seg_sheet_suffix(concept_xbrl, data["std"] if data["std"] != "nan" else None)
        sheet_name = f"Data_Seg_{suffix}"
        base = sheet_name
        n = 2
        while sheet_name in used_sheet_names:
            sheet_name = f"{base[:28]}_{n}"
            n += 1
        used_sheet_names.add(sheet_name)

        sorted_periods = sorted(data["periods"].keys())
        members_ordered = list(data["members"].keys())

        tables.append(StatementTable(
            sheet_name=sheet_name,
            quarter_labels=sorted_periods,
            filing_dates=[data["periods"][lbl][0] for lbl in sorted_periods],
            concepts=members_ordered,
            values=[[data["periods"][lbl][1].get(m) for lbl in sorted_periods]
                    for m in members_ordered],
            # B 欄借放維度軸，讓 segments.py 組長格式時能標出每一列屬於哪個軸
            labels=[data["axes"].get(m, "") for m in members_ordered],
        ))

    return tables


# ── Meta sheet ─────────────────────────────────────────────────────────────

# ── 期末流通股數（走封面頁 dei fact，不在三表裡）──────────────────────────
#
# 實測 ARLO / AAPL / NVDA / MSFT / COHR 五家都**沒有**在資產負債表 tag
# `us-gaap:CommonStockSharesOutstanding`——股數只寫在 `CommonStockValue` 的
# label 文字裡（"shares issued and outstanding: 108,745,373 at March 29, 2026"）。
#
# 真正拿得到的是封面頁的 `dei:EntityCommonStockSharesOutstanding`，走
# `Company.get_facts()`，ARLO 有 32 筆、AAPL 70 筆，2009 年起逐季都有。
#
# ⚠ 這個 fact 的日期是封面頁的「最近可行日期」，比財季結束**晚幾週**
#   （ARLO FY2025Q1 財季結束 2025-03-30，股數是 2025-05-02 的 103,400,957）。
#   它是公開資料裡最接近的時點股數，但不是財季結束當天的數字。用它算 BVPS
#   等於「期末權益 ÷ 幾週後的股數」，量級無虞但不是同一天。

_SHARES_CONCEPT = "dei:EntityCommonStockSharesOutstanding"


def _shares_label(fiscal_year, fiscal_period) -> str | None:
    """fact 的 (fiscal_year, fiscal_period) → 本專案的期間標籤。無法對映回 None。"""
    if fiscal_year is None or not fiscal_period:
        return None
    period = str(fiscal_period).strip().upper()
    try:
        year = int(fiscal_year)
    except (TypeError, ValueError):
        return None
    if period == "FY":
        return f"FY{year}"
    if re.fullmatch(r"Q[1-4]", period):
        return f"FY{year}{period}"
    return None


def _shares_map_from_records(records) -> dict[str, float]:
    """fact 記錄 → {期間標籤: 股數}。重複申報取最後一筆（更正後的）。"""
    out: dict[str, float] = {}
    for rec in records or []:
        label = _shares_label(rec.get("fiscal_year"), rec.get("fiscal_period"))
        if label is None:
            continue
        value = rec.get("numeric_value")
        if value is None:
            continue
        try:
            out[label] = float(value)
        except (TypeError, ValueError):
            continue
    return out


def _fetch_shares_outstanding(company) -> dict[str, float]:
    """抓封面頁流通股數的完整歷史序列。任何失敗回空 dict（該列留白即可）。"""
    try:
        facts = company.get_facts()
        df = facts.query().by_concept(_SHARES_CONCEPT).to_dataframe()
        if df is None or df.empty:
            return {}
        cols = [c for c in ("fiscal_year", "fiscal_period", "numeric_value")
                if c in df.columns]
        if len(cols) < 3:
            return {}
        return _shares_map_from_records(df[cols].to_dict("records"))
    except Exception as exc:
        print(f"[fetcher_gaap] 流通股數取得失敗: {type(exc).__name__}", file=sys.stderr)
        return {}


def _apply_shares_outstanding(tables: list[StatementTable],
                              shares_map: dict[str, float]) -> None:
    """把股數填進各表的 `Shares Outstanding` 列（就地修改）。

    沒有那一列的表直接跳過；對映不到的季度留 None，不用鄰近季度頂替。
    """
    if not shares_map:
        return
    for tbl in tables:
        if "Shares Outstanding" not in tbl.concepts:
            continue
        idx = tbl.concepts.index("Shares Outstanding")
        tbl.values[idx] = [shares_map.get(lbl) for lbl in tbl.quarter_labels]


def _build_meta_table(ticker: str, company_name: str,
                       tables: list[StatementTable],
                       fy_end_month: int = 12) -> StatementTable:
    """Build Data_Meta sheet with filing summary info."""
    n_quarters = 0
    quarter_labels: list[str] = []
    filing_dates: list[str] = []
    for tbl in tables:
        if tbl.sheet_name != "Data_Meta" and tbl.quarter_labels:
            n_quarters    = len(tbl.quarter_labels)
            quarter_labels = tbl.quarter_labels
            filing_dates   = tbl.filing_dates
            break

    # 品質檢查：9 個關鍵科目，各檢查「最近 4 期是否全部為空」——全空才算缺。
    # 只要有任一期有值就通過，所以 9/9 是「都至少抓到一期」，不代表每期都完整。
    score, missing_txt = "", ""
    q_tbl = next((t for t in tables if t.sheet_name == "Data_Financials(Q)"), None)
    if q_tbl is not None:
        missing = sorted(set(
            check_key_rows(q_tbl.concepts, q_tbl.values, "IS")
            + check_key_rows(q_tbl.concepts, q_tbl.values, "BS")
            + check_key_rows(q_tbl.concepts, q_tbl.values, "CF")
        ))
        from excel_formatter import ALL_KEY_ROWS as _ALL_KEY_ROWS
        total = len(_ALL_KEY_ROWS)
        score = f"{total - len(missing)}/{total}"
        missing_txt = "、".join(missing) if missing else "無"

    # 最新期間：這份檔案的資料抓到哪一季、那一季實際結束在哪天。
    latest_label, latest_end = "", ""
    if q_tbl is not None and q_tbl.quarter_labels:
        latest_label = q_tbl.quarter_labels[-1]
        ends = q_tbl.period_ends or []
        latest_end = ends[-1] if ends and ends[-1] else _period_end(latest_label, fy_end_month)

    # 財年起訖：結算月的下個月為起月。AAPL 9 月結算 → 財年 10 月起。
    start_month = fy_end_month % 12 + 1
    fy_span = f"{start_month} 月 – {fy_end_month} 月"

    return StatementTable(
        sheet_name="Data_Meta",
        quarter_labels=quarter_labels,
        filing_dates=filing_dates,
        # Fiscal Year End Month 是換算日曆季的依據——沒有它就無法把不同結算月
        # 公司的 FY 標籤對齊到同一個日曆季，是這張表唯一「程式在用」的欄位。
        # 品質檢查（原本在已移除的 Index sheet）也併到這裡。
        concepts=["Ticker", "Company Name", "Fetched Date", "Quarters Available",
                  "Fiscal Year End Month", "財年起訖", "最新期間", "最新期末日",
                  "Key Rows 完整度", "缺漏的 Key Rows"],
        values=[
            [ticker]            * n_quarters,
            [company_name]      * n_quarters,
            [str(date.today())] * n_quarters,
            [str(n_quarters)]   * n_quarters,
            [str(fy_end_month)] * n_quarters,
            [fy_span]           * n_quarters,
            [latest_label]      * n_quarters,
            [latest_end]        * n_quarters,
            [score]             * n_quarters,
            [missing_txt]       * n_quarters,
        ],
    )


# ── Public API ─────────────────────────────────────────────────────────────

def fetch_gaap_statements(ticker: str, identity: str,
                           max_filings: int = 80,
                           max_annual_filings: int = 20,
                           ai_config: dict | None = None,
                           start_year: int | None = None,
                           end_year: int | None = None,
                           fetch_quarterly: bool = True,
                           fetch_annual: bool = True,
                           excluded_sheets: set[str] | None = None) -> list[StatementTable]:
    """Fetch quarterly and/or annual GAAP statements for a ticker.

    Args:
        ticker:              Stock ticker, e.g. "AAPL"
        identity:            SEC EDGAR identity string
        max_filings:         Max 10-Q filings to process (default 80, ~20 years)
        max_annual_filings:  Max 10-K filings to process (default 20, ~20 years)
        ai_config:           AI config dict (provider/model/api_key) for E2 diagnosis
        start_year:          Only include filings from this year onwards (None = no limit)
        end_year:            Only include filings up to this year (None = no limit)
        fetch_quarterly:     Whether to fetch 10-Q data (default True)
        fetch_annual:        Whether to fetch 10-K data (default True)
        excluded_sheets:     Set of sheet names to skip in the output

    Returns:
        List of StatementTable

    Raises:
        ValueError: No filings found for the requested form type(s)
    """
    ai_config = ai_config or {}
    excluded_sheets = excluded_sheets or set()
    set_identity(identity)
    company = Company(ticker)

    filings_q = list(company.get_filings(form="10-Q", amendments=False)) if fetch_quarterly else []
    filings_k = list(company.get_filings(form="10-K", amendments=False)) if fetch_annual else []

    if fetch_quarterly and not filings_q:
        raise ValueError(
            f"No 10-Q filings found for ticker '{ticker}'. "
            "The ticker may be invalid or the company may not file 10-Qs."
        )
    if not fetch_quarterly and not filings_k:
        raise ValueError(
            f"No 10-K filings found for ticker '{ticker}'. "
            "The ticker may be invalid or the company may not file 10-Ks."
        )

    # Apply year range filter
    filings_q = _filter_filings_by_year(filings_q, start_year, end_year)
    filings_k = _filter_filings_by_year(filings_k, start_year, end_year)

    overrides = load_overrides(ticker)
    if filings_k:
        fy_end_month = _detect_fy_end_month(filings_k)
    elif fetch_quarterly and filings_q:
        _probe_k = list(company.get_filings(form="10-K", amendments=False))[:1]
        fy_end_month = _detect_fy_end_month(_probe_k) if _probe_k else 12
    else:
        fy_end_month = 12

    tables: list[StatementTable] = []

    if fetch_quarterly and filings_q:
        is_tbl, is_ng = _build_is_table(filings_q, max_filings, is_overrides=overrides.get("IS", {}), fy_end_month=fy_end_month)
        bs_tbl, bs_ng = _build_bs_table(filings_q, max_filings, bs_overrides=overrides.get("BS", {}), fy_end_month=fy_end_month)
        cf_tbl, cf_ng = _build_cf_table(filings_q, max_filings, cf_overrides=overrides.get("CF", {}), fy_end_month=fy_end_month)

        # Diagnose key rows that are all-None in recent quarters
        missing_is = check_key_rows(is_tbl.concepts, is_tbl.values, "IS")
        missing_bs = check_key_rows(bs_tbl.concepts, bs_tbl.values, "BS")
        missing_cf = check_key_rows(cf_tbl.concepts, cf_tbl.values, "CF")

        if missing_is or missing_bs or missing_cf:
            try:
                tenq_latest = filings_q[0].obj()
                latest_is_df = tenq_latest.financials.income_statement().to_dataframe()
                latest_bs_df = tenq_latest.financials.balance_sheet().to_dataframe()
                latest_cf_df = tenq_latest.financials.cashflow_statement().to_dataframe()
            except Exception as exc:
                print(f"[{ticker}] 診斷：無法取得最新 filing DataFrame — {exc!r}", file=sys.stderr)
                latest_is_df = latest_bs_df = latest_cf_df = None

            new_overrides: dict[str, dict] = {}
            if missing_is and latest_is_df is not None:
                fixes = run_diagnosis(ticker, "IS", latest_is_df, missing_is, ai_config)
                if fixes:
                    new_overrides["IS"] = fixes
            if missing_bs and latest_bs_df is not None:
                fixes = run_diagnosis(ticker, "BS", latest_bs_df, missing_bs, ai_config)
                if fixes:
                    new_overrides["BS"] = fixes
            if missing_cf and latest_cf_df is not None:
                fixes = run_diagnosis(ticker, "CF", latest_cf_df, missing_cf, ai_config)
                if fixes:
                    new_overrides["CF"] = fixes

            if new_overrides:
                total_fixes = sum(len(v) for v in new_overrides.values())
                print(f"[{ticker}] 自動修復：找到 {total_fixes} 項缺失指標修復方案，重新建表。", file=sys.stderr)
                overrides = load_overrides(ticker)
                is_tbl, is_ng = _build_is_table(filings_q, max_filings, is_overrides=overrides.get("IS", {}), fy_end_month=fy_end_month)
                bs_tbl, bs_ng = _build_bs_table(filings_q, max_filings, bs_overrides=overrides.get("BS", {}), fy_end_month=fy_end_month)
                cf_tbl, cf_ng = _build_cf_table(filings_q, max_filings, cf_overrides=overrides.get("CF", {}), fy_end_month=fy_end_month)
            else:
                remaining = missing_is + missing_bs + missing_cf
                if remaining:
                    no_key = "" if ai_config.get("api_key") else "（未設 AI API key，E2 診斷已跳過）"
                    print(f"[{ticker}] 警告：{remaining} 在 EDGAR 中無對應概念{no_key}。", file=sys.stderr)

        quarterly_tbl = _merge_financials(is_tbl, bs_tbl, cf_tbl, sheet_name="Data_Financials(Q)", fy_end_month=fy_end_month)
        tables.append(quarterly_tbl)
        if any(tbl.concepts for tbl in [is_ng, bs_ng, cf_ng]):
            ng_q_tbl = _merge_financials(is_ng, bs_ng, cf_ng, sheet_name="Data_Financials_NG(Q)", fy_end_month=fy_end_month)
            tables.append(ng_q_tbl)

    if fetch_annual and filings_k:
        is_ann, is_ann_ng = _build_is_table(filings_k, max_annual_filings, is_overrides=overrides.get("IS", {}), fy_end_month=fy_end_month)
        bs_ann, bs_ann_ng = _build_bs_table(filings_k, max_annual_filings, bs_overrides=overrides.get("BS", {}), fy_end_month=fy_end_month)
        cf_ann, cf_ann_ng = _build_cf_table(filings_k, max_annual_filings, cf_overrides=overrides.get("CF", {}), fy_end_month=fy_end_month)
        annual_tbl = _merge_financials(is_ann, bs_ann, cf_ann, sheet_name="Data_Financials(Y)", fy_end_month=fy_end_month)
        tables.append(annual_tbl)
        if any(tbl.concepts for tbl in [is_ann_ng, bs_ann_ng, cf_ann_ng]):
            ng_y_tbl = _merge_financials(is_ann_ng, bs_ann_ng, cf_ann_ng, sheet_name="Data_Financials_NG(Y)", fy_end_month=fy_end_month)
            tables.append(ng_y_tbl)

    if fetch_quarterly and filings_q:
        seg_tables = _build_segment_tables(filings_q, max_filings, fy_end_month=fy_end_month)
        tables.extend(t for t in seg_tables if t.sheet_name not in excluded_sheets)

    company_name = getattr(company, "name", ticker) or ticker
    # 期末流通股數不在三表裡，走封面頁 dei fact 另外補（見 _fetch_shares_outstanding）
    _apply_shares_outstanding(tables, _fetch_shares_outstanding(company))

    tables.append(_build_meta_table(ticker, company_name, tables, fy_end_month))

    for tbl in tables:
        tbl.ticker = ticker
    return tables


def preview_sheets(ticker: str, identity: str) -> list[str]:
    """Quick scan: fetch only the latest 10-Q to detect segment sheet names.

    Returns the predicted list of sheet names without performing a full fetch.
    Takes ~5-15 seconds (one HTTP request for the latest filing).

    Returns:
        List of sheet name strings. Fixed sheets (Financials Q/Y, Meta) are
        always included. Data_Seg_* sheets are detected from the latest 10-Q.
    """
    fixed = ["Data_Financials(Q)", "Data_Financials(Y)", "Data_Meta"]

    set_identity(identity)
    company = Company(ticker)
    filings_q = list(company.get_filings(form="10-Q", amendments=False))
    if not filings_q:
        return fixed

    try:
        seg_tables = _build_segment_tables([filings_q[0]], max_filings=1)
        seg_names = [t.sheet_name for t in seg_tables]
    except Exception as exc:
        print(f"[preview_sheets] Segment scan failed: {exc!r}", file=sys.stderr)
        seg_names = []

    return fixed + seg_names
