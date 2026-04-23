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
from datetime import date
from typing import Any

import pandas as pd
from edgar import Company, set_identity as set_identity


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


# ── Constants ─────────────────────────────────────────────────────────────

META_COLS: set[str] = {
    'concept', 'label', 'standard_concept', 'level', 'abstract',
    'dimension', 'is_breakdown', 'dimension_axis', 'dimension_member',
    'dimension_member_label', 'dimension_label', 'unit', 'point_in_time',
    'balance', 'weight', 'preferred_sign', 'parent_concept', 'parent_abstract_concept',
}

# Tuple: (label, std_concept, fallback_suffix, source, match, label_hint)
#   label         — display name (Col A)
#   std_concept   — primary: standard_concept == value
#   fallback_suffix — secondary: concept contains value
#   source        — "IS" | "CF" | "BS" | "DERIVED"
#   match         — "first" | "last"  (which occurrence when multiple rows match)
#   label_hint    — tertiary filter: prefer rows whose label contains this string
_T = tuple[str, str | None, str, str, str, str | None]

IS_TEMPLATE: list[_T] = [
    ("Revenue",                    "Revenue",                        "RevenueFromContractWithCustomer",                        "IS", "first", None),
    ("Cost of Revenue",            "CostOfGoodsAndServicesSold",     "CostOfGoodsSold",                                       "IS", "first", "cost"),
    ("Gross Profit",               "GrossProfit",                    "GrossProfit",                                            "IS", "first", None),
    ("R&D Expense",                "ResearchAndDevelopmentExpenses", "ResearchAndDevelopment",                                 "IS", "first", None),
    ("SG&A Expense",               "SellingGeneralAndAdminExpenses", "SellingGeneralAndAdmin",                                 "IS", "first", None),
    ("D&A (CF memo)",              "DepreciationExpense",            "DepreciationDepletionAndAmortization",                   "CF", "first", None),
    ("Other Operating Expense",    "OtherOperatingExpenses",         "OtherOperatingExpense",                                  "IS", "first", None),
    ("Total Operating Expense",    "TotalOperatingExpenses",         "OperatingExpenses",                                      "IS", "first", None),
    ("Operating Income",           "OperatingIncomeLoss",            "OperatingIncomeLoss",                                    "IS", "first", None),
    ("Interest Expense",           "InterestExpense",                "InterestExpense",                                        "IS", "first", None),
    ("Interest Income",            "InterestIncome",                 "InterestIncome",                                         "IS", "first", None),
    ("Other Non-op Inc/(Exp)",     None,                             "OtherNonoperatingIncome",                                "IS", "first", None),
    ("Total Non-op Income/(Loss)", "NonoperatingIncomeExpense",      "NonoperatingIncome",                                     "IS", "first", None),
    ("Pre-tax Income",             "PretaxIncomeLoss",               "IncomeLossFromContinuingOperationsBeforeIncomeTax",       "IS", "first", None),
    ("Income Tax",                 "IncomeTaxes",                    "IncomeTaxExpense",                                       "IS", "first", None),
    ("Net Income",                 "NetIncome",                      "NetIncomeLoss",                                          "IS", "first", None),
    ("Minority Interest",          None,                             "NetIncomeLossAttributableToNoncontrollingInterest",       "IS", "first", None),
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
]

CF_TEMPLATE: list[_T] = [
    # ── Operating ────────────────────────────────────────────────────────
    ("Net Income",                 "NetIncome",                          "NetIncomeLoss|ProfitLoss",                              "CF", "first", None),
    ("D&A",                        "DepreciationExpense",                "DepreciationDepletionAndAmortization",                  "CF", "first", None),
    ("SBC",                        "StockBasedCompensationExpense",      "ShareBasedCompensation",                                "CF", "first", None),
    ("Amortization of Intangibles","AmortizationOfIntangibles",          "AmortizationOfIntangibleAssets",                        "CF", "first", None),
    ("Change in Receivables",      "ChangeInReceivables",                "IncreaseDecreaseInAccountsReceivable",                  "CF", "first", "receivable"),
    ("Change in Inventories",      None,                                 "IncreaseDecreaseInInventories",                         "CF", "first", "inventories"),
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
    ("Dividends Paid",             "DistributionsToMinorityInterests",   "PaymentsOfDividends",                                   "CF", "first", "dividend"),
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


# ── Helpers ────────────────────────────────────────────────────────────────

def _col_to_quarter_label(col_name: str) -> str:
    """Convert edgartools period column name to FY label.

    Examples:
        "2023-03-31 (Q1)"  -> "FY2023Q1"
        "2024-12-31 (FY)"  -> "FY2024"
        "2023-03-31"       -> "2023-03-31"
    """
    m = re.match(r"(\d{4})-\d{2}-\d{2}\s+\((\w+)\)", col_name.strip())
    if m:
        year, period = m.group(1), m.group(2)
        return f"FY{year}" if period.upper() == "FY" else f"FY{year}{period}"
    return col_name


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


def _build_template_table(filings, template: list[_T], sheet_name: str,
                           stmt_method: str, max_filings: int) -> StatementTable:
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

        label = _col_to_quarter_label(q_col)
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

def _build_is_table(filings, max_filings: int) -> StatementTable:
    """Build Data_IS StatementTable from 10-Q filings using the fixed IS template."""
    periods: dict[str, tuple[str, dict[int, Any]]] = {}
    row_labels: dict[int, str] = {}

    for filing in filings:
        if len(periods) >= max_filings:
            break
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

        label = _col_to_quarter_label(q_col)
        if label in periods:
            continue

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

        row_vals: dict[int, Any] = {}
        for i, (_, std_concept, fallback, source, match, label_hint) in enumerate(IS_TEMPLATE):
            if source == "CF":
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
                val = _to_python_val(df.loc[idx, q_col]) if idx is not None else None
                if idx is not None and i not in row_labels:
                    raw = str(df.loc[idx, "label"] or "")
                    row_labels[i] = unicodedata.normalize("NFKC", raw)
            row_vals[i] = val

        # ── Post-processing: fallbacks not expressible in the 6-tuple ──

        # 1. Total Non-op: DERIVED = Pre-tax − Operating Income
        if row_vals.get(_NONOP_TOTAL_IDX) is None:
            op_val     = row_vals.get(_OP_INCOME_IDX)
            pretax_val = row_vals.get(_PRETAX_IDX)
            if op_val is not None and pretax_val is not None:
                row_vals[_NONOP_TOTAL_IDX] = pretax_val - op_val

        # 2. Net Income: ProfitLoss fallback (BA, TSLA, XOM, WMT)
        if row_vals.get(_NET_INCOME_IDX) is None:
            idx = _match_is_row(df, "ProfitLoss", "ProfitLoss")
            if idx is not None:
                row_vals[_NET_INCOME_IDX] = _to_python_val(df.loc[idx, q_col])
                if _NET_INCOME_IDX not in row_labels:
                    row_labels[_NET_INCOME_IDX] = unicodedata.normalize(
                        "NFKC", str(df.loc[idx, "label"] or ""))

        # 3. D&A label fallback: for companies where standard_concept = nan (TSLA)
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

        periods[label] = (str(filing.filing_date), row_vals)

    if not periods:
        return StatementTable(
            sheet_name="Data_IS",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in IS_TEMPLATE],
            values=[[] for _ in IS_TEMPLATE],
            labels=["" for _ in IS_TEMPLATE],
        )

    sorted_labels = sorted(periods.keys())
    filing_dates  = [periods[lbl][0] for lbl in sorted_labels]

    values: list[list[Any]] = []
    for i in range(len(IS_TEMPLATE)):
        values.append([periods[lbl][1].get(i) for lbl in sorted_labels])

    return StatementTable(
        sheet_name="Data_IS",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=[row[0] for row in IS_TEMPLATE],
        values=values,
        labels=[row_labels.get(i, "") for i in range(len(IS_TEMPLATE))],
    )


# ── BS: template-based fetch ────────────────────────────────────────────────

def _build_bs_table(filings, max_filings: int) -> StatementTable:
    """Build Data_BS StatementTable using the fixed BS template.

    Balance sheet columns in edgartools are instant (bare date, e.g. "2024-03-31")
    rather than period ("2024-03-31 (Q1)"), so _current_q_col cannot find them.
    We derive the quarter label from the IS statement (same filing) for merge alignment.
    """
    periods: dict[str, tuple[str, dict[int, Any]]] = {}
    row_labels: dict[int, str] = {}

    for filing in filings:
        if len(periods) >= max_filings:
            break
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

        label = _col_to_quarter_label(is_q_col) if is_q_col else _col_to_quarter_label(bs_col)
        if label in periods:
            continue

        row_vals: dict[int, Any] = {}
        for i, (_, std_concept, fallback, source, match, label_hint) in enumerate(BS_TEMPLATE):
            if source == "DERIVED":
                row_vals[i] = None
                continue
            idx = _match_is_row(df, std_concept, fallback, match=match, label_hint=label_hint)
            val = _to_python_val(df.loc[idx, bs_col]) if idx is not None else None
            row_vals[i] = val
            if idx is not None and i not in row_labels:
                raw = str(df.loc[idx, "label"] or "")
                row_labels[i] = unicodedata.normalize("NFKC", raw)

        periods[label] = (str(filing.filing_date), row_vals)

    if not periods:
        return StatementTable(
            sheet_name="Data_BS",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in BS_TEMPLATE],
            values=[[] for _ in BS_TEMPLATE],
            labels=["" for _ in BS_TEMPLATE],
        )

    sorted_labels = sorted(periods.keys())
    filing_dates = [periods[lbl][0] for lbl in sorted_labels]
    values = [[periods[lbl][1].get(i) for lbl in sorted_labels] for i in range(len(BS_TEMPLATE))]

    return StatementTable(
        sheet_name="Data_BS",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=[row[0] for row in BS_TEMPLATE],
        values=values,
        labels=[row_labels.get(i, "") for i in range(len(BS_TEMPLATE))],
    )


# ── CF: template-based fetch ────────────────────────────────────────────────

def _build_cf_table(filings, max_filings: int) -> StatementTable:
    """Build Data_CF StatementTable using the fixed CF template.

    Q1 and FY filings have standalone period columns (Q1/FY) and are used directly.
    Q2 and Q3 filings have YTD (cumulative) CF columns; standalone quarter values are
    derived by subtracting the prior period's YTD: Q2 = Q2_YTD − Q1, Q3 = Q3_YTD − Q2_YTD.
    The IS statement of each filing is consulted for the quarter label (same approach as BS).
    """
    # collected: label → (filing_date, {row_i: raw_value}, is_ytd)
    collected: dict[str, tuple[str, dict[int, Any], bool]] = {}
    # ytd_raw: stores raw values (standalone for Q1/FY, cumulative for YTD)
    # used as the subtraction base for the next YTD period
    ytd_raw: dict[str, dict[int, Any]] = {}
    row_labels: dict[int, str] = {}

    for filing in filings:
        if len(collected) >= max_filings:
            break
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
            label = _col_to_quarter_label(q_col)
            if label in collected:
                continue
            is_ytd = False
            data_col = q_col
        else:
            ytd_col = _ytd_col(df)
            if ytd_col is None or is_q_col is None:
                continue
            label = _col_to_quarter_label(is_q_col)
            if label in collected:
                continue
            is_ytd = True
            data_col = ytd_col

        row_vals: dict[int, Any] = {}
        for i, (_, std_concept, fallback, source, match, label_hint) in enumerate(CF_TEMPLATE):
            if source == "DERIVED":
                row_vals[i] = None
                continue
            idx = _match_is_row(df, std_concept, fallback, match=match, label_hint=label_hint)
            val = _to_python_val(df.loc[idx, data_col]) if idx is not None else None
            row_vals[i] = val
            if idx is not None and i not in row_labels:
                raw = str(df.loc[idx, "label"] or "")
                row_labels[i] = unicodedata.normalize("NFKC", raw)

        collected[label] = (str(filing.filing_date), row_vals, is_ytd)
        ytd_raw[label] = row_vals  # Q1 standalone doubles as Q1 YTD base

    if not collected:
        return StatementTable(
            sheet_name="Data_CF",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in CF_TEMPLATE],
            values=[[] for _ in CF_TEMPLATE],
            labels=["" for _ in CF_TEMPLATE],
        )

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

    return tbl


# ── Three-statement merge ───────────────────────────────────────────────────

def _merge_financials(is_tbl: StatementTable,
                       bs_tbl: StatementTable,
                       cf_tbl: StatementTable,
                       sheet_name: str = "Data_Financials(Q)") -> StatementTable:
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

    def _add_rows(tbl: StatementTable) -> None:
        q_idx = {q: j for j, q in enumerate(tbl.quarter_labels)}
        for i, concept in enumerate(tbl.concepts):
            concepts.append(concept)
            labels_col.append(tbl.labels[i] if tbl.labels else "")
            row = [_to_python_val(tbl.values[i][q_idx[q]])
                   if q in q_idx else None
                   for q in all_qs]
            values.append(row)

    _add_header("Income Statement")
    _add_rows(is_tbl)
    _add_blank()
    _add_header("Balance Sheet")
    _add_rows(bs_tbl)
    _add_blank()
    _add_header("Cash Flow")
    _add_rows(cf_tbl)

    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=all_qs,
        filing_dates=filing_dates,
        concepts=concepts,
        values=values,
        labels=labels_col,
    )


# ── BS/CF: dynamic row-union fetch (kept for reference / fallback) ──────────

def _build_dynamic_table(filings, stmt_method: str, sheet_name: str,
                          max_filings: int) -> StatementTable | None:
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

        label = _col_to_quarter_label(q_col)
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

def _build_segment_tables(filings, max_filings: int) -> list[StatementTable]:
    """Build one StatementTable per IS concept that has segment/dimension rows."""
    seg_data: dict[str, dict] = {}
    periods_seen: set[str] = set()

    for filing in filings:
        if len(periods_seen) >= max_filings:
            break
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

        period_label = _col_to_quarter_label(q_col)
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
                seg_data[concept_xbrl] = {"std": std, "members": {}, "periods": {}}

            seg_data[concept_xbrl]["members"].setdefault(member, member)

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
        ))

    return tables


# ── Meta sheet ─────────────────────────────────────────────────────────────

def _build_meta_table(ticker: str, company_name: str,
                       tables: list[StatementTable]) -> StatementTable:
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

    return StatementTable(
        sheet_name="Data_Meta",
        quarter_labels=quarter_labels,
        filing_dates=filing_dates,
        concepts=["Ticker", "Company Name", "Fetched Date", "Quarters Available"],
        values=[
            [ticker]            * n_quarters,
            [company_name]      * n_quarters,
            [str(date.today())] * n_quarters,
            [str(n_quarters)]   * n_quarters,
        ],
    )


# ── Public API ─────────────────────────────────────────────────────────────

def fetch_gaap_statements(ticker: str, identity: str,
                           max_filings: int = 80,
                           max_annual_filings: int = 20) -> list[StatementTable]:
    """Fetch quarterly and annual GAAP statements for a ticker.

    Args:
        ticker:              Stock ticker, e.g. "AAPL"
        identity:            SEC EDGAR identity string
        max_filings:         Max 10-Q filings to process (default 80, ~20 years)
        max_annual_filings:  Max 10-K filings to process (default 20, ~20 years)

    Returns:
        List of StatementTable: Data_Financials(Q), Data_Financials(Y), Data_Seg_*, Data_Meta

    Raises:
        ValueError: No 10-Q filings found for ticker
    """
    set_identity(identity)
    company = Company(ticker)

    filings_q = list(company.get_filings(form="10-Q", amendments=False))
    if not filings_q:
        raise ValueError(
            f"No 10-Q filings found for ticker '{ticker}'. "
            "The ticker may be invalid or the company may not file 10-Qs."
        )

    is_tbl = _build_is_table(filings_q, max_filings)
    bs_tbl = _build_bs_table(filings_q, max_filings)
    cf_tbl = _build_cf_table(filings_q, max_filings)
    quarterly_tbl = _merge_financials(is_tbl, bs_tbl, cf_tbl, sheet_name="Data_Financials(Q)")

    tables: list[StatementTable] = [quarterly_tbl]

    filings_k = list(company.get_filings(form="10-K", amendments=False))
    if filings_k:
        is_ann = _build_is_table(filings_k, max_annual_filings)
        bs_ann = _build_bs_table(filings_k, max_annual_filings)
        cf_ann = _build_cf_table(filings_k, max_annual_filings)
        annual_tbl = _merge_financials(is_ann, bs_ann, cf_ann, sheet_name="Data_Financials(Y)")
        tables.append(annual_tbl)

    tables.extend(_build_segment_tables(filings_q, max_filings))

    company_name = getattr(company, "name", ticker) or ticker
    tables.append(_build_meta_table(ticker, company_name, tables))

    for tbl in tables:
        tbl.ticker = ticker
    return tables
