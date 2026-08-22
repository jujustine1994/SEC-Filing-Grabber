"""facts_mapping.py — 模板列 → SEC companyfacts 的 us-gaap concept 對照表。

**這張表不是手填的。** 模板的 `std_concept` 欄是 edgartools 正規化過的名字
（`NetIncome`、`ResearchAndDevelopmentExpenses`），不是原始 us-gaap element name
（`NetIncomeLoss`、`ResearchAndDevelopmentExpense`）。憑印象填 95 列一定會錯，
而且錯了不會有人發現——數字看起來都很像。

產生方式（2026-08-22，`scripts/spike_derive_mapping.py`）：拿現行「逐份解 filing」
路徑抓到的數字當**答案卷**，對 51 家公司 × companyfacts 裡的每一個 concept 算
「同一個期末日、數字對得上的比例」，命中率最高的就是正確對應。順便自動偵測
正負號是否相反。每一列後面的註解就是它的證據強度。

51 家刻意涵蓋大中小型 × 跨產業，包含金融股（JPM/GS/BAC/SCHW）與 REIT（PLD）
——它們的報表結構跟製造業差很多，是檢驗模板列通不通用最有效的一群。

## 欄位

    concepts   依序試，第一個有資料的就用（同一列在不同公司/不同年代可能
               對到不同 concept，例如早年 `Revenues`、後來
               `RevenueFromContractWithCustomerExcludingAssessedTax`）
    kind       "quarter"（期間值）/ "instant"（時點值）/ "annual"
    unit       省略＝USD。EPS 是 "USD/shares"、股數是 "shares"
    taxonomy   省略＝us-gaap。流通股數在 "dei"
    negate     省略＝False。現行路徑對某些列做過符號正規化（Capex 記成
               現金流出的負數），companyfacts 給的是公司原始 tag 的正號

## 三張表分開、列序照模板

CTH 2026-08-22 要求「原本模板的格式架構要維持，包含排序方式、三表分別」。
這裡三個 dict 的**鍵的順序就是模板的列序**，`fetcher_facts.build_table()`
直接照這個順序產列，不重新排序。改動時請維持與
`fetcher_gaap.IS_TEMPLATE` / `BS_TEMPLATE` / `CF_TEMPLATE` 相同的順序。
"""

from __future__ import annotations

IS_MAPPING: dict[str, dict] = {
    'Revenue': {"concepts": ['RevenueFromContractWithCustomerExcludingAssessedTax', 'Revenues', 'SalesRevenueNet'], "kind": 'quarter'},  # 36/51 家（覆蓋 71%、命中 0.97）
    'Cost of Revenue': {"concepts": ['CostOfGoodsAndServicesSold', 'CostOfRevenue'], "kind": 'quarter'},  # 26/41 家（覆蓋 63%、命中 0.99）
    'Gross Profit': {"concepts": ['GrossProfit'], "kind": 'quarter'},  # 28/41 家（覆蓋 68%、命中 1.00）
    'R&D Expense': {"concepts": ['ResearchAndDevelopmentExpense', 'ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost'], "kind": 'quarter'},  # 31/37 家（覆蓋 84%、命中 1.00）
    'SG&A Expense': {"concepts": ['SellingGeneralAndAdministrativeExpense', 'SellingAndMarketingExpense', 'GeneralAndAdministrativeExpense'], "kind": 'quarter'},  # 28/47 家（覆蓋 60%、命中 1.00）
    'D&A (CF memo)': {"concepts": ['DepreciationDepletionAndAmortization', 'DepreciationAndAmortization', 'Depreciation', 'DepreciationAmortizationAndAccretionNet'], "kind": 'quarter'},  # 24/51 家（覆蓋 47%、命中 0.97）
    # 'Other Operating Expense'：51 家裡現行路徑一家都沒抓到（答案卷家數 0），facts 這邊也沒有對應 concept。無證據可判，保留列但不對應
    'Total Operating Expense': {"concepts": ['OperatingExpenses'], "kind": 'quarter'},  # 23/25 家（覆蓋 92%、命中 1.00）
    'Total Costs and Expenses': {"concepts": ['CostsAndExpenses'], "kind": 'quarter'},  # 14/14 家（覆蓋 100%、命中 1.00）
    'Operating Income': {"concepts": ['OperatingIncomeLoss'], "kind": 'quarter'},  # 39/47 家（覆蓋 83%、命中 1.00）
    'Interest Expense': {"concepts": ['InterestExpense', 'InterestExpenseNonoperating', 'InterestExpenseOperating', 'InterestExpenseDebt'], "kind": 'quarter'},  # 33/45 家（覆蓋 73%、命中 0.98）
    'Interest Income': {"concepts": ['InterestIncomeExpenseNonoperatingNet', 'InterestIncomeExpenseNet'], "kind": 'quarter'},  # 5/15 家（覆蓋 33%、命中 1.00）
    'Other Non-op Inc/(Exp)': {"concepts": ['OtherNonoperatingIncomeExpense'], "kind": 'quarter'},  # 29/34 家（覆蓋 85%、命中 1.00）
    'Total Non-op Income/(Loss)': {"concepts": ['OtherNonoperatingIncomeExpense'], "kind": 'quarter'},  # 26/49 家（覆蓋 53%、命中 0.94）
    'Pre-tax Income': {"concepts": ['IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest', 'IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments'], "kind": 'quarter'},  # 41/48 家（覆蓋 85%、命中 1.00）
    'Income Tax': {"concepts": ['IncomeTaxExpenseBenefit'], "kind": 'quarter'},  # 50/51 家（覆蓋 98%、命中 0.99）
    'Net Income': {"concepts": ['NetIncomeLoss', 'NetIncomeLossAvailableToCommonStockholdersBasic', 'NetIncomeLossAttributableToParentDiluted'], "kind": 'quarter'},  # 47/51 家（覆蓋 92%、命中 0.99）
    'Minority Interest': {"concepts": ['NetIncomeLossAttributableToNoncontrollingInterest', 'ComprehensiveIncomeNetOfTaxAttributableToNoncontrollingInterest', 'IncomeLossFromContinuingOperationsAttributableToNoncontrollingEntity'], "kind": 'quarter'},  # 18/21 家（覆蓋 86%、命中 0.98）
    'Net Income incl. NCI': {"concepts": ['ProfitLoss'], "kind": 'quarter'},  # 20/21 家（覆蓋 95%、命中 0.99）
    'SBC': {"concepts": ['ShareBasedCompensation'], "kind": 'quarter'},  # 43/44 家（覆蓋 98%、命中 0.99）
    'Basic EPS': {"concepts": ['EarningsPerShareBasic'], "kind": 'quarter', "unit": 'USD/shares'},  # 51/51 家（覆蓋 100%、命中 0.98）
    'Diluted EPS': {"concepts": ['EarningsPerShareDiluted'], "kind": 'quarter', "unit": 'USD/shares'},  # 51/51 家（覆蓋 100%、命中 0.98）
    'Basic Shares': {"concepts": ['WeightedAverageNumberOfSharesOutstandingBasic'], "kind": 'quarter', "unit": 'shares'},  # 45/45 家（覆蓋 100%、命中 1.00）
    'Diluted Shares': {"concepts": ['WeightedAverageNumberOfDilutedSharesOutstanding'], "kind": 'quarter', "unit": 'shares'},  # 46/46 家（覆蓋 100%、命中 0.99）
}


BS_MAPPING: dict[str, dict] = {
    'Cash': {"concepts": ['CashAndCashEquivalentsAtCarryingValue', 'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents', 'CashAndCashEquivalentsFairValueDisclosure', 'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsIncludingDisposalGroupAndDiscontinuedOperations'], "kind": 'instant'},  # 45/45 家（覆蓋 100%、命中 1.00）
    'Short-term Investments': {"concepts": ['AvailableForSaleSecuritiesDebtSecuritiesCurrent', 'MarketableSecuritiesCurrent', 'ShortTermInvestments', 'AvailableForSaleSecuritiesDebtSecurities'], "kind": 'instant'},  # 14/36 家（覆蓋 39%、命中 0.99）
    'Accounts Receivable': {"concepts": ['AccountsReceivableNetCurrent', 'ReceivablesNetCurrent'], "kind": 'instant'},  # 31/36 家（覆蓋 86%、命中 1.00）
    'Inventories': {"concepts": ['InventoryNet'], "kind": 'instant'},  # 32/37 家（覆蓋 86%、命中 1.00）
    'Other Current Assets': {"concepts": ['PrepaidExpenseAndOtherAssetsCurrent', 'OtherAssetsCurrent'], "kind": 'instant'},  # 24/41 家（覆蓋 59%、命中 1.00）
    'Total Current Assets': {"concepts": ['AssetsCurrent'], "kind": 'instant'},  # 46/46 家（覆蓋 100%、命中 1.00）
    'PP&E, net': {"concepts": ['PropertyPlantAndEquipmentNet', 'PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization', 'NoncurrentAssets'], "kind": 'instant'},  # 42/49 家（覆蓋 86%、命中 1.00）
    'Operating Lease ROU Assets': {"concepts": ['OperatingLeaseRightOfUseAsset'], "kind": 'instant'},  # 21/24 家（覆蓋 88%、命中 1.00）
    'Long-term Investments': {"concepts": ['LongTermInvestments', 'AvailableForSaleSecuritiesDebtSecuritiesNoncurrent', 'OtherLongTermInvestments'], "kind": 'instant'},  # 5/22 家（覆蓋 23%、命中 1.00）
    'Goodwill': {"concepts": ['Goodwill'], "kind": 'instant'},  # 45/47 家（覆蓋 96%、命中 0.99）
    'Intangible Assets, net': {"concepts": ['IntangibleAssetsNetExcludingGoodwill', 'FiniteLivedIntangibleAssetsNet', 'IntangibleAssetsNetIncludingGoodwill'], "kind": 'instant'},  # 31/40 家（覆蓋 78%、命中 0.99）
    'Deferred Tax Assets': {"concepts": ['DeferredIncomeTaxAssetsNet'], "kind": 'instant'},  # 23/26 家（覆蓋 88%、命中 1.00）
    'Other Non-current Assets': {"concepts": ['OtherAssetsNoncurrent'], "kind": 'instant'},  # 43/45 家（覆蓋 96%、命中 1.00）
    'Total Non-current Assets': {"concepts": ['AssetsNoncurrent'], "kind": 'instant'},  # 6/46 家（覆蓋 13%、命中 1.00）
    'Total Assets': {"concepts": ['Assets', 'LiabilitiesAndStockholdersEquity'], "kind": 'instant'},  # 51/51 家（覆蓋 100%、命中 1.00）
    'Accounts Payable': {"concepts": ['AccountsPayableCurrent', 'AccountsPayableAndAccruedLiabilitiesCurrent'], "kind": 'instant'},  # 39/48 家（覆蓋 81%、命中 1.00）
    'Short-term Debt': {"concepts": ['DebtCurrent', 'ShortTermBorrowings', 'ConvertibleDebtCurrent', 'CommercialPaper', 'OtherShortTermBorrowings'], "kind": 'instant'},  # 17/42 家（覆蓋 40%、命中 1.00）
    'Current Portion of LT Debt': {"concepts": ['LongTermDebtCurrent', 'LongTermDebtAndCapitalLeaseObligationsCurrent'], "kind": 'instant'},  # 15/21 家（覆蓋 71%、命中 1.00）
    'Op. Lease Liabilities, current': {"concepts": ['OperatingLeaseLiabilityCurrent'], "kind": 'instant'},  # 12/12 家（覆蓋 100%、命中 0.99）
    'Accrued Compensation': {"concepts": ['EmployeeRelatedLiabilitiesCurrent'], "kind": 'instant'},  # 16/16 家（覆蓋 100%、命中 1.00）
    'Deferred Revenue, current': {"concepts": ['ContractWithCustomerLiabilityCurrent'], "kind": 'instant'},  # 3/3 家（覆蓋 100%、命中 1.00）
    'Income Tax Payable': {"concepts": ['AccruedIncomeTaxesCurrent'], "kind": 'instant'},  # 14/16 家（覆蓋 88%、命中 1.00）
    'Other Current Liabilities': {"concepts": ['OtherLiabilitiesCurrent', 'LiabilitiesOfDisposalGroupIncludingDiscontinuedOperationCurrent'], "kind": 'instant'},  # 16/34 家（覆蓋 47%、命中 0.99）
    'Total Current Liabilities': {"concepts": ['LiabilitiesCurrent'], "kind": 'instant'},  # 46/49 家（覆蓋 94%、命中 1.00）
    'Long-term Debt': {"concepts": ['LongTermDebtNoncurrent', 'LongTermDebtAndCapitalLeaseObligations', 'LongTermDebt', 'LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities'], "kind": 'instant'},  # 26/40 家（覆蓋 65%、命中 0.98）
    'Op. Lease Liabilities, LT': {"concepts": ['OperatingLeaseLiabilityNoncurrent'], "kind": 'instant'},  # 22/22 家（覆蓋 100%、命中 1.00）
    # ↓ 人工補：只有 1 家有答案卷（現行路徑幾乎抓不到這列），但 concept 名稱一對一對應
    'Finance Lease Liabilities, LT': {"concepts": ['FinanceLeaseLiabilityNoncurrent'], "kind": 'instant'},
    'Deferred Revenue, LT': {"concepts": ['ContractWithCustomerLiabilityNoncurrent'], "kind": 'instant'},  # 11/11 家（覆蓋 100%、命中 1.00）
    'Deferred Tax Liability, LT': {"concepts": ['DeferredIncomeTaxLiabilitiesNet', 'AccruedIncomeTaxesNoncurrent', 'LiabilityForUncertainTaxPositionsNoncurrent'], "kind": 'instant'},  # 17/34 家（覆蓋 50%、命中 1.00）
    'Pension & Retirement Oblig.': {"concepts": ['PensionAndOtherPostretirementDefinedBenefitPlansLiabilitiesNoncurrent'], "kind": 'instant'},  # 5/6 家（覆蓋 83%、命中 1.00）
    'Other Non-current Liabilities': {"concepts": ['OtherLiabilitiesNoncurrent'], "kind": 'instant'},  # 39/45 家（覆蓋 87%、命中 1.00）
    'Total Non-current Liabilities': {"concepts": ['LiabilitiesNoncurrent'], "kind": 'instant'},  # 5/49 家（覆蓋 10%、命中 1.00）
    'Total Liabilities': {"concepts": ['Liabilities', 'LiabilitiesAndStockholdersEquity'], "kind": 'instant'},  # 42/51 家（覆蓋 82%、命中 0.99）
    'Preferred Stock': {"concepts": ['PreferredStockValue', 'LossContingencyAccrualAtCarryingValue', 'FinanceLeaseLiabilityPaymentsDueAfterYearFive'], "kind": 'instant'},  # 30/34 家（覆蓋 88%、命中 1.00）
    'Common Stock & APIC': {"concepts": ['CommonStockValue', 'CommonStocksIncludingAdditionalPaidInCapital'], "kind": 'instant'},  # 39/51 家（覆蓋 76%、命中 1.00）
    'Additional Paid-in Capital': {"concepts": ['AdditionalPaidInCapitalCommonStock', 'AdditionalPaidInCapital'], "kind": 'instant'},  # 20/39 家（覆蓋 51%、命中 1.00）
    'Treasury Stock': {"concepts": ['TreasuryStockValue'], "kind": 'instant'},  # 20/23 家（覆蓋 87%、命中 0.98）
    'Retained Earnings': {"concepts": ['RetainedEarningsAccumulatedDeficit'], "kind": 'instant'},  # 51/51 家（覆蓋 100%、命中 1.00）
    'AOCI': {"concepts": ['AccumulatedOtherComprehensiveIncomeLossNetOfTax'], "kind": 'instant'},  # 51/51 家（覆蓋 100%、命中 1.00）
    'Total Equity — Parent': {"concepts": ['StockholdersEquity'], "kind": 'instant'},  # 47/51 家（覆蓋 92%、命中 1.00）
    'Noncontrolling Interests': {"concepts": ['MinorityInterest'], "kind": 'instant'},  # 19/21 家（覆蓋 90%、命中 0.99）
    'Total Equity incl. NCI': {"concepts": ['StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest'], "kind": 'instant'},  # 22/22 家（覆蓋 100%、命中 1.00）
    'Total Liabilities & Equity': {"concepts": ['Assets', 'LiabilitiesAndStockholdersEquity'], "kind": 'instant'},  # 51/51 家（覆蓋 100%、命中 1.00）
    'Shares Outstanding': {"concepts": ['EntityCommonStockSharesOutstanding'], "kind": 'instant', "unit": 'shares', "taxonomy": 'dei'},  # 12/49 家（覆蓋 24%、命中 1.00）
}


CF_MAPPING: dict[str, dict] = {
    'Net Income': {"concepts": ['NetIncomeLoss', 'NetIncomeLossAvailableToCommonStockholdersBasic', 'NetIncomeLossAttributableToParentDiluted'], "kind": 'quarter'},  # 47/51 家（覆蓋 92%、命中 0.99）
    'D&A': {"concepts": ['DepreciationDepletionAndAmortization', 'DepreciationAndAmortization', 'Depreciation', 'DepreciationAmortizationAndAccretionNet'], "kind": 'quarter'},  # 24/48 家（覆蓋 50%、命中 0.97）
    'SBC': {"concepts": ['ShareBasedCompensation'], "kind": 'quarter'},  # 43/44 家（覆蓋 98%、命中 0.99）
    # ↓ 人工補：9 家命中、平均 0.89，差 0.01 沒過自動門檻。concept 名稱與語意一對一對應，且沒有第二個候選
    'Amortization of Intangibles': {"concepts": ['AmortizationOfIntangibleAssets'], "kind": 'quarter'},
    'Change in Receivables': {"concepts": ['IncreaseDecreaseInAccountsReceivable', 'IncreaseDecreaseInReceivables'], "kind": 'quarter', "negate": True},  # 23/35 家（覆蓋 66%、命中 0.97）
    'Change in Inventories': {"concepts": ['IncreaseDecreaseInInventories'], "kind": 'quarter', "negate": True},  # 24/24 家（覆蓋 100%、命中 0.95）
    'Change in Accounts Payable': {"concepts": ['IncreaseDecreaseInAccountsPayable', 'IncreaseDecreaseInAccountsPayableAndAccruedLiabilities'], "kind": 'quarter'},  # 26/38 家（覆蓋 68%、命中 0.99）
    'Change in Prepaid & Other Assets': {"concepts": ['IncreaseDecreaseInPrepaidDeferredExpenseAndOtherAssets'], "kind": 'quarter', "negate": True},  # 16/16 家（覆蓋 100%、命中 0.98）
    'Change in Other Operating Assets': {"concepts": ['IncreaseDecreaseInOtherOperatingAssets'], "kind": 'quarter', "negate": True},  # 16/18 家（覆蓋 89%、命中 0.98）
    'Change in Deferred Revenue': {"concepts": ['IncreaseDecreaseInContractWithCustomerLiability', 'IncreaseDecreaseInDeferredRevenue'], "kind": 'quarter'},  # 16/21 家（覆蓋 76%、命中 0.99）
    'Other Working Capital': {"concepts": ['IncreaseDecreaseInOtherOperatingCapitalNet', 'IncreaseDecreaseInOperatingCapital'], "kind": 'quarter', "negate": True},  # 9/25 家（覆蓋 36%、命中 0.97）
    'Other Non-cash Items': {"concepts": ['OtherNoncashIncomeExpense'], "kind": 'quarter', "negate": True},  # 29/31 家（覆蓋 94%、命中 0.98）
    'Operating Cash Flow': {"concepts": ['NetCashProvidedByUsedInOperatingActivities', 'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations'], "kind": 'quarter'},  # 51/51 家（覆蓋 100%、命中 1.00）
    'Capex': {"concepts": ['PaymentsToAcquirePropertyPlantAndEquipment', 'SegmentExpenditureAdditionToLongLivedAssets'], "kind": 'quarter', "negate": True},  # 31/39 家（覆蓋 79%、命中 0.98）
    'Acquisitions': {"concepts": ['PaymentsToAcquireBusinessesNetOfCashAcquired'], "kind": 'quarter'},  # 17/40 家（覆蓋 42%、命中 0.97）
    'Investment Purchases': {"concepts": ['PaymentsToAcquireInvestments', 'PaymentsToAcquireMarketableSecurities', 'PaymentsToAcquireShortTermInvestments'], "kind": 'quarter', "negate": True},  # 9/36 家（覆蓋 25%、命中 0.98）
    'Investment Proceeds': {"concepts": ['ProceedsFromSaleAndMaturityOfMarketableSecurities', 'ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities', 'ProceedsFromSaleMaturityAndCollectionOfShorttermInvestments', 'ProceedsFromSaleOfAvailableForSaleSecuritiesDebt'], "kind": 'quarter'},  # 7/36 家（覆蓋 19%、命中 0.97）
    'Investing Cash Flow': {"concepts": ['NetCashProvidedByUsedInInvestingActivities', 'NetCashProvidedByUsedInInvestingActivitiesContinuingOperations'], "kind": 'quarter'},  # 49/50 家（覆蓋 98%、命中 1.00）
    'Debt Proceeds': {"concepts": ['ProceedsFromIssuanceOfLongTermDebt'], "kind": 'quarter'},  # 14/37 家（覆蓋 38%、命中 0.99）
    'Debt Repayments': {"concepts": ['RepaymentsOfLongTermDebt'], "kind": 'quarter', "negate": True},  # 10/46 家（覆蓋 22%、命中 0.95）
    'Share Repurchases': {"concepts": ['PaymentsForRepurchaseOfCommonStock'], "kind": 'quarter', "negate": True},  # 31/37 家（覆蓋 84%、命中 0.98）
    'Dividends Paid': {"concepts": ['PaymentsOfDividends', 'PaymentsOfDividendsCommonStock', 'PaymentsOfOrdinaryDividends'], "kind": 'quarter', "negate": True},  # 17/40 家（覆蓋 42%、命中 0.96）
    'Financing Cash Flow': {"concepts": ['NetCashProvidedByUsedInFinancingActivities', 'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations'], "kind": 'quarter'},  # 50/50 家（覆蓋 100%、命中 1.00）
    'FX Effect on Cash': {"concepts": ['EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents', 'EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsIncludingDisposalGroupAndDiscontinuedOperations', 'EffectOfExchangeRateOnCashAndCashEquivalents'], "kind": 'quarter'},  # 28/39 家（覆蓋 72%、命中 1.00）
    'Net Change in Cash': {"concepts": ['CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect', 'CashAndCashEquivalentsPeriodIncreaseDecrease'], "kind": 'quarter'},  # 51/51 家（覆蓋 100%、命中 1.00）
    # ↓ 人工補：期末現金是時點值不是期間值。覆蓋率低是因為各家「含不含受限現金」口徑不同，兩個 concept 排成 fallback
    'Ending Cash': {"concepts": ['CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents', 'CashAndCashEquivalentsAtCarryingValue'], "kind": 'instant'},
    'Cash Taxes Paid': {"concepts": ['IncomeTaxesPaidNet'], "kind": 'quarter'},  # 10/17 家（覆蓋 59%、命中 1.00）
    'Cash Interest Paid': {"concepts": ['InterestPaidNet'], "kind": 'quarter'},  # 13/16 家（覆蓋 81%、命中 0.98）
    # 'Free Cash Flow'：OCF − Capex。XBRL 沒有這個 tag，必須用算的（模板本來就標 source=DERIVED）
}


# `fetcher_facts.build_statement_tables()` 吃的形狀
ALL_MAPPINGS: dict[str, dict] = {"IS": IS_MAPPING, "BS": BS_MAPPING, "CF": CF_MAPPING}
