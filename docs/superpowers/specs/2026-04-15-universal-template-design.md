# Universal Financial Statement Template Design

**Date:** 2026-04-15  
**Status:** Exploration complete — ready for implementation planning  
**Goal:** Design a fixed IS/BS/CF template that works across all industries, merge three statements into one sheet (`Data_Financials`), add Original Item (B column).

---

## Target Output Format (reference: MSFT model)

- Single sheet `Data_Financials` containing IS + BS + CF in order
- Section headers as separator rows: "Income Statement", "Balance Sheet", "Cash Flow"
- **Col A** = Std Name (our standardized label)
- **Col B** = Original Item (company's XBRL label from edgartools)
- **Col C+** = quarterly data, oldest → newest
- Row 1: ticker (A1) + quarter labels (C1+)
- Row 2: empty (A2) + filing dates (C2+)
- Segment breakdowns remain in separate `Data_Seg_*` sheets

---

## Known Issues & Fixes

### Issue 1: CF aggregate rows have duplicate standard_concept
**Problem:** `NetCashFromOperatingActivities` appears 3–4 times per filing (intermediate items + final total). Current `_match_is_row` takes first match → picks wrong intermediate row.  
**Affected:** BA (4x), AMD (3x), also Investing CF and Financing CF.  
**Fix:** Add 5th field `match: "first"|"last"` to template tuple. Aggregate CF rows use `"last"` (final total is always the last occurrence).  
**Also affects:** `CashAndCashEquivalents` (beginning + ending balance — want last/ending).

### Issue 2: BS rows with same standard_concept
**Problem:** Multiple BS rows share the same standard_concept:
- `CashAndMarketableSecurities`: GS has 4 rows (Cash, Trading assets, AFS, HTM)
- `TradeReceivables`: BA has 3 rows (AR, Unbilled, Financing receivables)
- `OtherNonOperatingCurrentAssets`: duplicated in most companies
**Fix:** Add `label_hint` optional field to template — prefer rows whose label contains the hint string (case-insensitive). E.g., Cash row uses `label_hint="cash and cash equivalents"`.

### Issue 3: BA Net Income uses `ProfitLoss` not `NetIncome`
**Problem:** Boeing CF uses `standard_concept = "ProfitLoss"` for net income instead of `"NetIncome"`. IS template uses `"NetIncome"` → Boeing would show None.  
**Fix:** IS template Net Income row: add `ProfitLoss` as additional fallback std_concept to try.

### Issue 4: AMD D&A wrongly mapped by edgartools
**Problem:** AMD's "Depreciation and amortization" in CF is mapped to `NonoperatingIncomeExpense` (wrong) instead of `DepreciationExpense`.  
**Fix:** D&A template row: add label-based fallback matching rows whose label contains "depreciation".

### Issue 5: AdditionalPaidInCapital — split vs combined
**Problem:** AAPL/MSFT combine Common Stock + APIC into single `CommonEquity` row. NVDA/AMD/GS/JPM/BA have separate `AdditionalPaidInCapital`.  
**Fix:** BS template includes BOTH rows. Companies that combine will show None for `AdditionalPaidInCapital`.

### Issue 9: `ProfitLoss` is more common than expected
**Problem:** BA, TSLA, XOM, WMT all use `ProfitLoss` (Net income including noncontrolling interests) instead of `NetIncome`. Affects both IS and CF templates.  
**Fix:** Both Net Income rows try `NetIncome` first, fallback to `ProfitLoss`.

### Issue 10: TSLA D&A has no standard_concept
**Problem:** TSLA CF "Depreciation, amortization and impairment" has `standard_concept = nan`. `DepreciationExpense` match fails.  
**Fix:** D&A matching: `DepreciationExpense` std → fallback concept suffix `DepreciationDepletion` → fallback label containing "depreciation".

### Issue 11: Treasury Stock embedded in CommonEquity for XOM/JNJ
**Problem:** XOM and JNJ label their treasury stock rows as `CommonEquity` standard_concept (e.g. "Common stock held in treasury"). First-match would pick the par value row, which is correct — but the treasury row also matches.  
**Fix:** `label_hint="common stock"` (not "treasury") on the CommonEquity row.

### Issue 12: WMT LongTermDebt appears twice
**Problem:** WMT has both "Long-term debt" and "Long-term finance lease obligations" as `LongTermDebt`.  
**Fix:** `label_hint="long-term debt"` (not "lease") on the LongTermDebt row; add separate Finance Lease LT row.

### Issue 13: GOOGL encoding error
**Problem:** Some GOOGL BS label contains a non-ASCII character causing `cp950` codec error during output.  
**Fix:** Ensure `_to_python_val` and all string processing uses UTF-8 or strips/replaces non-encodable characters gracefully.

### Issue 6: Treasury Stock not universal
**Problem:** AMD, GS, BA have explicit `TreasuryShares`; others embed it elsewhere.  
**Fix:** Include Treasury Stock row in BS template; None for companies that don't report separately.

### Issue 7: InvestmentProceeds appears multiple times
**Problem:** Multiple investment proceeds rows (maturities, sales, etc.) all share `InvestmentProceeds`. No single aggregate row.  
**Current decision:** Take first match (total investment proceeds not separately aggregated in XBRL). Flag as known limitation.

### Issue 8: Three statements → one sheet
**Problem:** Current architecture produces 3 separate StatementTable objects (Data_IS, Data_BS, Data_CF).  
**Fix:** New merge function in fetcher_gaap.py assembles IS + BS + CF into single `StatementTable` with section header rows. excel_writer.py updated to support B-column Original Item (data shifts to C+).

---

## Template Matching Enhancement

Current tuple: `(label, std_concept, fallback_suffix, source)`  
Proposed tuple: `(label, std_concept, fallback_suffix, source, match, label_hint)`

| Field | Type | Description |
|-------|------|-------------|
| `label` | str | Display name (Col A) |
| `std_concept` | str\|None | Primary match: `standard_concept == value` |
| `fallback_suffix` | str | Secondary match: `concept.contains(value)` |
| `source` | "IS"\|"CF"\|"BS" | Which statement to pull from |
| `match` | "first"\|"last" | Which occurrence to use when multiple rows match |
| `label_hint` | str\|None | Tertiary filter: prefer rows whose label contains this string |

---

## Financial Companies

**Finding:** GS/JPM have fundamentally different IS and BS structure:
- IS: Net Interest Income, Provision for Credit Losses, Non-interest Revenue/Expense
- BS: No current/non-current split; dominated by Trading Assets, Loans, Deposits, Repos
- CF: Structure similar but operating items completely different

**Decision:**
- Universal template works for non-financials (tech, industrial, energy, pharma, retail)
- Financial companies need a separate `FINANCIAL_IS_TEMPLATE` and `FINANCIAL_BS_TEMPLATE`
- Auto-detection: check if BS contains `TotalDeposits` standard_concept → financial company
- UI: after fetch, if financial company detected → show warning and offer to switch to financial template

**Financial template** design deferred until universal template is complete.

---

## Exploration Status

Companies explored for BS/CF standard_concept coverage:

| Company | Industry | BS | CF | Notes |
|---------|----------|----|----|-------|
| AAPL | Tech | ✅ | ✅ | |
| MSFT | Tech | ✅ | ✅ | |
| NVDA | Semiconductor | ✅ | ✅ | |
| AMD | Semiconductor | ✅ | ✅ | D&A mapping bug |
| GS | Investment Bank | ✅ | ✅ | Financial co. |
| JPM | Bank | ✅ | ✅ | Financial co. |
| BA | Industrial/Aerospace | ✅ | ✅ | ProfitLoss issue |
| META | Internet/Advertising | ✅ | ✅ | |
| GOOGL | Internet/Advertising | ✅ | ✅ | Encoding error in BS (partial) |
| TSLA | Auto/Manufacturing | ✅ | ✅ | ProfitLoss + D&A nan issue |
| XOM | Energy | ✅ | ✅ | ProfitLoss; Treasury in CommonEquity |
| JNJ | Pharma | ✅ | ✅ | Treasury in CommonEquity |
| WMT | Retail | ✅ | ✅ | ProfitLoss |

---

## Proposed IS Template (21 rows — unchanged from current)

Current 21-row template is adequate. Will be reviewed after exploration complete.

---

## Proposed BS Template (~41 rows — finalized)

**Current Row Candidates:**

Assets:
1. Cash (`CashAndMarketableSecurities`, label_hint="cash and cash equivalents")
2. Short-term Investments (`ShortTermInvestments`)
3. Accounts Receivable (`TradeReceivables`, label_hint="accounts receivable")
4. Inventories (`Inventories`, match="first")
5. Other Current Assets (`OtherNonOperatingCurrentAssets`, label_hint="other current")
6. Total Current Assets (`CurrentAssetsTotal`)
7. PP&E, net (`PlantPropertyEquipmentNet`)
8. Operating Lease ROU Assets (`OperatingLeaseRightOfUseAsset`)
9. Long-term Investments (`LongtermInvestments`)
10. Goodwill (`Goodwill`)
11. Intangible Assets, net (`IntangibleAssets`)
12. Deferred Tax Assets (`DeferredTaxNoncurrentAssets`)
13. Other Non-current Assets (`OtherNonOperatingNonCurrentAssets`, label_hint="other", match="last")
14. Total Assets (`Assets`, match="last")

Liabilities:
15. Accounts Payable (`TradePayables`)
16. Short-term Debt (`ShortTermDebt`, match="first")
17. Current Portion of LT Debt (`CurrentPortionOfLongTermDebt`)
18. Operating Lease Liabilities, current (`OperatingLeaseCurrentDebtEquivalent`)
19. Accrued Compensation (`AccruedCompensation`)
20. Deferred Revenue, current (`OtherOperatingCurrentLiabilities`, label_hint="unearned revenue")
21. Income Tax Payable (`AccruedIncomeTaxes`)
22. Other Current Liabilities (`OtherNonOperatingCurrentLiabilities`, match="first")
23. Total Current Liabilities (`CurrentLiabilitiesTotal`)
24. Long-term Debt (`LongTermDebt`, label_hint="long-term debt")
25. Operating Lease Liabilities, LT (`OperatingLeaseNonCurrentDebtEquivalent`)
26. Finance Lease Liabilities, LT (label fallback "finance lease", LT)
27. Deferred Revenue, LT (`ContractLiabilities`)
28. Deferred Tax Liability, LT (`DeferredTaxNonCurrentLiabilities`)
29. Pension & Retirement Obligations (`PensionObligations`)
30. Other Non-current Liabilities (`OtherNonOperatingNonCurrentLiabilities`, match="first")
31. Total Liabilities (`Liabilities`, match="last")

Equity:
32. Preferred Stock (`PreferredStock`)
33. Common Stock & APIC (`CommonEquity`, label_hint="common stock")
34. Additional Paid-in Capital (`AdditionalPaidInCapital`)
35. Treasury Stock (`TreasuryShares`)
36. Retained Earnings (`RetainedEarnings`)
37. AOCI (`AccumulatedOtherComprehensiveIncome`)
38. Total Equity — Parent (`AllEquityBalance`, match="first")
39. Noncontrolling Interests (`MinorityInterestBalance`)
40. Total Equity incl. NCI (`AllEquityBalanceIncludingMinorityInterest`)
41. Total Liabilities & Equity (`LiabilitiesAndEquity`)

---

## Proposed CF Template (~25 rows — finalized)

Operating:
1. Net Income (`NetIncome`, fallback `ProfitLoss`)
2. D&A (`DepreciationExpense`, fallback label "depreciation")
3. SBC (`StockBasedCompensationExpense`)
4. Amortization of Intangibles (`AmortizationOfIntangibles`)
5. Change in Receivables (`ChangeInReceivables`)
6. Change in Inventories (label fallback "inventories")
7. Change in Deferred Revenue (`ChangeInDeferredRevenue`)
8. Other Working Capital (`ChangeInOtherWorkingCapital`)
9. Other Non-cash Items (`OtherNonCashItemsCF`)
10. **Operating Cash Flow** (`NetCashFromOperatingActivities`, match="last")

Investing:
11. Capex (`CapitalExpenses`, match="first", label_hint="property")
12. Acquisitions (`AcquisitionsNet`)
13. Investment Purchases (`InvestmentPurchases`, match="first")
14. Investment Proceeds (`InvestmentProceeds`, match="first")
15. **Investing Cash Flow** (`NetCashFromInvestingActivities`, match="last")

Financing:
16. Debt Proceeds (`DebtProceeds`)
17. Debt Repayments (`DebtRepayments`)
18. Share Repurchases (`EquityExpenseIncome(BuybackIssued)`, label_hint="repurchas")
19. Dividends Paid (`DistributionsToMinorityInterests`, label_hint="dividend")
20. **Financing Cash Flow** (`NetCashFromFinancingActivities`, match="last")

Other:
21. FX Effect on Cash (`ForeignExchangeEffectOnCash`)
22. **Net Change in Cash** (`NetChangeInCash`)
23. Ending Cash (`CashAndCashEquivalents`, match="last")
24. Cash Taxes Paid (`IncomeTaxes`, source=CF, label_hint="income tax")
25. Cash Interest Paid (`InterestExpense`, source=CF, label_hint="interest paid") — financial co.

*Derived (computed, not from XBRL):*
26. Free Cash Flow = Operating CF − Capex

---

## Next Steps

1. ✅ Run exploration on META, GOOGL, TSLA, XOM, JNJ, WMT
2. ✅ Finalize BS and CF templates based on full exploration
3. ✅ Write implementation plan
4. ✅ Implement: template enhancement (`match`, `label_hint`), three-statement merge, B-column Original Item
5. ⬜ Real-company smoke test: AAPL, TSLA, BA, XOM — verify Data_Financials output
6. ⬜ Verify main.py: GUI references to Data_IS/BS/CF need updating to Data_Financials
7. ⬜ Design and implement financial company template (GS, JPM)
8. ⬜ Excel Template colouring feature (template.xlsx-based)
