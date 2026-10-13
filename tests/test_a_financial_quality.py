from datetime import date
from types import SimpleNamespace
from unittest.mock import patch

import pandas as pd
import pytest
import fetcher_gaap as fg
from override_engine import e1_fuzzy_match


def frame(rows, column='2026-06-27 (Q2)'):
    return pd.DataFrame([dict(concept=c, standard_concept=s, label=l,
                              abstract=False, dimension=False, is_breakdown=False, **{column: v})
                         for c, s, l, v in rows])


def filing(df):
    stmt = SimpleNamespace(to_dataframe=lambda: df.copy(deep=True))
    fin = SimpleNamespace(income_statement=lambda: stmt, balance_sheet=lambda: stmt,
                          cashflow_statement=lambda: None)
    return SimpleNamespace(filing_date=date(2026, 8, 5), accession_no='test',
                           obj=lambda: SimpleNamespace(financials=fin))


def value(table, name):
    return table.values[table.concepts.index(name)][0]


def test_q4_never_subtracts_weighted_shares_or_eps():
    names = ['Revenue', 'Basic Shares', 'Diluted Shares', 'Basic EPS', 'Diluted EPS']
    q = fg.StatementTable('Data_IS', ['FY2024Q1', 'FY2024Q2', 'FY2024Q3'], ['', '', ''],
                          names, [[10, 20, 30], [100, 110, 120], [105, 115, 125],
                                  [0.1, 0.2, 0.3], [0.09, 0.19, 0.29]])
    a = fg.StatementTable('Data_IS', ['FY2024'], [], names,
                          [[100], [115], [120], [0.9], [0.8]])
    result = fg._synthesize_q4(q, a, len(names), False)
    assert result.values[0][-1] == 40
    assert [row[-1] for row in result.values[1:]] == [None] * 4


def test_cost_uses_reported_total_not_excluding_da_child():
    df = frame([
        ('us-gaap_Revenues', 'Revenue', 'Revenue', 100),
        ('us-gaap_CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization',
         'CostOfGoodsAndServicesSold', 'Cost of sales', 40),
        ('us-gaap_CostDepreciationAmortizationAndDepletion', 'DepreciationExpense', 'Amortization', 5),
        ('us-gaap_CostOfGoodsAndServicesSold', 'CostOfGoodsAndServicesSold', 'Total cost of sales', 45),
        ('us-gaap_GrossProfit', 'GrossProfit', 'Gross profit', 55),
    ])
    with patch.object(fg, '_filing_obj', side_effect=lambda f: f.obj()):
        table, _ = fg._build_is_table([filing(df)], 1)
    assert value(table, 'Cost of Revenue') == 45


@pytest.mark.parametrize('nci,total_equity,expected', [(None, None, 30), (2, None, None), (2, 72, 28)])
def test_liabilities_excludes_equity_and_guards_nci_scope(nci, total_equity, expected):
    rows = [('us-gaap_Assets', 'Assets', 'Total assets', 100),
            ('us-gaap_StockholdersEquity', 'AllEquityBalance', 'Parent equity', 70),
            ('us-gaap_LiabilitiesAndStockholdersEquity', 'LiabilitiesAndEquity', 'Total liabilities and equity', 100)]
    if nci is not None:
        rows.append(('us-gaap_MinorityInterest', 'MinorityInterestBalance', 'Noncontrolling interest', nci))
    if total_equity is not None:
        rows.append(('us-gaap_StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
                     'AllEquityBalanceIncludingMinorityInterest', 'Total equity', total_equity))
    df = frame(rows, '2026-06-27')
    with patch.object(fg, '_filing_obj', side_effect=lambda f: f.obj()):
        table, _ = fg._build_bs_table([filing(df)], 1,
                                    bs_overrides={'Total Liabilities': {'fix_type': 'concept_override', 'std_concept': 'LiabilitiesAndEquity'}})
    assert value(table, 'Total Liabilities') == expected


def test_fuzzy_liabilities_does_not_select_equity_total():
    df = frame([('us-gaap_LiabilitiesAndStockholdersEquity', 'LiabilitiesAndEquity',
                 'Total liabilities and stockholders equity', 100)])
    assert e1_fuzzy_match(df, 'Total Liabilities') is None


def test_income_bridge_preserves_distinct_reported_concepts():
    df = frame([
        ('us-gaap_IncomeLossFromEquityMethodInvestments', 'EquityMethodInvestmentIncome', 'Equity income', 6),
        ('us-gaap_IncomeLossFromContinuingOperations', 'IncomeLossContinuingOperations', 'Continuing net income', 2305),
        ('us-gaap_IncomeLossFromDiscontinuedOperationsNetOfTaxAttributableToReportingEntity',
         'DiscontinuedOperationsIncome', 'Discontinued, net of tax', -8),
    ])
    with patch.object(fg, '_filing_obj', side_effect=lambda f: f.obj()):
        table, _ = fg._build_is_table([filing(df)], 1)
    assert value(table, 'Equity Method Income/(Loss)') == 6
    assert value(table, 'Net Income from Continuing Operations') == 2305
    assert value(table, 'Discontinued Operations, net of tax') == -8


def test_eps_selects_total_not_continuing_operations():
    df = frame([
        ('custom_EarningsPerShareBasicFromContinuingOperations', None, 'Continuing basic EPS', 1.42),
        ('us-gaap_EarningsPerShareBasic', None, 'Basic earnings per share', 1.41),
        ('custom_EarningsPerShareDilutedFromContinuingOperations', None, 'Continuing diluted EPS', 1.40),
        ('us-gaap_EarningsPerShareDiluted', None, 'Diluted earnings per share', 1.39),
    ])
    with patch.object(fg, '_filing_obj', side_effect=lambda f: f.obj()):
        table, _ = fg._build_is_table([filing(df)], 1)
    assert value(table, 'Basic EPS') == 1.41
    assert value(table, 'Diluted EPS') == 1.39


def test_cost_total_does_not_conflict_with_product_cost_child():
    # MSFT historical layout: both product cost and the complete goods/services
    # total use recognized GAAP concepts; their calculation scopes differ.
    df = frame([
        ('us-gaap_CostOfGoodsSold', 'CostOfGoodsAndServicesSold', 'Product', 40),
        ('example_CostOfServices', None, 'Service and other', 5),
        ('us-gaap_CostOfRevenue', 'CostOfGoodsAndServicesSold', 'Total cost of revenue', 45),
    ])
    df['parent_concept'] = ['us-gaap_CostOfRevenue', 'us-gaap_CostOfRevenue', 'us-gaap_GrossProfit']
    with patch.object(fg, '_filing_obj', side_effect=lambda f: f.obj()):
        table, _ = fg._build_is_table([filing(df)], 1)
    assert value(table, 'Cost of Revenue') == 45


@pytest.mark.parametrize('parent', ['us-gaap:CostOfRevenue', 'us-gaap_CostOfGoodsSold'])
def test_cost_fallback_cannot_reintroduce_excluded_child(parent):
    df = frame([('custom_DirectCost', 'CostOfGoodsAndServicesSold', 'Cost of sales', 40)])
    df['parent_concept'] = parent
    with patch.object(fg, '_filing_obj', side_effect=lambda f: f.obj()):
        table, _ = fg._build_is_table([filing(df)], 1)
    assert value(table, 'Cost of Revenue') is None
