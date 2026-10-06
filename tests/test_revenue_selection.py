import pandas as pd
import pytest
from unittest.mock import MagicMock
import fetcher_gaap as fg
from fetch_ledger import FetchLedger


def frame(concepts, values, standards=None):
    return pd.DataFrame({'concept': concepts, 'label': ['Revenue'] * len(concepts),
        'standard_concept': standards or ['Revenue'] * len(concepts),
        'abstract': False, 'is_breakdown': False, 'dimension_member_label': None,
        '2020-12-31 (FY)': values})


@pytest.mark.parametrize('component,total', [(530_000_000,10_908_000_000), (83_000_000,2_161_000_000)])
def test_source_confirmed_management_fees_do_not_win_over_total(component,total):
    df=frame(['us-gaap_ManagementFeesBaseRevenue','us-gaap_Revenues'],[component,total])
    before=df.copy(deep=True)
    assert fg._match_revenue_row(df,'2020-12-31 (FY)') == (1,False)
    pd.testing.assert_frame_equal(df,before)


def test_total_is_selected_by_concept_even_when_smaller():
    df=frame(['us-gaap_SalesRevenueGoodsNet','us-gaap_SalesRevenueNet'],[300,200])
    assert fg._match_revenue_row(df,'2020-12-31 (FY)') == (1,False)


def test_raw_total_does_not_require_correct_standardization():
    df=frame(['us-gaap_ManagementFeesBaseRevenue','us-gaap_Revenues'],[5,100],['Revenue',None])
    assert fg._match_revenue_row(df,'2020-12-31 (FY)') == (1,False)


def test_broad_total_precedes_contract_customer_subset():
    df=frame(['us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax','us-gaap_Revenues'],[50,100])
    assert fg._match_revenue_row(df,'2020-12-31 (FY)') == (1,False)


def test_dimensional_total_is_not_consolidated_authority():
    df=frame(['us-gaap_Revenues','us-gaap_Revenues'],[20,100])
    df.loc[0,'dimension_member_label']='Segment A'
    assert fg._match_revenue_row(df,'2020-12-31 (FY)') == (1,False)


def test_duplicate_total_with_equal_current_values_is_safe():
    df=frame(['us-gaap_Revenues','us-gaap_Revenues'],[100,100])
    assert fg._match_revenue_row(df,'2020-12-31 (FY)') == (0,False)


def test_conflicting_totals_cannot_be_resolved_by_order():
    df=frame(['us-gaap_Revenues','us-gaap_SalesRevenueNet'],[100,90])
    assert fg._match_revenue_row(df,'2020-12-31 (FY)') == (None,True)


def test_missing_total_value_does_not_fall_back_to_component():
    df=frame(['us-gaap_ManagementFeesBaseRevenue','us-gaap_Revenues'],[5,None])
    assert fg._match_revenue_row(df,'2020-12-31 (FY)') == (1,False)


def test_no_identified_total_retains_existing_policy():
    df=frame(['company_CustomRevenue'],[100])
    assert fg._match_revenue_row(df,'2020-12-31 (FY)') == (0,False)


def test_builder_uses_total_and_preserves_component_overflow():
    df=frame(['us-gaap_ManagementFeesBaseRevenue','us-gaap_Revenues'],[5,100])
    filing=MagicMock();filing.filing_date='2021-02-01'
    fin=filing.obj.return_value.financials
    fin.income_statement.return_value.to_dataframe.return_value=df
    fin.cashflow_statement.return_value=None
    table,_=fg._build_is_table([filing])
    assert table.values[table.concepts.index('Revenue')]==[100]
    assert table.values[table.concepts.index('us-gaap_ManagementFeesBaseRevenue')]==[5]


def test_ambiguous_total_is_data_gap_without_network_probe():
    df=frame(['us-gaap_Revenues','us-gaap_SalesRevenueNet'],[100,90])
    filing=MagicMock();filing.filing_date='2021-02-01'
    fin=filing.obj.return_value.financials
    fin.income_statement.return_value.to_dataframe.return_value=df
    fin.cashflow_statement.return_value=None
    probe=MagicMock(side_effect=AssertionError('No network probe'))
    with fg.collect_gaps(FetchLedger(probe=probe)) as ledger:
        table,_=fg._build_is_table([filing])
    assert table.values[table.concepts.index('Revenue')]==[None]
    assert any(g.exc_name=='AmbiguousRevenueTotal' and g.kind=='data' for g in ledger.gaps)
    probe.assert_not_called()
