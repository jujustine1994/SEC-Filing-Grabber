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


def test_no_identified_total_accepts_unique_exact_revenue_label():
    df=frame(['company_CustomRevenue'],[100])
    assert fg._match_revenue_row(df,'2020-12-31 (FY)') == (0,False)


@pytest.mark.parametrize('label', ['Revenue', 'REVENUES', ' Revenue: ', 'Total revenue',
                                    'Total Revenues', 'Total net revenues'])
def test_custom_revenue_requires_exact_original_label(label):
    df = frame(['company_CustomRevenue'], [100])
    df['label'] = [label]
    assert fg._match_revenue_row(df, '2020-12-31 (FY)') == (0, False)


@pytest.mark.parametrize('label', ['Management Fees Revenue', 'Total Revenues and Other Income',
                                    'Sales Revenue Goods', 'Unrelated'])
def test_normalized_revenue_does_not_authorize_a_component(label):
    df = frame(['company_CustomRevenue'], [100])
    df['label'] = [label]
    assert fg._match_revenue_row(df, '2020-12-31 (FY)') == (None, False)


def test_bare_revenue_with_competing_component_is_ambiguous():
    df = frame(['company_CustomRevenue', 'company_ManagementFeesRevenue'], [100, 10])
    df['label'] = ['Revenue', 'Management Fees Revenue']
    assert fg._match_revenue_row(df, '2020-12-31 (FY)') == (None, True)


def test_cost_of_revenue_is_not_a_competing_revenue_candidate():
    df = frame(['us-gaap_SalesRevenueGoodsNet', 'us-gaap_CostOfGoodsSold'], [100, 60])
    df['label'] = ['Revenue', 'Cost of Revenue']
    df['standard_concept'] = ['Revenue', 'CostOfGoodsAndServicesSold']
    assert fg._match_revenue_row(df, '2020-12-31 (FY)') == (0, False)


def test_exact_custom_totals_with_different_values_are_ambiguous():
    df = frame(['company_A', 'company_B'], [100, 90])
    df['label'] = ['Total Revenue', 'Total revenues']
    assert fg._match_revenue_row(df, '2020-12-31 (FY)') == (None, True)


def test_custom_revenue_cannot_be_replaced_by_legacy_override():
    df = frame(['company_CustomRevenue', 'company_Fees'], [100, 10])
    df['label'] = ['Total Revenue', 'Management Fees Revenue']
    df['standard_concept'] = ['Revenue', 'Fees']
    filing = MagicMock()
    filing.obj.return_value.financials.income_statement.return_value.to_dataframe.return_value = df
    filing.obj.return_value.financials.cashflow_statement.return_value = None
    table, _ = fg._build_is_table([filing], max_filings=1, is_overrides={
        'Revenue': {'fix_type': 'concept_override', 'std_concept': 'Fees'}})
    assert table.values[table.concepts.index('Revenue')] == [100]


def test_builder_uses_total_and_preserves_component_overflow():
    df=frame(['us-gaap_ManagementFeesBaseRevenue','us-gaap_Revenues'],[5,100])
    filing=MagicMock();filing.filing_date='2021-02-01'
    fin=filing.obj.return_value.financials
    fin.income_statement.return_value.to_dataframe.return_value=df
    fin.cashflow_statement.return_value=None
    table,_=fg._build_is_table([filing],max_filings=1)
    assert table.values[table.concepts.index('Revenue')]==[100]
    assert table.values[table.labels.index('us-gaap_ManagementFeesBaseRevenue')]==[5]


def test_ambiguous_total_is_data_gap_without_network_probe():
    df=frame(['us-gaap_Revenues','us-gaap_SalesRevenueNet'],[100,90])
    filing=MagicMock();filing.filing_date='2021-02-01'
    fin=filing.obj.return_value.financials
    fin.income_statement.return_value.to_dataframe.return_value=df
    fin.cashflow_statement.return_value=None
    probe=MagicMock(side_effect=AssertionError('No network probe'))
    with fg.collect_gaps(FetchLedger(probe=probe)) as ledger:
        table,_=fg._build_is_table([filing],max_filings=1)
    assert table.values[table.concepts.index('Revenue')]==[None]
    assert any(g.exc_name=='AmbiguousRevenueTotal' and g.kind=='data' for g in ledger.gaps)
    probe.assert_not_called()


@pytest.mark.parametrize('override',[
    {'fix_type':'concept_override','std_concept':'Revenue'},
    {'fix_type':'structural_absence'},
])
def test_recognized_total_survives_existing_or_diagnosis_override(override):
    df=frame(['us-gaap_ManagementFeesBaseRevenue','us-gaap_Revenues'],[5,100])
    filing=MagicMock();filing.filing_date='2021-02-01'
    fin=filing.obj.return_value.financials
    fin.income_statement.return_value.to_dataframe.return_value=df
    fin.cashflow_statement.return_value=None
    table,_=fg._build_is_table([filing],max_filings=1,is_overrides={'Revenue':override})
    assert table.values[table.concepts.index('Revenue')]==[100]
    assert table.values[table.labels.index('us-gaap_ManagementFeesBaseRevenue')]==[5]


def test_diagnosis_rebuild_cannot_fill_rejected_conflicting_totals():
    df=frame(['us-gaap_Revenues','us-gaap_SalesRevenueNet'],[100,90])
    filing=MagicMock();filing.filing_date='2021-02-01'
    fin=filing.obj.return_value.financials
    fin.income_statement.return_value.to_dataframe.return_value=df
    fin.cashflow_statement.return_value=None
    with fg.collect_gaps(FetchLedger(probe=lambda:True)) as ledger:
        before,_=fg._build_is_table([filing],max_filings=1)
        rebuilt,_=fg._build_is_table([filing],max_filings=1,is_overrides={
            'Revenue':{'fix_type':'concept_override','std_concept':'Revenue'}})
    assert before.values[before.concepts.index('Revenue')]==[None]
    assert rebuilt.values[rebuilt.concepts.index('Revenue')]==[None]
    assert len([g for g in ledger.gaps if g.exc_name=='AmbiguousRevenueTotal'])==2


def test_calculation_child_is_not_an_independent_conflicting_total():
    df=frame(['us-gaap_SalesRevenueNet','us-gaap_Revenues'],[31_624_000_000,32_361_000_000])
    df['parent_concept']=['us-gaap_Revenues','us-gaap_OperatingIncomeLoss']
    df['label']=['Net sales','Total revenue']
    assert fg._match_revenue_row(df,'2020-12-31 (FY)')==(1,False)


def test_total_including_other_income_is_not_automatic_revenue_authority():
    df=frame(['us-gaap_SalesRevenueNet','us-gaap_Revenues'],[29_106_000_000,32_584_000_000])
    df['parent_concept']=['us-gaap_Revenues','us-gaap_OperatingIncomeLoss']
    df['label']=['Sales and other operating revenues','Total Revenues and Other Income']
    assert fg._match_revenue_row(df,'2020-12-31 (FY)')==(None,True)


def test_absent_value_is_not_a_conflicting_reported_amount():
    df=frame(['us-gaap_SalesRevenueNet','us-gaap_Revenues'],[None,37_666_000_000])
    assert fg._match_revenue_row(df,'2020-12-31 (FY)')==(1,False)


def test_missing_parent_total_does_not_fall_back_to_its_sales_child():
    df=frame(['us-gaap_SalesRevenueNet','us-gaap_Revenues'],[90,None])
    df['parent_concept']=['us-gaap_Revenues','us-gaap_OperatingIncomeLoss']
    assert fg._match_revenue_row(df,'2020-12-31 (FY)')==(1,False)
