import pandas as pd
import pytest
import fetcher_gaap as fg


@pytest.mark.parametrize('concept', ['jnj_Restructuringchargepercenttosales',
                                    'jnj_RestructuringChargePercentToSales'])
def test_percentage_is_not_an_independent_usd_revenue_component(concept):
    df = pd.DataFrame({
        'concept': ['us-gaap_SalesRevenueGoodsNet', concept],
        'label': ['Sales to customers (Note 9)', 'Restructuring charge percent to sales'],
        'standard_concept': ['Revenue', None],
        'parent_concept': ['us-gaap_GrossProfit', 'jnj_IncomeBeforeTaxesPercentToSales'],
        'abstract': False, 'is_breakdown': False, 'dimension_member_label': None,
        '2018-07-01 (Q3)': [20_830_000_000, 0.003],
    })
    assert fg._match_revenue_row(df, '2018-07-01 (Q3)') == (0, False)
