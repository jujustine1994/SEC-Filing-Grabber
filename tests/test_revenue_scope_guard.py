"""A goods/services net subtotal does not authorize the entire company's sales."""
import pandas as pd
import fetcher_gaap as fg


def test_goods_subtotal_with_independent_custom_services_is_not_total():
    df = pd.DataFrame({
        'concept': ['us-gaap_SalesRevenueGoodsNet', 'company_ServicesRevenue'],
        'label': ['Products', 'Services'],
        'standard_concept': ['Revenue', 'Revenue'],
        'parent_concept': ['us-gaap_GrossProfit', 'us-gaap_GrossProfit'],
        'abstract': False, 'is_breakdown': False, 'dimension_member_label': None,
        '2020-12-31 (FY)': [100, 20],
    })
    assert fg._match_revenue_row(df, '2020-12-31 (FY)') == (None, True)
