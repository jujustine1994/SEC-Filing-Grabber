"""Revenue competitors must not cross CamelCase word boundaries."""
import pandas as pd
import pytest
import fetcher_gaap as fg


@pytest.mark.parametrize('concept', [
    'us-gaap_OtherComprehensiveIncomeAvailableforsaleSecuritiesAdjustmentNetOfTaxPortionAttributableToParent',
    'us-gaap_OtherComprehensiveIncomeAvailableForSaleSecuritiesAdjustmentNetOfTax',
])
def test_available_for_sale_securities_cannot_compete_with_revenue(concept):
    df = pd.DataFrame({
        'concept': ['us-gaap_FoodAndBeverageRevenue', concept],
        'label': ['Revenue', 'Unrealized gain (loss) on investments'],
        'standard_concept': ['Revenue', 'OtherComprehensiveIncome'],
        'abstract': False, 'is_breakdown': False, 'dimension_member_label': None,
        '2016-12-31 (FY)': [3_904_384_000, 1_402_000],
    })
    before = df.copy(deep=True)
    assert fg._match_revenue_row(df, '2016-12-31 (FY)') == (0, False)
    pd.testing.assert_frame_equal(df, before)


@pytest.mark.parametrize('concept,label', [
    ('company_ProductSales', 'Products'),
    ('company_ManagementFeesRevenue', 'Management fees'),
    ('company_OTHER_REVENUES', 'Other'),
])
def test_real_revenue_word_remains_a_competitor(concept, label):
    df = pd.DataFrame({
        'concept': ['company_CustomRevenue', concept],
        'label': ['Revenue', label],
        'abstract': False, 'is_breakdown': False, 'dimension_member_label': None,
        '2020-12-31 (FY)': [100, 10],
    })
    assert fg._match_revenue_row(df, '2020-12-31 (FY)') == (None, True)
