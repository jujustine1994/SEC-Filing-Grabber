import pandas as pd
import pytest
import fetcher_gaap as fg


def group(values):
    return pd.DataFrame({
        'concept': ['us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax',
                    'us-gaap_RevenueFromContractWithCustomerIncludingAssessedTax',
                    'us-gaap_RevenueFromRelatedParties', 'company_RevenuesAndOtherIncome'],
        'label': ['Sales and other operating revenues', 'Sales and other operating revenues',
                  'Related party sales', 'Total revenues and other income'],
        'standard_concept': ['Revenue', 'Revenue', 'Revenue', None],
        'parent_concept': ['company_RevenuesAndOtherIncome'] * 3 + ['us-gaap_OperatingIncomeLoss'],
        'weight': [1, 1, 1, 1], 'abstract': False, 'is_breakdown': False,
        'dimension_member_label': None, '2020-12-31 (FY)': values,
    })


def test_tax_aliases_are_alternatives_not_additive_components():
    df = group([100, None, 5, 120])
    indices, ambiguous = fg._revenue_components(df, '2020-12-31 (FY)')
    assert not ambiguous
    assert sum(df.loc[i, '2020-12-31 (FY)'] for i in indices) == 105


def test_current_tax_aliases_with_different_values_remain_ambiguous():
    df = group([100, 110, 5, 120])
    assert fg._match_revenue_row(df, '2020-12-31 (FY)') == (None, True)


def test_no_current_tax_alias_cannot_be_replaced_by_related_party_sales():
    df = group([None, None, 5, 120])
    assert fg._match_revenue_row(df, '2020-12-31 (FY)') == (None, False)


@pytest.mark.parametrize('past_fact', [None, 10])
def test_only_fully_empty_unlinked_placeholder_is_ignored(past_fact):
    df = group([100, None, 5, 120])
    extra = df.iloc[[0]].copy()
    extra.index = [100]
    extra['concept'] = 'us-gaap_FinancialServicesRevenue'
    extra['2020-12-31 (FY)'] = None
    extra['2019-12-31 (FY)'] = past_fact
    extra['weight'] = None
    df = pd.concat([df, extra])
    indices, ambiguous = fg._revenue_components(df, '2020-12-31 (FY)')
    if past_fact is None:
        assert not ambiguous and indices
    else:
        assert ambiguous and not indices
