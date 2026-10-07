"""Tax exclusions in reported GAAP revenue are not adjusted performance."""
import pandas as pd
import pytest
from fetcher_gaap import _collect_overflow, _is_nongaap_label


@pytest.mark.parametrize('concept,label', [
    ('us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax',
     'Sales and other operating revenues, excluding consumer excise taxes'),
    ('us-gaap_SalesRevenueNet', 'Revenue excluding assessed tax'),
    ('us-gaap_Revenues', 'Revenue excluding sales-based taxes'),
])
def test_reported_gaap_tax_exclusion_stays_in_gaap_overflow(concept,label):
    gaap, ng = {}, {}
    df = pd.DataFrame([dict(concept=concept,label=label,abstract=False,
                           is_breakdown=False,dimension_member_label=None,
                           **{'2025-12-31 (FY)':123})])
    _collect_overflow(df,set(),'2025-12-31 (FY)','FY2025',gaap,ng)
    assert gaap[concept]['periods']=={'FY2025':123}
    assert not ng


@pytest.mark.parametrize('label', [
    'Adjusted revenue excluding assessed tax',
    'Revenue excluding assessed tax and excluding SBC',
    'Revenue excluding discontinued operations',
])
def test_tax_exclusion_does_not_hide_a_separate_adjustment(label):
    assert _is_nongaap_label(label,'us-gaap_Revenues')


def test_unknown_custom_tax_label_does_not_gain_gaap_authority():
    assert _is_nongaap_label('Revenue excluding assessed tax','company_CustomRevenue')
