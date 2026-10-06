from dataclasses import replace

import pytest

from fetcher_gaap import StatementTable, IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE, _merge_financials


def statements():
    return [StatementTable(sheet_name=name, quarter_labels=['FY2024Q1'],
        filing_dates=['2024-05-01'], period_ends=['2024-03-31'],
        concepts=[row[0] for row in template],
        labels=['Original source label'] * len(template),
        values=[[100.0 + i] for i in range(len(template))])
        for name, template in [('Data_IS', IS_TEMPLATE), ('Data_BS', BS_TEMPLATE), ('Data_CF', CF_TEMPLATE)]]


@pytest.mark.parametrize('statement_index', [0, 1, 2])
def test_same_named_overflow_cannot_move_fixed_statement_rows(statement_index):
    tables = statements()
    baseline = _merge_financials(*tables)
    original = tables[statement_index]
    tables[statement_index] = replace(original,
        concepts=original.concepts + [original.concepts[0]],
        labels=original.labels + ['company_SourceComponent'],
        values=original.values + [[123456.0]])
    actual = _merge_financials(*tables)
    count = len(baseline.concepts)
    assert actual.concepts[:count] == baseline.concepts
    assert actual.values[:count] == baseline.values
    assert actual.concepts.index('Balance Sheet') == baseline.concepts.index('Balance Sheet')
    assert actual.concepts.index('Cash Flow') == baseline.concepts.index('Cash Flow')
    source_row = actual.labels.index('company_SourceComponent')
    assert source_row > actual.concepts.index('Other (as reported)')
    assert actual.values[source_row] == [123456.0]
    assert len(original.concepts) == len([IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE][statement_index])


def test_nongaap_same_named_revenue_has_no_fixed_template_identity():
    tables = [StatementTable(sheet_name=name, quarter_labels=['FY2024Q1'],
        filing_dates=['2024-05-01'], period_ends=['2024-03-31'],
        concepts=[], labels=[], values=[])
        for name in ['Data_IS_NG', 'Data_BS_NG', 'Data_CF_NG']]
    tables[0] = replace(tables[0], concepts=['Revenue'],
        labels=['company_AdjustedRevenue'], values=[[321.0]])
    actual = _merge_financials(*tables, sheet_name='Data_Financials_NG(Q)')
    row = actual.concepts.index('Revenue')
    assert row > actual.concepts.index('Other (as reported)')
    assert actual.values[row] == [321.0]
