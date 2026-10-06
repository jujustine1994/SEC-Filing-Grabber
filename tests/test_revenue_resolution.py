"""Source-certified Revenue cases and guards for mechanical derivations."""
import json
from pathlib import Path
from unittest.mock import MagicMock

import pandas as pd
import pytest
import fetcher_gaap as fg

CASES = json.loads((Path(__file__).parent / 'fixtures/revenue-resolution-cases.json').read_text(encoding='utf-8'))


def dataframe(case):
    rows = [{**row, case['column']: row['value']} for row in case['rows']]
    df = pd.DataFrame(rows).set_index('index')
    df['abstract'] = False
    df['is_breakdown'] = False
    df['dimension_member_label'] = None
    return df


@pytest.mark.parametrize('case', CASES, ids=lambda c: c['ticker'] + '-' + c['accession'])
def test_source_certified_total_or_explicit_derivation(case):
    df = dataframe(case)
    before = df.copy(deep=True)
    index, ambiguous = fg._match_revenue_row(df, case['column'])
    assert not ambiguous
    if index is not None:
        value = df.loc[index, case['column']]
    else:
        components, ambiguous = fg._revenue_components(df, case['column'])
        assert not ambiguous and components
        value = sum(df.loc[i, case['column']] for i in components)
    assert value == case['expected_value']
    pd.testing.assert_frame_equal(df, before)


def test_bank_formula_is_before_credit_loss_provision_and_preserves_components():
    case = next(c for c in CASES if c['ticker']=='BK' and c['accession']=='0001193125-10-042948')
    df = dataframe(case)
    filing = MagicMock()
    filing.obj.return_value.financials.income_statement.return_value.to_dataframe.return_value = df
    filing.obj.return_value.financials.cashflow_statement.return_value = None
    table, _ = fg._build_is_table([filing], max_filings=1)
    assert table.values[table.concepts.index('Revenue')] == [7_687_000_000]
    assert any('derived' in label.casefold() for label in table.labels)
    assert 'us-gaap_NoninterestIncome' in table.labels
    assert 'us-gaap_InterestIncomeExpenseNet' in table.labels


@pytest.mark.parametrize('ticker', ['BK', 'MPC'])
def test_missing_component_cannot_be_replaced_with_zero(ticker):
    case = next(c for c in CASES if c['ticker']==ticker)
    df = dataframe(case)
    component = case['facts'][0]['concept']
    df.loc[df['concept']==component, case['column']] = None
    indices, ambiguous = fg._revenue_components(df, case['column'])
    assert not indices


def test_duplicate_bank_components_with_different_values_are_ambiguous():
    case = next(c for c in CASES if c['ticker']=='BK')
    df = dataframe(case)
    duplicate = df[df['concept']=='us-gaap_NoninterestIncome'].copy()
    duplicate.index = [1000]
    duplicate[case['column']] = 1
    df = pd.concat([df, duplicate])
    assert fg._revenue_components(df, case['column']) == ([], True)


def test_explicit_reported_total_with_missing_value_blocks_bank_derivation():
    case = next(c for c in CASES if c['ticker']=='MS')
    df = dataframe(case)
    df.loc[df['concept']=='us-gaap_RevenuesNetOfInterestExpense', case['column']] = None
    index, ambiguous = fg._match_revenue_row(df, case['column'])
    assert index is not None and not ambiguous
    assert pd.isna(df.loc[index, case['column']])
