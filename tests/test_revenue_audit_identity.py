from copy import deepcopy
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from compare_revenue_audits import compare
from verify_fixed_financial_rows import source_contract


def pair(display):
    contract = source_contract()
    names = ['Fiscal Quarter', 'Calendar Quarter', 'Period End', '', 'Income Statement']
    for template, following in [('IS_TEMPLATE', 'Balance Sheet'), ('BS_TEMPLATE', 'Cash Flow')]:
        names += contract[template] + [''] * contract['SECTION_GAP'] + [following]
    names += contract['CF_TEMPLATE'] + [''] * contract['SECTION_GAP'] + ['Other (as reported)', display]
    labels = ['Reported item'] * len(names)
    labels[-1] = 'company_SourceComponent'
    values = [[100.0 + i] if name in sum((contract[k] for k in ['IS_TEMPLATE', 'BS_TEMPLATE', 'CF_TEMPLATE']), []) else [None]
              for i, name in enumerate(names)]
    values[-1] = [987.0]
    table = dict(sheet_name='Data_Financials(Q)', concepts=names, labels=labels,
        values=values, quarter_labels=['FY2024Q1'], period_ends=['2024-03-31'], filing_dates=['2024-05-01'])
    new = dict(tables=[table])
    old = deepcopy(new)
    old_table = old['tables'][0]
    # Reproduce the old merger promoting an IS component after fixed IS slots.
    position = 5 + len(contract['IS_TEMPLATE'])
    for key in ['concepts', 'labels', 'values']:
        old_table[key].insert(position, old_table[key].pop())
    return old, new


def test_moving_same_named_source_does_not_change_cf_net_income():
    old, new = pair('Net Income')
    assert compare(old, new)['value_changes'] == []


def test_cf_change_is_distinct_from_is_and_source_net_income():
    old, new = pair('Net Income')
    table = new['tables'][0]
    index = table['concepts'].index('Cash Flow') + 1
    table['values'][index][0] += 1
    changes = compare(old, new)['value_changes']
    assert len(changes) == 1
    assert changes[0]['after'] - changes[0]['before'] == 1


def test_same_named_source_change_keeps_its_source_identity():
    old, new = pair('Net Income')
    new['tables'][0]['values'][-1][0] += 2
    changes = compare(old, new)['value_changes']
    assert len(changes) == 1
    assert changes[0]['after'] - changes[0]['before'] == 2
