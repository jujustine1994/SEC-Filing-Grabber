import hashlib
import json

import pytest
from quarterly_source import read_quarter_source_pack, supplement_quarter_table


HTML = '''<h2>CONDENSED CONSOLIDATED STATEMENTS OF OPERATIONS</h2>
<p>(Millions except per share amounts) (Unaudited)</p>
<table><tr><td></td><td colspan="2">Three Months Ended</td><td>Year Ended</td></tr>
<tr><td></td><td>December 28, 2030</td><td>September 28, 2030</td><td>December 28, 2030</td></tr>
<tr><td>Revenue</td><td>100</td><td>90</td><td>360</td></tr>
<tr><td>Gross profit</td><td>60</td><td>50</td><td>220</td></tr>
<tr><td>Operating income</td><td>30</td><td>20</td><td>100</td></tr>
<tr><td>Income before income taxes</td><td>25</td><td>18</td><td>90</td></tr>
<tr><td>Income tax provision</td><td>5</td><td>3</td><td>20</td></tr>
<tr><td>Net income</td><td>20</td><td>15</td><td>70</td></tr>
<tr><td>Earnings per share:</td></tr>
<tr><td>Basic</td><td>0.20</td><td>0.15</td><td>0.70</td></tr>
<tr><td>Diluted</td><td>0.19</td><td>0.14</td><td>0.67</td></tr>
<tr><td>Shares used in per share calculation</td></tr>
<tr><td>Basic</td><td>100</td><td>99</td><td>98</td></tr>
<tr><td>Diluted</td><td>105</td><td>104</td><td>103</td></tr></table>'''


def pack(tmp_path, content=HTML):
    body = content.encode('utf-8')
    (tmp_path / 'quarter.htm').write_bytes(body)
    manifest = dict(cik=1234, sources=[dict(filename='quarter.htm', sha256=hashlib.sha256(body).hexdigest(),
        document_type='EX-99.1', form='8-K', accession='0000001234-31-000001',
        url='https://www.sec.gov/Archives/edgar/data/1234/000000123431000001/quarter.htm',
        filing_date='2031-01-20')])
    (tmp_path / 'manifest.json').write_text(json.dumps(manifest), encoding='utf-8')
    return tmp_path


def test_reads_standalone_quarters_in_base_units_and_keeps_provenance(tmp_path):
    records = read_quarter_source_pack(pack(tmp_path), 1234)
    q4 = records['2030-12-28']
    assert q4['values']['Basic Shares'] == 100_000_000
    assert q4['values']['Basic EPS'] == 0.2
    assert q4['values']['Revenue'] == 100_000_000
    assert q4['table'] == 1 and q4['column'] == 1
    assert q4['sha256'] and q4['accession']
    assert records['2030-09-28']['values']['Basic Shares'] == 99_000_000


@pytest.mark.parametrize('change', ['issuer', 'hash', 'url', 'path'])
def test_rejects_invalid_source_identity_or_bytes(tmp_path, change):
    folder = pack(tmp_path)
    manifest = json.loads((folder / 'manifest.json').read_text())
    if change == 'issuer': manifest['cik'] = 9999
    if change == 'hash': manifest['sources'][0]['sha256'] = '0' * 64
    if change == 'url': manifest['sources'][0]['url'] = 'https://example.com/not-sec.htm'
    if change == 'path': manifest['sources'][0]['filename'] = '../quarter.htm'
    (folder / 'manifest.json').write_text(json.dumps(manifest))
    with pytest.raises(ValueError): read_quarter_source_pack(folder, 1234)


@pytest.mark.parametrize('content', [HTML.replace('Three Months Ended', 'Nine Months Ended'),
    HTML.replace('Millions except per share amounts', 'Amounts'),
    HTML.replace('STATEMENTS OF OPERATIONS', 'NON-GAAP STATEMENTS OF OPERATIONS'),
    HTML.replace('<td>100</td><td>99</td>', '<td></td><td>99</td>')])
def test_unsupported_scope_units_or_missing_shares_never_produce_a_value(tmp_path, content):
    assert not read_quarter_source_pack(pack(tmp_path, content), 1234)


def test_supplement_fills_missing_without_mutating_source_or_reported_values(tmp_path):
    from fetcher_gaap import StatementTable
    records = read_quarter_source_pack(pack(tmp_path), 1234)
    table = StatementTable('Data_Financials(Q)', ['FY2030Q4'], ['2031-01-20'],
        ['Revenue', 'Gross Profit', 'Operating Income', 'Pre-tax Income', 'Income Tax', 'Net Income', 'Basic Shares', 'Basic EPS'],
        [[100_000_000], [60_000_000], [30_000_000], [25_000_000], [5_000_000], [20_000_000], [None], [None]],
        period_ends=['2030-12-28'])
    result, trace = supplement_quarter_table(table, records)
    assert result.values[-2:] == [[100_000_000], [0.2]]
    assert table.values[-2:] == [[None], [None]]
    assert [r['metric'] for r in trace] == ['Basic EPS', 'Basic Shares']
    assert all(r['raw_label'] == 'Basic' for r in trace)


def test_scope_conflict_prevents_all_supplementation(tmp_path):
    from fetcher_gaap import StatementTable
    records = read_quarter_source_pack(pack(tmp_path), 1234)
    table = StatementTable('Data_Financials(Q)', ['FY2030Q4'], ['2031-01-20'],
        ['Revenue', 'Basic Shares'], [[99_000_000], [None]], period_ends=['2030-12-28'])
    result, trace = supplement_quarter_table(table, records)
    assert result.values == table.values
    assert trace[0]['status'] == 'scope-conflict'


def test_conflicting_same_filed_sources_are_rejected(tmp_path):
    folder = pack(tmp_path)
    manifest = json.loads((folder / 'manifest.json').read_text())
    body = HTML.replace('<td>0.20</td>', '<td>0.21</td>').encode()
    (folder / 'second.htm').write_bytes(body)
    manifest['sources'].append({**manifest['sources'][0], 'filename': 'second.htm',
                               'sha256': hashlib.sha256(body).hexdigest()})
    (folder / 'manifest.json').write_text(json.dumps(manifest))
    with pytest.raises(ValueError, match='Conflicting'):
        read_quarter_source_pack(folder, 1234)


def test_year_first_layout_uses_aligned_quarter_group_not_repeated_date(tmp_path):
    from lxml import html
    doc = html.fromstring(HTML)
    rows = doc.xpath('//tr')
    rows[0][1].text = 'Year Ended'
    rows[0][1].set('colspan', '1')
    rows[0][2].text = 'Three Months Ended'
    rows[0][2].set('colspan', '2')
    for row in rows[1:]:
        if len(row) == 4:
            last = row[-1]
            row.remove(last)
            row.insert(1, last)
    records = read_quarter_source_pack(pack(tmp_path, html.tostring(doc, encoding='unicode')), 1234)
    assert records['2030-12-28']['values']['Basic EPS'] == 0.2
    assert records['2030-12-28']['values']['Basic Shares'] == 100_000_000


def test_mixed_share_units_are_rejected_without_scaling_guess(tmp_path):
    content = HTML.replace('Millions except per share amounts',
                           'Dollars in millions, shares in thousands, except per share amounts')
    assert not read_quarter_source_pack(pack(tmp_path, content), 1234)


def test_continuing_eps_section_is_not_total_eps(tmp_path):
    content = HTML.replace('Earnings per share:', 'Earnings per share — continuing operations:')
    assert not read_quarter_source_pack(pack(tmp_path, content), 1234)


def test_no_existing_anchors_cannot_authorize_official_supplement(tmp_path):
    from fetcher_gaap import StatementTable
    records = read_quarter_source_pack(pack(tmp_path), 1234)
    table = StatementTable('Data_Financials(Q)', ['FY2030Q4'], ['2031-01-20'],
        ['Basic Shares'], [[None]], period_ends=['2030-12-28'])
    result, trace = supplement_quarter_table(table, records)
    assert result.values == [[None]]
    assert trace[0]['status'] == 'scope-conflict'


def test_eps_continuing_subheading_does_not_supply_total_eps(tmp_path):
    content = HTML.replace('<tr><td>Basic</td><td>0.20',
        '<tr><td>Continuing operations:</td></tr><tr><td>Basic</td><td>0.20')
    assert not read_quarter_source_pack(pack(tmp_path, content), 1234)


@pytest.mark.parametrize('label', ['Equity loss', 'Equity loss:', ' (Equity loss) '])
def test_explicit_equity_loss_is_signed_loss_with_original_value_retained(tmp_path, label):
    content = HTML.replace('<tr><td>Net income</td>',
        '<tr><td>' + label + '</td><td>6</td><td>4</td><td>10</td></tr><tr><td>Net income</td>')
    record = read_quarter_source_pack(pack(tmp_path, content), 1234)['2030-12-28']
    assert record['values']['Equity Method Income/(Loss)'] == -6_000_000
    assert record['reported_values']['Equity Method Income/(Loss)'] == 6


@pytest.mark.parametrize('label', ['Gain on sale of discontinued operations, net of tax',
    'Income from discontinued operations, net of tax attributable to noncontrolling interests',
    'Income from continuing operations, net of tax attributable to parent'])
def test_bridge_component_and_attribution_labels_are_not_full_totals(tmp_path, label):
    content = HTML.replace('<tr><td>Net income</td>',
        '<tr><td>' + label + '</td><td>6</td><td>4</td><td>10</td></tr><tr><td>Net income</td>')
    record = read_quarter_source_pack(pack(tmp_path, content), 1234)['2030-12-28']
    assert 'Discontinued Operations, net of tax' not in record['values']
    assert 'Net Income from Continuing Operations' not in record['values']
