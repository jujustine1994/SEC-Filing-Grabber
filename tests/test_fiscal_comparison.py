import sys
import json
import pytest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from compare_fiscal_audits import compare
import compare_fiscal_audits as cli


def result(labels, values, ends=None):
    return dict(tables=[dict(sheet_name='Data_Financials(Q)', quarter_labels=labels,
                            period_ends=ends or ['2021-01-03'], filing_dates=['2021-02-22'],
                            concepts=['Revenue'], values=[values])])


def test_date_key_distinguishes_label_correction_from_financial_change():
    change=compare(result(['FY2021Q4'],[100]), result(['FY2020Q4'],[100]))
    assert not change['value_changes']
    assert len(change['label_changes']) == 1


def test_blank_protection_is_a_real_change_not_equal_to_zero():
    change=compare(result(['FY2020Q4'],[100]), result(['FY2020Q4'],[None]))
    assert change['value_changes'][0]['before'] == 100
    assert change['value_changes'][0]['after'] is None
    assert compare(result(['FY2020Q4'],[None]),result(['FY2020Q4'],[0]))['value_changes']


def test_duplicate_concept_rows_are_never_overwritten():
    before=result(['FY2020Q4'],[100]);after=result(['FY2020Q4'],[100])
    for x,v in ((before,10),(after,20)):
        x['tables'][0]['concepts'].append('Revenue');x['tables'][0]['values'].append([v])
    assert len(compare(before,after)['value_changes']) == 1


def test_added_nonempty_period_is_reported():
    before=result(['FY2020Q4'],[100]);after=result(['FY2020Q4','FY2021Q1'],[100,120],['2021-01-03','2021-04-04'])
    after['tables'][0]['filing_dates'].append('2021-05-01')
    assert compare(before,after)['value_changes'][0]['after'] == 120


def test_same_filing_date_for_multiple_periods_uses_financial_end_provenance():
    source=result(['FY2018Q4','FY2019Q4'],[100,120],['2018-06-30','2019-06-30'])
    source['tables'][0]['filing_dates']=['2019-08-09']*2
    source['tables'].append(dict(sheet_name='Data_Meta',quarter_labels=['FY2018Q4','FY2019Q4'],
                                filing_dates=['2019-08-09']*2,period_ends=[],concepts=['Source'],values=[['a','b']]))
    assert not compare(source,source)['metadata_changes']


@pytest.mark.parametrize('mode',['errors','empty','stale'])
def test_cli_refuses_failed_or_empty_measurements(tmp_path,monkeypatch,mode):
    before=tmp_path/'before';after=tmp_path/'after'
    before.mkdir();after.mkdir()
    if mode=='stale':
        for folder in (before,after):
            (folder/'TEST.json').write_text(json.dumps(result(['FY2020Q4'],[100])),encoding='utf-8')
    if mode!='empty':
        for folder in (before,after):
            (folder/'TEST.error.json').write_text('{"error":"failed"}',encoding='utf-8')
    monkeypatch.setattr(sys,'argv',['compare',str(before),str(after),'--output',str(tmp_path/'comparison')])
    assert cli.main()!=0
