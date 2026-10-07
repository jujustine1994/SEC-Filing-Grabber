from types import SimpleNamespace
from unittest.mock import MagicMock
from datetime import date
import pandas as pd
import pytest
import fetcher_gaap as fg
import main


def filing(acc,end,form,year,period):
    frame=pd.DataFrame({'concept':['dei_DocumentPeriodEndDate','dei_DocumentFiscalYearFocus','dei_DocumentFiscalPeriodFocus'],end:[end,str(year),period]})
    obj=SimpleNamespace(financials=SimpleNamespace(cover=lambda:SimpleNamespace(to_dataframe=lambda:frame)))
    return SimpleNamespace(accession_no=acc,form=form,period_of_report=end,report_date=end,filing_date=date(2026,1,1),obj=MagicMock(return_value=obj)),obj


def test_preview_uses_the_same_named_fiscal_year_as_pipeline_without_downloading_history(monkeypatch):
    q,qobj=filing('0000000001-25-000001','2025-05-03','10-Q',2025,'Q1')
    k,kobj=filing('0000000001-25-000002','2025-02-01','10-K',2024,'FY')
    company=SimpleNamespace(cik=1,fiscal_year_end='0131')
    monkeypatch.setattr(fg,'Company',lambda _:company)
    monkeypatch.setattr(fg,'set_identity',lambda _:None)
    monkeypatch.setattr(fg,'_list_filings',lambda c,f:[q] if f=='10-Q' else [k])
    monkeypatch.setattr(fg,'_build_segment_tables',lambda *a,**k:[])
    monkeypatch.setattr(fg.filing_cache,'load_filing',lambda t,a,c,**kw:{'object':qobj if a==q.accession_no else kobj})
    monkeypatch.setattr(fg.filing_cache,'cached_filing',lambda entry:entry['object'])
    result=fg.preview_sheets('RETAIL','identity')
    assert result['latest_label']=='FY2025Q1'
    assert not result['label_estimated']
    q.obj.assert_not_called();k.obj.assert_not_called()


def test_incomplete_preview_inputs_are_explicitly_estimated(monkeypatch):
    q,_=filing('0000000001-25-000001','2025-05-03','10-Q',2025,'Q1')
    company=SimpleNamespace(cik=1,fiscal_year_end='0131')
    monkeypatch.setattr(fg,'Company',lambda _:company)
    monkeypatch.setattr(fg,'set_identity',lambda _:None)
    monkeypatch.setattr(fg,'_list_filings',lambda c,f:[q] if f=='10-Q' else [])
    monkeypatch.setattr(fg,'_build_segment_tables',lambda *a,**k:[])
    monkeypatch.setattr(fg.filing_cache,'load_filing',lambda *a,**k:None)
    assert fg.preview_sheets('NEW','identity')['label_estimated']


def app(ticker):
    obj=main.SECFetcherApp.__new__(main.SECFetcherApp)
    obj.ticker_var=SimpleNamespace(get=lambda:ticker)
    obj._sheet_panel_frame=MagicMock()
    obj._SHEET_PANEL_TITLE_BASE='Sheets'
    obj._sheet_check_vars={'OLD':False}
    obj._sheet_panel_ticker='OLD'
    obj._scan_btn=MagicMock();obj._scan_hint_label=MagicMock()
    obj._scan_running=True;obj._build_sheet_panel=MagicMock()
    return obj


def test_old_ticker_scan_cannot_replace_current_company_panel():
    obj=app('NEW')
    obj._show_preview_result('OLD',dict(sheets=['WRONG'],latest_label='FY2025Q1',latest_period_end='2025-03-31',filing_date='2025-05-01'))
    obj._build_sheet_panel.assert_not_called()
    assert not obj._scan_running


def test_ticker_change_clears_old_sheet_exclusions():
    obj=app('NEW')
    obj._invalidate_sheet_preview()
    assert obj._sheet_check_vars=={}
    assert obj._sheet_panel_ticker is None
