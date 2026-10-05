import pandas as pd
from types import SimpleNamespace
from period_identity import cover_focus, reported_label, build_period_map


def financials(end='2021-01-03',year=2020,quarter='FY'):
    df=pd.DataFrame({'concept':['dei_DocumentPeriodEndDate','dei_DocumentFiscalYearFocus','dei_DocumentFiscalPeriodFocus'],
                     end+' (FY)':[end,year,quarter]})
    return SimpleNamespace(cover=lambda:SimpleNamespace(to_dataframe=lambda:df))


def test_reported_annual_year_is_not_calendar_end_year():
    assert reported_label(cover_focus(financials()),'2021-01-03 (FY)',annual=True)=='FY2020'


def test_sixteen_week_q1_uses_reported_quarter():
    assert reported_label(cover_focus(financials('2010-05-22',2010,'Q1')),'2010-05-22 (YTD)')=='FY2010Q1'


def test_comparative_column_never_borrows_document_focus():
    assert reported_label(cover_focus(financials()),'2020-01-05 (FY)',annual=True) is None


def test_conflicting_cover_focus_is_rejected():
    fin=financials()
    df=fin.cover().to_dataframe()
    df['other']=df.iloc[:,1]
    # Non-period columns cannot supply focus values.
    df.iloc[:,1]=None
    fin.cover=lambda:SimpleNamespace(to_dataframe=lambda:df)
    assert cover_focus(fin) is None


def test_complete_reported_year_recovers_missing_sixteen_week_quarter_focus():
    records=[('2009-01-31','10-K',('2009-01-31',2008,'FY')),
             ('2010-01-30','10-K',('2010-01-30',2009,'FY')),
             ('2009-05-23','10-Q',None),('2009-08-15','10-Q',None),('2009-11-07','10-Q',None)]
    result=build_period_map(records)
    assert result[('2009-05-23',False)]=='FY2009Q1'
    assert result[('2010-01-30',False)]=='FY2009Q4'


def test_missing_quarter_is_not_renumbered_as_q1():
    records=[('2009-01-31','10-K',('2009-01-31',2008,'FY')),
             ('2010-01-30','10-K',('2010-01-30',2009,'FY')),
             ('2009-08-15','10-Q',None),('2009-11-07','10-Q',None)]
    assert ('2009-08-15',False) not in build_period_map(records)


def test_complete_anchored_sequence_rejects_stale_quarter_dei():
    records=[('2020-12-31','10-K',('2020-12-31',2020,'FY')),
             ('2021-12-31','10-K',('2021-12-31',2021,'FY')),
             ('2021-03-31','10-Q',('2021-03-31',2020,'Q3')),
             ('2021-06-30','10-Q',None),('2021-09-30','10-Q',None)]
    assert build_period_map(records)[('2021-03-31',False)]=='FY2021Q1'


def test_conflicting_annual_chain_revokes_direct_labels_before_deduplication():
    records=[(f'{y}-01-31','10-K',(f'{y}-01-31',focus,'FY')) for y,focus in [(2024,2024),(2025,2025),(2026,2025)]]
    records.append(('2026-04-30','10-Q',('2026-04-30',2026,'Q1')))
    assert not build_period_map(records)


def test_financial_six_k_gets_same_complete_chain_protection_as_ten_q():
    records=[('2024-03-31','20-F',('2024-03-31',2024,'FY')),
             ('2025-03-31','20-F',('2025-03-31',2025,'FY')),
             ('2024-06-30','6-K',('2024-06-30',2024,'Q1')),
             ('2024-09-30','6-K',('2024-09-30',2024,'Q2')),
             ('2024-12-31','6-K',('2024-12-31',2024,'Q3'))]
    assert build_period_map(records)[('2024-06-30',False)]=='FY2025Q1'


def test_partial_year_rejects_quarter_focus_for_already_ended_fiscal_year():
    records=[('2024-03-31','20-F',('2024-03-31',2024,'FY')),
             ('2025-03-31','20-F',('2025-03-31',2025,'FY')),
             ('2025-06-30','6-K',('2025-06-30',2025,'Q1'))]
    assert ('2025-06-30',False) not in build_period_map(records)


def test_verified_source_correction_is_accession_and_value_guarded():
    from period_identity import corrected_focus
    focus=('2024-02-03',2024,'FY')
    assert corrected_focus('0001558370-24-004603', focus)==('2024-02-03',2023,'FY')
    assert corrected_focus('unrelated-accession', focus)==focus
    assert corrected_focus('0001558370-24-004603',('2024-02-03',2023,'FY'))==('2024-02-03',2023,'FY')
    assert corrected_focus('0001558370-24-004603',('2024-02-04',2024,'FY'))==('2024-02-04',2024,'FY')
    assert corrected_focus('0001558370-25-004267',('2025-02-01',2025,'FY'))==('2025-02-01',2024,'FY')
