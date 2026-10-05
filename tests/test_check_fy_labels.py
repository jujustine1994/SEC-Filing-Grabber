import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import check_fy_labels as checker
from fiscal_audit import label_anomalies


def test_stable_fiscal_month_does_not_exempt_company_from_verification(monkeypatch):
    calls = []
    monkeypatch.setattr(checker, 'scan', lambda ticker: dict(ticker=ticker, is_week_based=False))
    def verify(ticker, identity=None):
        calls.append(ticker)
        return label_anomalies(['2010-05-22'], ['2010-05-22'])
    monkeypatch.setattr(checker, 'verify', verify)
    assert checker.main(['--verify', 'KR']) != 0
    assert calls == ['KR']


def test_failed_verification_never_exits_successfully(monkeypatch):
    monkeypatch.setattr(checker, 'scan', lambda ticker: dict(ticker=ticker, is_week_based=True))
    def failed(*args):
        raise ValueError('invalid cached input')
    monkeypatch.setattr(checker, 'verify', failed)
    assert checker.main(['--verify', 'KR']) != 0


def test_missing_official_metadata_is_inconclusive_even_with_consistent_labels(monkeypatch,capsys):
    monkeypatch.setattr(checker,'scan',lambda ticker:{})
    result=label_anomalies(['FY2025Q1','FY2025Q2','FY2025Q3','FY2025Q4'],['2025-03-31','2025-06-30','2025-09-30','2025-12-31'])
    result['metadata_complete']=False
    monkeypatch.setattr(checker,'verify',lambda ticker:result)
    assert checker.main(['--verify','TEST'])==0
    assert 'INCONCLUSIVE' in capsys.readouterr().out


def test_no_observed_periods_never_counts_as_consistent(monkeypatch,capsys):
    monkeypatch.setattr(checker,'scan',lambda ticker:{})
    result=label_anomalies([],[])
    result.update(metadata_complete=True,observed_periods=0)
    monkeypatch.setattr(checker,'verify',lambda ticker:result)
    assert checker.main(['--verify','TEST'])==0
    assert 'INCONCLUSIVE' in capsys.readouterr().out
