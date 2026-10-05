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
