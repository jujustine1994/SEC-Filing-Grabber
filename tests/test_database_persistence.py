import pytest

import local_db
from database import DatabaseError
from fetch_ledger import FetchLedger


def test_meta_write_failure_is_reported_per_company(monkeypatch):
    monkeypatch.setattr(local_db,'write_meta',lambda *a: (_ for _ in ()).throw(DatabaseError('full')))
    report = local_db.update_local_db(['NVDA','AAPL'], 'test',
        list_filings=lambda *a: ({'10-Q': [], '10-K': []},1045810),
        fetch=lambda *a: FetchLedger())
    assert report.failed == 2


def test_fetch_persistence_failure_does_not_report_updated(monkeypatch):
    ledger = FetchLedger(persistence_errors=['accession'])
    report = local_db.update_local_db(['NVDA'], 'test',
        list_filings=lambda *a: ({'10-Q': [('0001045810-25-000001','2025-01-01')], '10-K': []},1045810),
        fetch=lambda *a: ledger)
    assert report.failed == 1
    assert report.updated == 0


def test_persistence_failures_are_visible_in_report_summary():
    assert FetchLedger(persistence_errors=['accession']).summary()


def test_retry_persistence_failure_survives_network_recovery():
    from fetch_ledger import Gap
    ledger = FetchLedger(gaps=[Gap('accession', 'network', 'Timeout')])
    ledger.absorbed_by_retry(FetchLedger(persistence_errors=['accession']))
    assert not ledger.has_gaps
    assert ledger.persistence_errors == ['accession']


def test_persistence_only_failure_counts_as_warning():
    ledger = FetchLedger(persistence_errors=['accession'])
    assert ledger.has_warnings
    assert not ledger.has_gaps
