import hashlib
import json
import subprocess
import sys
from pathlib import Path

import pytest

from database import DatabaseError, create_database
from database_io import commit_filing, database_lock, clean_staging


ACC = '0001045810-25-000001'


def entry(value):
    return dict(accession_no=ACC, cik=1045810, schema_version=2, form='10-Q',
                has_financials=True, dataframes={'income_statement': {'data': value}},
                edgartools_version='test', fetched_keys=['income_statement'])


def test_history_keeps_exact_previous_bytes(isolated_database):
    root = isolated_database
    commit_filing(root, 'NVDA', ACC, entry(1))
    path = root / 'filings/NVDA' / (ACC + '.json')
    old = path.read_bytes()
    commit_filing(root, 'NVDA', ACC, entry(2))
    history = root / 'history/NVDA' / ACC / (hashlib.sha256(old).hexdigest() + '.json')
    assert history.read_bytes() == old
    assert json.loads(path.read_bytes())['dataframes']['income_statement']['data'] == 2


def test_corrupt_history_prevents_overwrite(isolated_database):
    root = isolated_database
    commit_filing(root, 'NVDA', ACC, entry(1))
    path = root / 'filings/NVDA' / (ACC + '.json')
    old = path.read_bytes()
    history = root / 'history/NVDA' / ACC / (hashlib.sha256(old).hexdigest() + '.json')
    history.parent.mkdir(parents=True)
    history.write_bytes(b'corrupt')
    with pytest.raises(DatabaseError):
        commit_filing(root, 'NVDA', ACC, entry(2))
    assert path.read_bytes() == old


def test_archive_failure_keeps_current(isolated_database, monkeypatch):
    import database_io
    root = isolated_database
    commit_filing(root, 'NVDA', ACC, entry(1))
    path = root / 'filings/NVDA' / (ACC + '.json')
    old = path.read_bytes()
    monkeypatch.setattr(database_io, '_preserve', lambda *a: (_ for _ in ()).throw(OSError('full')))
    with pytest.raises(DatabaseError):
        commit_filing(root, 'NVDA', ACC, entry(2))
    assert path.read_bytes() == old


def test_corrupt_current_is_preserved_before_repair(isolated_database):
    root = isolated_database
    path = root / 'filings/NVDA' / (ACC + '.json')
    path.parent.mkdir()
    path.write_bytes(b'{truncated')
    commit_filing(root, 'NVDA', ACC, entry(1))
    assert next((root / 'history/NVDA' / ACC).iterdir()).read_bytes() == b'{truncated'


def test_duplicate_content_does_not_grow_history(isolated_database):
    root = isolated_database
    commit_filing(root, 'NVDA', ACC, dict(entry(1), cached_at='a'))
    commit_filing(root, 'NVDA', ACC, dict(entry(1), cached_at='b'))
    assert not list((root / 'history').rglob('*.json'))


def test_formal_clear_apis_always_reject(isolated_database):
    import filing_cache
    for call in (lambda: filing_cache.clear_all(), lambda: filing_cache.clear_ticker('NVDA')):
        with pytest.raises(DatabaseError):
            call()
    assert (isolated_database / 'database.json').exists()


def test_cleanup_only_staging(isolated_database):
    import os, time
    root = isolated_database
    formal = root / 'filings/keep.tmp'
    formal.write_bytes(b'keep')
    tmp = root / 'staging/old.tmp'
    tmp.write_bytes(b'partial')
    os.utime(tmp, (time.time()-7200,)*2)
    assert clean_staging(root) == 1
    assert formal.read_bytes() == b'keep'


def test_traversal_is_rejected(isolated_database):
    with pytest.raises(DatabaseError):
        commit_filing(isolated_database, '../escape', ACC, entry(1))


def test_two_processes_preserve_both_versions(isolated_database):
    root = isolated_database
    src = str(Path(__file__).parents[1] / 'src')
    code = "import sys,json;sys.path.insert(0,sys.argv[1]);from database_io import commit_filing;from pathlib import Path;commit_filing(Path(sys.argv[2]),'NVDA',sys.argv[3],json.loads(sys.argv[4]))"
    procs = [subprocess.Popen([sys.executable, '-c', code, src, str(root), ACC, json.dumps(entry(n))]) for n in (1,2)]
    assert [p.wait(timeout=40) for p in procs] == [0,0]
    values = [json.loads(p.read_bytes())['dataframes']['income_statement']['data']
              for p in [root / 'filings/NVDA' / (ACC+'.json'), *list((root/'history/NVDA'/ACC).glob('*.json'))]]
    assert sorted(values) == [1,2]
