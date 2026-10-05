import json
from pathlib import Path

import pytest
from database import DatabaseError
from database_transfer import (inventory, migrate_database, create_snapshot,
                               verify_snapshot, restore_snapshot)


def legacy(tmp_path):
    source = tmp_path / 'legacy'
    (source / 'NVDA').mkdir(parents=True)
    (source / 'NVDA/0001045810-25-000001.json').write_bytes(b'{"cik":1045810}')
    (source / 'NVDA/_meta.json').write_bytes(b'{}')
    return source


def test_migration_keeps_every_byte_and_source(tmp_path):
    source = legacy(tmp_path)
    before = inventory(source)
    dest = tmp_path / 'independent'
    cfg = tmp_path / 'cfg.json'
    cfg.write_text('{"identity":"keep"}')
    report = migrate_database(source, dest, ['NVDA','NO-DATA'], config_path=cfg)
    assert report['verified']
    assert inventory(dest / 'filings') == before
    assert all((source / name).exists() for name in before)
    assert json.loads(cfg.read_text())['identity'] == 'keep'
    assert json.loads((dest / 'metadata/update_list.json').read_text())['tickers'] == ['NVDA','NO-DATA']


def test_unknown_nonempty_destination_not_overwritten(tmp_path):
    source = legacy(tmp_path)
    dest = tmp_path / 'other'
    dest.mkdir()
    (dest/'keep').write_bytes(b'keep')
    with pytest.raises(DatabaseError):
        migrate_database(source, dest, [], config_path=tmp_path/'cfg')
    assert (dest/'keep').read_bytes() == b'keep'


def test_changed_source_blocks_connection(tmp_path, monkeypatch):
    import database_transfer
    source = legacy(tmp_path)
    cfg = tmp_path / 'cfg.json'
    cfg.write_text('{}')
    original = database_transfer._copy_verified
    def change(*args):
        result = original(*args)
        (source/'NVDA/_meta.json').write_bytes(b'{"changed":1}')
        return result
    monkeypatch.setattr(database_transfer, '_copy_verified', change)
    with pytest.raises(DatabaseError):
        migrate_database(source, tmp_path/'dest', [], config_path=cfg)
    assert cfg.read_text() == '{}'


def test_snapshot_restores_to_new_folder(isolated_database, tmp_path):
    root = isolated_database
    (root/'filings/keep.json').write_bytes(b'{"keep":1}')
    snapshot = create_snapshot(root)
    assert verify_snapshot(snapshot)['verified']
    restored = tmp_path / 'restored'
    restore_snapshot(snapshot, restored)
    assert (restored/'filings/keep.json').read_bytes() == b'{"keep":1}'
    with pytest.raises(DatabaseError):
        restore_snapshot(snapshot, root)


def test_failed_destination_verification_never_publishes_identity(isolated_database, tmp_path, monkeypatch):
    import database_transfer
    (isolated_database/'filings/keep.json').write_bytes(b'{"keep":1}')
    snapshot = create_snapshot(isolated_database)
    destination = tmp_path/'restored'
    real = database_transfer._copy_verified
    def corrupt_earlier_copy(source, target, expected):
        real(source, target, expected)
        if target.parent.name == 'metadata':
            (destination/'filings/keep.json').write_bytes(b'corrupted')
    monkeypatch.setattr(database_transfer, '_copy_verified', corrupt_earlier_copy)
    with pytest.raises(DatabaseError, match='differs'):
        restore_snapshot(snapshot, destination)
    assert not (destination/'database.json').exists()
    assert verify_snapshot(snapshot)['verified']


def test_snapshot_corruption_and_escape_rejected(isolated_database, tmp_path):
    snapshot = create_snapshot(isolated_database)
    manifest = snapshot/'snapshot.json'
    data = json.loads(manifest.read_text())
    data['files']['../escape'] = {'size': 1, 'sha256': '0'*64}
    manifest.write_text(json.dumps(data))
    with pytest.raises(DatabaseError):
        restore_snapshot(snapshot, tmp_path/'restored')
    assert not (tmp_path/'escape').exists()


def test_migration_overlap_rejected(tmp_path):
    source = legacy(tmp_path)
    with pytest.raises(DatabaseError):
        migrate_database(source, source/'child', [], config_path=tmp_path/'cfg')


def test_repeated_completed_migration_keeps_new_update_list(tmp_path):
    source = legacy(tmp_path)
    dest = tmp_path/'dest'
    cfg = tmp_path/'cfg'
    migrate_database(source, dest, ['NVDA'], config_path=cfg)
    from local_db import write_update_list
    write_update_list(dest, ['NEW'])
    migrate_database(source, dest, ['NVDA'], config_path=cfg)
    from local_db import read_update_list
    assert read_update_list(dest) == ['NEW']


def test_partial_copy_can_resume_without_overwrite(tmp_path, monkeypatch):
    import database_transfer
    source = legacy(tmp_path)
    dest = tmp_path/'dest'
    cfg = tmp_path/'cfg'
    real = database_transfer._copy_verified
    calls = []
    def stop(*args):
        if calls:
            raise OSError('disk full')
        calls.append(1)
        return real(*args)
    monkeypatch.setattr(database_transfer, '_copy_verified', stop)
    with pytest.raises(OSError):
        migrate_database(source, dest, [], config_path=cfg)
    assert not cfg.exists()
    monkeypatch.setattr(database_transfer, '_copy_verified', real)
    assert migrate_database(source, dest, [], config_path=cfg)['verified']


def test_corrupt_snapshot_bytes_cannot_restore(isolated_database, tmp_path):
    snapshot = create_snapshot(isolated_database)
    (snapshot/'metadata/update_list.json').write_bytes(b'broken')
    with pytest.raises(DatabaseError):
        restore_snapshot(snapshot, tmp_path/'restore')
    assert not (tmp_path/'restore').exists()
