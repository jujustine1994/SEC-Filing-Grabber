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
