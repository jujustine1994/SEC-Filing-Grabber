import json
from pathlib import Path

import pytest

from database import (DatabaseError, connect_database, create_database,
                      database_root, read_marker)


def test_missing_registered_database_does_not_create_empty(tmp_path):
    cfg = tmp_path / 'config.json'
    missing = tmp_path / 'missing'
    cfg.write_text(json.dumps({'database_path': str(missing),
                               'database_id': '00000000-0000-4000-8000-000000000000'}))
    with pytest.raises(DatabaseError):
        database_root(config_path=cfg)
    assert not missing.exists()


def test_connect_preserves_settings_and_checks_id(tmp_path, monkeypatch):
    root = tmp_path / 'formal'
    marker = create_database(root)
    cfg = tmp_path / 'config.json'
    cfg.write_text(json.dumps({'identity': 'private', 'language': 'ja'}))
    connect_database(root, config_path=cfg)
    assert database_root(config_path=cfg) == root.resolve()
    assert json.loads(cfg.read_text())['identity'] == 'private'
    monkeypatch.delenv('SEC_LOCAL_DB_ROOT', raising=False)
    marker['database_id'] = '00000000-0000-4000-8000-000000000000'
    (root / 'database.json').write_text(json.dumps(marker))
    with pytest.raises(DatabaseError):
        database_root(config_path=cfg)


def test_corrupt_config_does_not_discover_other_database(tmp_path, monkeypatch):
    import database
    root = tmp_path / 'default'
    create_database(root)
    monkeypatch.setattr(database, 'default_database_path', lambda: root)
    cfg = tmp_path / 'config.json'
    cfg.write_text('{broken')
    with pytest.raises(DatabaseError):
        database_root(config_path=cfg)


def test_override_requires_marker(tmp_path, monkeypatch):
    monkeypatch.setenv('SEC_LOCAL_DB_ROOT', str(tmp_path))
    with pytest.raises(DatabaseError):
        database_root()
    assert not (tmp_path / 'filings').exists()


def test_create_does_not_relabel_existing_folder_as_test(tmp_path):
    (tmp_path / 'keep.txt').write_text('keep')
    with pytest.raises(DatabaseError):
        create_database(tmp_path, test_mode=True)
    assert (tmp_path / 'keep.txt').read_text() == 'keep'


def test_formal_database_rejected_inside_project(tmp_path):
    import database
    with pytest.raises(DatabaseError):
        create_database(database.project_root() / 'forbidden-database')


def test_empty_valid_database_is_connected(tmp_path, monkeypatch):
    root = tmp_path / 'empty-valid'
    marker = create_database(root, test_mode=True)
    monkeypatch.setenv('SEC_LOCAL_DB_ROOT', str(root))
    assert database_root() == root.resolve()
    assert read_marker(root)['database_id'] == marker['database_id']
    assert json.loads((root / 'metadata/update_list.json').read_text())['tickers'] == []


def test_explicit_reconnect_preserves_corrupt_settings(tmp_path):
    root = tmp_path/'formal'
    create_database(root)
    cfg = tmp_path/'config.json'
    cfg.write_bytes(b'{broken')
    connect_database(root, config_path=cfg)
    assert database_root(config_path=cfg) == root.resolve()
    assert next(tmp_path.glob('config.json.corrupt-*.bak')).read_bytes() == b'{broken'
