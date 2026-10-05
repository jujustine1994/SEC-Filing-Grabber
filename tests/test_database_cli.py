import json

import pytest
import cli
from database import create_database


def test_cli_connect_rejects_unmarked_folder(tmp_path, capsys):
    assert cli.main(['db-connect', str(tmp_path), '--config-path', str(tmp_path/'cfg.json')]) != 0
    assert not (tmp_path/'cfg.json').exists()


def test_cli_connect_status_and_snapshot(tmp_path, monkeypatch, capsys):
    root = tmp_path/'formal'
    marker = create_database(root)
    cfg = tmp_path/'cfg.json'
    assert cli.main(['db-connect', str(root), '--config-path', str(cfg)]) == 0
    capsys.readouterr()
    monkeypatch.delenv('SEC_LOCAL_DB_ROOT')
    assert cli.main(['db-status','--config-path',str(cfg),'--json','-']) == 0
    data = json.loads(capsys.readouterr().out)
    assert data['database_id'] == marker['database_id']
    dest = tmp_path/'snapshot'
    assert cli.main(['db-snapshot','--config-path',str(cfg),'--destination',str(dest)]) == 0
    assert (dest/'snapshot.json').exists()


def test_cli_missing_registered_database_returns_connection_error(tmp_path, monkeypatch, capsys):
    cfg = tmp_path/'cfg.json'
    cfg.write_text(json.dumps({'database_path':str(tmp_path/'missing'),'database_id':'lost'}))
    monkeypatch.delenv('SEC_LOCAL_DB_ROOT')
    assert cli.main(['db-status','--config-path',str(cfg)]) != 0
    assert 'db-connect' in capsys.readouterr().err


def test_gui_disconnected_can_open_and_has_no_clear_buttons(tmp_path, monkeypatch):
    import tkinter as tk
    import main
    monkeypatch.setenv('SEC_LOCAL_DB_ROOT',str(tmp_path/'missing'))
    monkeypatch.setattr(main, '_migrate_config_if_needed',lambda:None)
    root = tk.Tk(); root.withdraw()
    try:
        app = main.SECFetcherApp(root)
        assert not app._database_connected()
        assert not getattr(app,'_cache_clear_all_btn',None)
        assert 'disabled' == str(app.btn_run_single.cget('state'))
    finally:
        root.destroy()
