import json
import pytest
from pathlib import Path
from config import load_config, save_config, DEFAULT_CONFIG


def test_load_returns_defaults_when_no_file(tmp_path):
    cfg = load_config(tmp_path / "nonexistent.json")
    assert cfg["output_dir"] == "output"
    assert cfg["watchlist"] == []
    assert cfg["ai"]["provider"] == "google"
    assert cfg["ai"]["model"] == "gemini-flash-latest"
    assert cfg["ai"]["api_key"] == ""


def test_load_reads_existing_file(tmp_path):
    data = {
        "identity": "Test User test@example.com",
        "output_dir": "my_output",
        "watchlist": [{"ticker": "AAPL", "name": "Apple Inc."}],
        "ai": {"provider": "google", "model": "gemini-flash-latest", "api_key": "abc123"},
    }
    cfg_path = tmp_path / "config.json"
    cfg_path.write_text(json.dumps(data), encoding="utf-8")
    cfg = load_config(cfg_path)
    assert cfg["identity"] == "Test User test@example.com"
    assert cfg["output_dir"] == "my_output"
    assert cfg["watchlist"][0]["ticker"] == "AAPL"
    assert cfg["ai"]["api_key"] == "abc123"


def test_save_writes_file(tmp_path):
    cfg_path = tmp_path / "config.json"
    cfg = {"identity": "X x@x.com", "output_dir": "output", "watchlist": [],
           "ai": {"provider": "google", "model": "gemini-flash-latest", "api_key": ""}}
    save_config(cfg, cfg_path)
    assert cfg_path.exists()
    loaded = json.loads(cfg_path.read_text(encoding="utf-8"))
    assert loaded["identity"] == "X x@x.com"


def test_load_merges_missing_keys(tmp_path):
    # partial config missing the ai block
    data = {"identity": "User user@example.com", "output_dir": "output", "watchlist": []}
    cfg_path = tmp_path / "config.json"
    cfg_path.write_text(json.dumps(data), encoding="utf-8")
    cfg = load_config(cfg_path)
    assert "ai" in cfg
    assert cfg["ai"]["model"] == "gemini-flash-latest"
