"""
config.py — Load and save config.json.
Merges loaded data with defaults so missing keys are always present.
"""

import json
import copy
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent

DEFAULT_CONFIG: dict = {
    "identity": "",
    "output_dir": "output",
    "watchlist": [],
    "ai": {
        "provider": "google",
        "model": "gemini-flash-latest",
        "api_key": "",
    },
}


def load_config(path: Path | None = None) -> dict:
    """Load config.json, merging with defaults for any missing keys."""
    if path is None:
        path = SCRIPT_DIR / "config.json"
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    if Path(path).exists():
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        for key, default_val in DEFAULT_CONFIG.items():
            if key in data:
                if isinstance(default_val, dict):
                    cfg[key].update(data[key])
                else:
                    cfg[key] = data[key]
    return cfg


def save_config(cfg: dict, path: Path | None = None) -> None:
    """Save config dict to config.json as UTF-8 JSON."""
    if path is None:
        path = SCRIPT_DIR / "config.json"
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
