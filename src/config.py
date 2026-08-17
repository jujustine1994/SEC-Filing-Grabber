"""
config.py — Load and save config.json.
Merges loaded data with defaults so missing keys are always present.
"""

import json
import copy
import os
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent


def _default_config_path() -> Path:
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "SEC Financial Tools" / "config.json"
    return Path.home() / ".sec_financial_tools" / "config.json"

CONFIG_PATH = _default_config_path()

DEFAULT_CONFIG: dict = {
    # 介面與 Excel 顯示語言。代號清單見 i18n.LANGUAGES。
    #
    # 預設是**空字串而不是 "zh_tw"**：空字串代表「使用者還沒選過」，
    # main._pick_language_on_first_run() 靠它決定首次啟動要不要問。填了
    # "zh_tw" 就分不出「他選了繁中」和「他沒選過」，只能再加一個布林值，
    # 而兩個欄位描述同一件事遲早會不同步。
    #
    # 空字串餵給 i18n.set_lang() 會退回預設語言，所以就算問語言那步被跳過
    # （例如 cli.py 這種沒有 GUI 的路徑），程式照樣跑得動。
    "language": "",
    "identity": "",
    "output_dir": "output",
    "ticker_paths": {},
    "watchlist": [],
    "filename_format": "ticker_name",
    "filename_custom": "",
    "max_filings": 80,
    "template_path": "",
    # 輸出檔已存在時，抓取前先跳一次確認。使用者在那個對話框勾「不再提醒」
    # 會把這裡寫成 False，進階設定可以再打開。
    #
    # 預設 True 而不是 False：既有使用者的 config.json 沒有這個鍵，
    # load_config 會補上預設值，升級後第一次覆蓋要提醒得到。
    "warn_on_overwrite": True,
    "ai": {
        "provider": "google",
        "model": "gemini-flash-latest",
        "api_key": "",
    },
}


def load_config(path: Path | None = None) -> dict:
    """Load config.json, merging with defaults for any missing keys."""
    if path is None:
        path = CONFIG_PATH
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    if Path(path).exists():
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            # Malformed or unreadable config — proceed with defaults
            return cfg
        for key, default_val in DEFAULT_CONFIG.items():
            if key in data:
                if isinstance(default_val, dict) and isinstance(data[key], dict):
                    cfg[key].update(data[key])
                elif not isinstance(default_val, dict):
                    cfg[key] = data[key]
    return cfg


def save_config(cfg: dict, path: Path | None = None) -> None:
    """Save config dict to config.json as UTF-8 JSON."""
    if path is None:
        path = CONFIG_PATH
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
