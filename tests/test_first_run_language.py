"""首次啟動選語言：問一次、記住、不再跳。

這個對話框只在「還沒選過」時出現，判斷依據是 config.json 的 `language`
不是合法代號。測試不真的開視窗（CI 沒有顯示器，而且 wait_window 會卡住），
改為驗證那個判斷本身與存檔行為——真正會出錯的是這兩件事，不是 tkinter 版面。
"""

import json

import pytest

import config
import i18n
import main


@pytest.fixture
def cfg_path(tmp_path):
    return tmp_path / "config.json"


def _saved(path):
    return json.loads(path.read_text(encoding="utf-8"))


# ── 「還沒選過」的判斷 ────────────────────────────────────────────────────

def test_fresh_config_has_no_language_chosen():
    """預設值必須是空字串。填 "zh_tw" 就分不出「選了繁中」與「沒選過」，
    首次啟動的對話框永遠不會出現。"""
    assert config.DEFAULT_CONFIG["language"] == ""
    assert not i18n.is_supported(config.DEFAULT_CONFIG["language"])


def test_missing_key_counts_as_not_chosen(cfg_path):
    """既有使用者的 config.json 沒有 language 這個鍵——他們也沒選過，
    下次啟動該問一次。"""
    cfg_path.write_text(json.dumps({"identity": "A b@c.com"}), encoding="utf-8")
    cfg = config.load_config(cfg_path)
    assert not i18n.is_supported(cfg.get("language", ""))


@pytest.mark.parametrize("value", ["zh_tw", "zh_cn", "en", "ja"])
def test_a_real_choice_is_never_asked_again(cfg_path, value):
    cfg_path.write_text(json.dumps({"language": value}), encoding="utf-8")
    cfg = config.load_config(cfg_path)
    assert i18n.is_supported(cfg["language"])


@pytest.mark.parametrize("value", ["", "kl_ingon", "zh", "EN"])
def test_garbage_language_values_ask_again(cfg_path, value):
    """舊版留下的怪值、手改壞的值——當成沒選過再問一次，比靜默用預設好。"""
    cfg_path.write_text(json.dumps({"language": value}), encoding="utf-8")
    cfg = config.load_config(cfg_path)
    assert not i18n.is_supported(cfg["language"])


# ── 選完之後 ──────────────────────────────────────────────────────────────

def test_choice_survives_a_round_trip(cfg_path):
    cfg = config.load_config(cfg_path)
    cfg["language"] = "ja"
    config.save_config(cfg, cfg_path)
    assert _saved(cfg_path)["language"] == "ja"
    assert config.load_config(cfg_path)["language"] == "ja"


def test_saving_a_choice_keeps_the_rest_of_the_config(cfg_path):
    """語言是在主視窗建起來**之前**存的，那時 identity / API key 都還沒載入
    到 GUI。存檔用的是剛讀出來的整份 cfg，不可以只寫語言那一欄把其他洗掉。"""
    cfg_path.write_text(json.dumps({
        "identity": "CTH x@y.com",
        "max_filings": 40,
        "ai": {"provider": "google", "model": "m", "api_key": "SECRET"},
    }), encoding="utf-8")

    cfg = config.load_config(cfg_path)
    cfg["language"] = "en"
    config.save_config(cfg, cfg_path)

    got = _saved(cfg_path)
    assert got["language"] == "en"
    assert got["identity"] == "CTH x@y.com"
    assert got["max_filings"] == 40
    assert got["ai"]["api_key"] == "SECRET"


def test_picker_is_skipped_when_already_chosen(cfg_path, monkeypatch):
    """已選過時 `_pick_language_on_first_run` 必須在建任何 widget 之前就 return。

    用一個會爆的 Toplevel 釘住：只要它動手建視窗就會失敗。
    """
    cfg_path.write_text(json.dumps({"language": "en"}), encoding="utf-8")
    monkeypatch.setattr(main, "CONFIG_PATH", cfg_path)
    monkeypatch.setattr(main, "_migrate_config_if_needed", lambda: None)

    def _boom(*a, **k):
        raise AssertionError("已經選過語言了，不該再跳對話框")

    monkeypatch.setattr(main.tk, "Toplevel", _boom)
    main._pick_language_on_first_run(root=None)      # 沒建 widget 就用不到 root


def test_language_menu_offers_every_registered_language():
    """對話框的按鈕是從 LANGUAGES 生出來的，新增語言不必改這支函式。"""
    codes = [c for c, _ in i18n.available_languages()]
    assert codes == [c for c, _, _ in i18n.LANGUAGES]
    assert len(codes) >= 4
