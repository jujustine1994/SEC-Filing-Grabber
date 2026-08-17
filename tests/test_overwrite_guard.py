"""覆蓋既有輸出檔前先提示（CTH 2026-08-17 要求）。

現況是不問就覆蓋。雖然 Data_* 以外的 sheet（My_* 分析頁）會保留、
而且覆蓋前有備份，但使用者事前完全不知道自己要蓋掉誰——尤其
`_build_output_path()` 的檔名是程式組的（`AAPL Apple Inc data.xlsx`），
使用者未必記得上次存到哪。

跟 test_first_run_language.py 一樣不開視窗：真正會出錯的是「什麼時候
該問」這個判斷，不是對話框版面。
"""

from pathlib import Path

import pytest

import config
import i18n
import main

LANGS = [code for code, _, _ in i18n.LANGUAGES]


@pytest.fixture
def existing(tmp_path):
    p = tmp_path / "AAPL Apple Inc data.xlsx"
    p.write_bytes(b"not really xlsx")
    return p


@pytest.fixture
def absent(tmp_path):
    return tmp_path / "NVDA.xlsx"


# ── 什麼時候該問 ──────────────────────────────────────────────────────────

def test_warns_when_file_exists(existing):
    assert main.should_warn_overwrite({"warn_on_overwrite": True}, existing)


def test_silent_when_file_is_new(absent):
    """第一次抓這家公司沒有東西被蓋掉，不該多一次點擊。"""
    assert not main.should_warn_overwrite({"warn_on_overwrite": True}, absent)


def test_silent_when_user_opted_out(existing):
    assert not main.should_warn_overwrite({"warn_on_overwrite": False}, existing)


def test_missing_key_still_warns(existing):
    """既有使用者的 config.json 沒有這個鍵。升級後預設要提醒，
    不可以因為鍵不存在就當成「他選過不要提醒」。"""
    assert main.should_warn_overwrite({}, existing)


def test_default_config_has_the_warning_on():
    assert config.DEFAULT_CONFIG["warn_on_overwrite"] is True


def test_flag_survives_a_config_round_trip(tmp_path):
    """關掉提醒之後要真的記住——存檔再讀回來不能被預設值蓋回 True。"""
    path = tmp_path / "config.json"
    cfg = config.load_config(path)
    cfg["warn_on_overwrite"] = False
    config.save_config(cfg, path)
    assert config.load_config(path)["warn_on_overwrite"] is False


# ── 批量更新：一次問完，不逐檔問 ──────────────────────────────────────────

def test_batch_lists_only_the_files_that_exist(tmp_path, existing, absent):
    other = tmp_path / "MSFT.xlsx"
    other.write_bytes(b"x")
    found = main.existing_outputs([existing, absent, other])
    assert found == [existing, other]


def test_batch_returns_empty_when_nothing_would_be_overwritten(absent):
    assert main.existing_outputs([absent]) == []


def test_batch_preserves_input_order(tmp_path):
    """訊息裡要列出檔名，順序跟使用者的 watchlist 一致才好對。"""
    paths = []
    for name in ("C.xlsx", "A.xlsx", "B.xlsx"):
        p = tmp_path / name
        p.write_bytes(b"x")
        paths.append(p)
    assert main.existing_outputs(paths) == paths


# ── 訊息文字 ──────────────────────────────────────────────────────────────

@pytest.mark.parametrize("lang", LANGS)
def test_overwrite_messages_exist_in_every_language(lang):
    s = i18n._strings(lang)
    for key in ("gui.dlg.overwrite_title", "gui.msg.overwrite_single",
                "gui.msg.overwrite_batch", "gui.chk.dont_warn_again",
                "gui.chk.warn_on_overwrite", "gui.btn.overwrite_continue"):
        assert key in s, f"{lang} 缺 {key}"


@pytest.mark.parametrize("lang", LANGS)
def test_single_message_names_the_file_and_the_backup(lang):
    """使用者要能從訊息本身看出「蓋掉哪個檔」與「備份叫什麼」，
    否則提示了也不知道出事後去哪裡救。"""
    msg = i18n._strings(lang)["gui.msg.overwrite_single"]
    assert "{name}" in msg
    assert "{backup}" in msg
