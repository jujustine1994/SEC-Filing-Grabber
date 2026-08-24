"""缺漏帳本：抓不到的期數要被記下來並講出來（CTH 2026-08-17 定案）。

演進過程值得記著，免得日後有人改回去：

  第一版  抓不到就 `except Exception: continue`，靜默。使用者拿到少一季
          的 Excel 不會發現——這是要修的 bug。
  第二版  網路問題一律中止整趟、不寫檔。CTH 否決：「不希望抓得太嚴格讓
          資料永遠抓不出來」。
  定案    照抓，抓不到就留空，**但把缺了哪幾期主動講出來**，並附上
          「可能是網路問題」的判斷。使用者不必自己去發現。

分類不靠猜例外類別名稱（那份名單要跟著 httpx 版本走，漏一個就誤判），
失敗當下直接戳一次 SEC：戳得通＝資料問題，戳不通＝網路問題。
"""

import pytest

from fetch_ledger import FetchLedger


def _net_probe(reachable):
    return lambda: reachable


class Boom(Exception):
    pass


# ── 記錄 ──────────────────────────────────────────────────────────────────

def test_a_clean_run_has_no_gaps():
    led = FetchLedger()
    assert led.gaps == []
    assert not led.has_gaps


def test_records_what_could_not_be_fetched():
    """使用者要知道少了哪幾期才知道該不該重抓。"""
    led = FetchLedger(probe=_net_probe(True))
    led.record("FY2025Q2", Boom("bad XBRL"))
    led.record("FY2024Q4", Boom("bad XBRL"))
    assert [g.where for g in led.gaps] == ["FY2025Q2", "FY2024Q4"]
    assert led.has_gaps


def test_reachable_sec_means_the_data_is_the_problem():
    led = FetchLedger(probe=_net_probe(True))
    led.record("FY2025Q2", Boom("bad XBRL"))
    assert led.gaps[0].kind == "data"


def test_unreachable_sec_means_the_network_is_the_problem():
    led = FetchLedger(probe=_net_probe(False))
    led.record("FY2025Q2", Boom("bad XBRL"))
    assert led.gaps[0].kind == "network"


def test_the_probe_runs_at_most_once_per_run():
    """每一期失敗都戳一次 SEC 的話，整批斷線會多打幾十次沒必要的請求，
    而且每次都要等逾時。判斷一次就夠了。"""
    calls = []

    def counting_probe():
        calls.append(1)
        return False

    led = FetchLedger(probe=counting_probe)
    for i in range(5):
        led.record(f"FY2025Q{i}", Boom("x"))
    assert len(calls) == 1


def test_exhausted_retries_always_count_as_network():
    """實機驗收踩到的：NetworkDownError 代表「已退避重試三次都連不上」，
    那就是最強的斷網證據。但探測是在事後跑的，這時網路可能已經恢復——
    去問 SEC 就會得到「連得上」，於是斷網被報成「資料問題，重抓也一樣」，
    方向完全相反。這種例外不可以再走探測。"""
    from net_retry import NetworkDownError

    led = FetchLedger(probe=_net_probe(True))       # 探測說連得上
    led.record("FY2025Q2", NetworkDownError("ConnectTimeout after 3 attempts"))
    assert led.gaps[0].kind == "network"
    assert led.network_blamed


def test_probe_is_skipped_when_the_exception_already_says_network():
    """例外自己就長得像斷線時不必再戳，省一次逾時等待。"""
    calls = []

    class ConnectTimeout(Exception):
        pass

    led = FetchLedger(probe=lambda: calls.append(1) or True)
    led.record("FY2025Q2", ConnectTimeout("timed out"))
    assert led.gaps[0].kind == "network"
    assert calls == []


# ── 煞車：整個網路斷掉時不要乾等 ──────────────────────────────────────────

def test_no_brake_while_things_are_fine():
    led = FetchLedger(probe=_net_probe(True))
    for i in range(10):
        led.record(f"FY{i}", Boom("bad XBRL"))
    assert not led.give_up_retrying, "資料問題再多也不代表網路斷了"


def test_brake_engages_after_consecutive_network_failures():
    """整個網路斷掉時，40 份財報每份重試 2+4 秒 = 乾等 4 分鐘才給你一份
    空檔。連續幾期都是網路問題就別再重試了，快速跑完剩下的。"""
    led = FetchLedger(probe=_net_probe(False), brake_after=3)
    led.record("FY2025Q1", Boom("x"))
    led.record("FY2024Q4", Boom("x"))
    assert not led.give_up_retrying
    led.record("FY2024Q3", Boom("x"))
    assert led.give_up_retrying


def test_a_success_releases_the_brake():
    """中間抓到了就代表網路還在，前面只是閃斷。"""
    led = FetchLedger(probe=_net_probe(False), brake_after=3)
    led.record("FY2025Q1", Boom("x"))
    led.record("FY2024Q4", Boom("x"))
    led.succeeded()
    led.record("FY2024Q3", Boom("x"))
    assert not led.give_up_retrying


# ── 給人看的摘要 ──────────────────────────────────────────────────────────

def test_summary_is_empty_when_nothing_was_missed():
    assert FetchLedger().summary() == ""


def test_summary_names_the_periods():
    led = FetchLedger(probe=_net_probe(True))
    led.record("FY2025Q2", Boom("x"))
    led.record("FY2024Q4", Boom("x"))
    s = led.summary()
    assert "FY2025Q2" in s and "FY2024Q4" in s


def test_summary_says_when_the_network_was_to_blame():
    """網路造成的缺漏重抓有救，資料造成的重抓一樣沒救——要分得出來。"""
    led = FetchLedger(probe=_net_probe(False))
    led.record("FY2025Q2", Boom("x"))
    assert led.network_blamed
    assert not FetchLedger(probe=_net_probe(True)).network_blamed


def test_long_gap_lists_are_truncated():
    """抓 20 年的範圍整個斷線時會有上百期，全列出來擠爆畫面。"""
    led = FetchLedger(probe=_net_probe(False))
    for i in range(50):
        led.record(f"FY2020Q{i}", Boom("x"))
    s = led.summary()
    assert len(s) < 300
    assert "50" in s, "沒列完的要用數字交代總數"


def test_summary_never_leaks_exception_messages():
    """windows-tool.md 的規則：例外訊息會挾帶完整 URL 與 response 片段，
    不可以原樣落進 log 或畫面。"""
    led = FetchLedger(probe=_net_probe(True))
    led.record("FY2025Q2", Boom("https://sec.gov/x?apikey=SECRET"))
    assert "SECRET" not in led.summary()
    assert "sec.gov" not in led.summary()


# ── D11-B：網路缺漏自動重試一次，重試結果吸收回帳本 ─────────────────────────
#
# 只重試一次、只重試 network 那類（data 類重試沒用，見 fetch_ledger.py 開頭
# 的分類說明）。判斷「這期救回來了沒」不用猜——直接看同一個 `where` 標籤
# 在重試那輪的帳本裡還在不在：不在＝救回來了，在＝還是沒抓到。

def test_retry_absorbs_gaps_that_succeeded_the_second_time():
    """一期第一輪失敗、重試那輪沒有再記到同一個 where，代表救回來了。"""
    led = FetchLedger(probe=_net_probe(False))
    led.record("FY2025Q1", Boom("x"))
    led.record("FY2025Q2", Boom("x"))
    retry_led = FetchLedger(probe=_net_probe(False))
    retry_led.record("FY2025Q2", Boom("x"))   # 這期重試還是失敗
    led.absorbed_by_retry(retry_led)
    assert [g.where for g in led.gaps] == ["FY2025Q2"]


def test_retry_absorbs_all_gaps_when_second_pass_is_clean():
    led = FetchLedger(probe=_net_probe(False))
    led.record("FY2025Q1", Boom("x"))
    retry_led = FetchLedger(probe=_net_probe(False))   # 重試那輪一期都沒失敗
    led.absorbed_by_retry(retry_led)
    assert led.gaps == []


def test_retry_does_not_absorb_data_kind_gaps():
    """資料類缺漏本來就不指望重試救得回（重試也不會真的去補這種），
    不該被當成「救回來了」而移除。"""
    led = FetchLedger(probe=_net_probe(True))   # kind="data"
    led.record("FY2025Q1", Boom("x"))
    retry_led = FetchLedger(probe=_net_probe(True))
    led.absorbed_by_retry(retry_led)
    assert [g.where for g in led.gaps] == ["FY2025Q1"]
