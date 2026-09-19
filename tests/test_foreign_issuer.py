"""外國私人發行人（FPI）的 6-K／20-F 抓取路徑（TODO D9 A 路線）。

ARM 這類 FPI 不交 10-Q／10-K，交的是 **6-K（季報）＋ 20-F（年報）**。
6-K 是大雜燴——ARM 33 份裡只有 9 份含財報，其餘是股東會通知、人事公告。
判準是 `R*.htm` 的數量（2026-09-18 實測 33/33 完全分離：有財報的 61~72 個、
沒有的**全部都是 1 個**），R 檔是 SEC 財報檢視器從 XBRL 三表產生的。

⚠ **判定必須是動態的、而且要落檔**。無條件把 6-K 加進清單（原本傾向的 B
路線）對日本公司會很糟：Sony 有 1,055 份 6-K、野村 1,244 份，而且**一份都
沒有財報**。反過來每輪重問也不行——那是每家每輪上千次 index 請求。

⚠ **範圍：只做 us-gaap 的 FPI（ARM）**。IFRS 的 20-F（TM／SONY／HMC）
edgartools 解析得出來，但比對層的科目對照表對不上（損益表 24 列只命中 10 列），
那是另一件事，CTH 2026-09-19 決定先不做。
"""
from unittest.mock import MagicMock

import pytest

import filing_cache
from fetcher_gaap import (_filings_with_statements, _r_file_count,
                          _resolve_listings)

ACC_WITH = "0001045810-25-000123"
ACC_WITHOUT = "0001045810-25-000124"


@pytest.fixture
def cache_dir(tmp_path, monkeypatch):
    monkeypatch.setenv("SEC_LOCAL_DB_ROOT", str(tmp_path))
    return tmp_path / "filing_cache"


def _attachment(document: str, document_type: str = "HTML"):
    att = MagicMock()
    att.document = document
    att.document_type = document_type
    return att


def _fake_6k(acc: str, r_files: int, filing_date="2026-07-29"):
    """一份假的 6-K。`attachments` 被碰幾次是這組測試的主要斷言——
    碰一次就是一次 SEC index 請求。"""
    filing = MagicMock()
    filing.accession_no = acc
    filing.filing_date = filing_date
    filing.form = "6-K"
    atts = [_attachment(f"R{i}.htm") for i in range(1, r_files + 1)]
    atts.append(_attachment("arm-20260630.htm"))       # 主文件，不是 R 檔
    atts.append(_attachment("R2.xml", document_type="XML"))  # 不是 HTML
    type(filing).attachments = property(lambda self: atts)
    return filing


# ── R 檔數量 ─────────────────────────────────────────────────────────────

def test_r_file_count_only_counts_html_files_named_R_something():
    assert _r_file_count(_fake_6k(ACC_WITH, r_files=68)) == 68


def test_r_file_count_is_none_when_the_index_cannot_be_read():
    """⚠ **「讀到了，0 個」跟「讀不到」必須分開**（2026-09-19 實測抓到）。
    ARM 沒財報的 6-K 是 **1 個** R 檔，TM（豐田）的是 **0 個**——把 0 當成
    「讀不到」就會走保守路徑，TM 的 634 份 6-K 全部被留下來下載。"""
    filing = MagicMock()
    type(filing).attachments = property(
        lambda self: (_ for _ in ()).throw(OSError("index unavailable")))
    assert _r_file_count(filing) is None


def test_r_file_count_is_zero_for_a_filing_that_really_has_no_r_files():
    filing = MagicMock()
    type(filing).attachments = property(
        lambda self: [_attachment("tm-6k.htm"), _attachment("ex99.htm")])
    assert _r_file_count(filing) == 0


# ── 過濾 ─────────────────────────────────────────────────────────────────

def test_a_6k_carrying_financial_statements_is_kept(cache_dir):
    keep = _fake_6k(ACC_WITH, r_files=68)
    assert _filings_with_statements("ARM", [keep]) == [keep]


def test_a_6k_with_only_the_cover_page_is_filtered_out(cache_dir):
    """ARM 沒有財報的 6-K 都是 1 個 R 檔（封面那張）。"""
    assert _filings_with_statements("ARM", [_fake_6k(ACC_WITHOUT, r_files=1)]) == []


def test_a_6k_with_no_r_files_at_all_is_filtered_out(cache_dir):
    """TM（豐田）2026-09-03 那份實測就是 0 個 R 檔。TM 有 634 份 6-K，
    這條沒守住就是每輪多 634 份下載。"""
    assert _filings_with_statements("TM", [_fake_6k(ACC_WITHOUT, r_files=0)]) == []


def test_the_index_is_requested_only_once_per_filing(cache_dir):
    """第二輪要零請求。Sony 1,055 份每輪重問就是這條沒守住。"""
    calls = []
    filing = _fake_6k(ACC_WITHOUT, r_files=1)
    atts = filing.attachments
    type(filing).attachments = property(
        lambda self: (calls.append(1), atts)[1])

    _filings_with_statements("ARM", [filing])
    _filings_with_statements("ARM", [filing])

    assert len(calls) == 1
    assert filing_cache.load_sixk_probe("ARM") == {ACC_WITHOUT: 1}


def test_a_filing_whose_index_fails_is_kept_and_not_remembered(cache_dir):
    """判定失敗要往「留著」倒：多下載一份無用的 6-K 只是浪費一次請求，
    誤丟掉一份有財報的 6-K 是那一季**永久消失而且沒有症狀**。
    而且失敗不寫進快取——下次還要再問一次。"""
    filing = MagicMock()
    filing.accession_no = ACC_WITH
    type(filing).attachments = property(
        lambda self: (_ for _ in ()).throw(OSError("index unavailable")))

    assert _filings_with_statements("ARM", [filing]) == [filing]
    assert filing_cache.load_sixk_probe("ARM") == {}


def test_a_filing_without_an_accession_is_kept_without_being_probed(cache_dir):
    """拿不到 accession 就沒有快取鍵，跟 `_filing_obj()` 同一個處理：
    不快取、照樣往下走。"""
    filing = MagicMock()
    filing.accession_no = None
    assert _filings_with_statements("ARM", [filing]) == [filing]
    assert filing_cache.load_sixk_probe("ARM") == {}


# ── 表單來源的選擇（沒有 10-Q 才退到 6-K）────────────────────────────────

def _listing_stub(table: dict, calls: list):
    def _listing(form: str):
        calls.append(form)
        return list(table.get(form, []))
    return _listing


def test_a_normal_10q_filer_never_asks_for_6k(cache_dir):
    """正常美股 214 家都走這條。多問一次 6-K 清單對 Sony 那種公司是
    上千份的差別，對正常公司則是白白多一次請求。"""
    calls = []
    listing = _listing_stub({"10-Q": ["q1"], "10-K": ["k1"]}, calls)

    result = _resolve_listings("NVDA", listing)

    assert result["quarterly"] == ["q1"]
    assert result["annual"] == ["k1"]
    assert result["forms"] == ("10-Q", "10-K")
    assert calls == ["10-Q", "10-K"]


def test_a_filer_with_no_10q_falls_back_to_6k_and_20f(cache_dir):
    calls = []
    six_k = [_fake_6k(ACC_WITH, r_files=68)]
    listing = _listing_stub({"6-K": six_k, "20-F": ["annual"]}, calls)

    result = _resolve_listings("ARM", listing)

    assert result["quarterly"] == six_k
    assert result["annual"] == ["annual"]
    assert result["forms"] == ("6-K", "20-F")


def test_the_6k_fallback_is_filtered_by_r_file_count(cache_dir):
    """ARM 33 份 6-K 只有 9 份有財報。沒過濾就會把 24 份公告當季報下載。"""
    calls = []
    keep = _fake_6k(ACC_WITH, r_files=68)
    drop = _fake_6k(ACC_WITHOUT, r_files=1)
    listing = _listing_stub({"6-K": [keep, drop], "20-F": ["annual"]}, calls)

    assert _resolve_listings("ARM", listing)["quarterly"] == [keep]


def test_a_ticker_with_nothing_anywhere_reports_both_form_families(cache_dir):
    """SEC 查無此代號（ADR Level I，例如 NTDOY／IFX）。錯誤訊息要講得出
    「10-Q 跟 6-K 都沒有」，不然使用者會以為只是這個工具沒支援。"""
    calls = []
    listing = _listing_stub({}, calls)

    with pytest.raises(ValueError, match="10-Q"):
        _resolve_listings("NTDOY", listing)
    assert "6-K" in calls


def test_annual_only_mode_still_finds_the_20f(cache_dir):
    calls = []
    listing = _listing_stub({"20-F": ["annual"]}, calls)

    result = _resolve_listings("ARM", listing, fetch_quarterly=False)

    assert result["annual"] == ["annual"]
    assert result["forms"] == ("6-K", "20-F")
    assert "10-Q" not in calls
