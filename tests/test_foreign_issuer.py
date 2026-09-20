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


# ── 財年結束月：沒有年報清單時的探測 ──────────────────────────────────────
#
# ⚠ **這是 D9 之後才長出來的漏洞，而且完全靜默**（2026-09-20 實測發現）。
# 只抓季報（`--quarterly-only`／GUI 只勾 10-Q）時 `filings_k` 是空的，
# `_fetch_gaap_impl()` 會探一份年報來判財年結束月——原本寫死問 `"10-K"`，
# FPI 交的是 20-F，探不到就 fallback 成 12（日曆年）。
#
# ARM 財年 3 月底結束，實測兩種抓法對**同一個期末日**給出不同標籤：
#
#   期末日 2024-09-30 → 季+年「FY2025Q2」（對）、只抓季報「FY2024Q3」（錯 2 季）
#   期末日 2024-12-31 → 季+年「FY2025Q3」（對）、只抓季報「FY2024Q4」（錯 2 季）
#
# 不報錯、不警告，整份 Excel 的期間標籤靜靜地錯掉——正是這個專案最不能接受的
# 失效模式（跟 G13 同一類）。

def test_fy_end_month_probe_asks_for_the_annual_form_of_this_filer_type(monkeypatch):
    """FPI 要探 20-F。寫死 10-K 的話 ARM 探不到，財年結束月會 fallback 成 12。"""
    import fetcher_gaap

    asked = []
    monkeypatch.setattr(fetcher_gaap, "_list_filings",
                        lambda company, form: asked.append(form) or ["annual"])
    monkeypatch.setattr(fetcher_gaap, "_detect_fy_end_month", lambda filings: 3)

    month = fetcher_gaap._probe_fy_end_month(object(), "20-F")

    assert asked == ["20-F"]
    assert month == 3


def test_fy_end_month_probe_still_asks_for_10k_for_a_domestic_filer(monkeypatch):
    """214 家的常態不可以被改壞。"""
    import fetcher_gaap

    asked = []
    monkeypatch.setattr(fetcher_gaap, "_list_filings",
                        lambda company, form: asked.append(form) or ["annual"])
    monkeypatch.setattr(fetcher_gaap, "_detect_fy_end_month", lambda filings: 9)

    assert fetcher_gaap._probe_fy_end_month(object(), "10-K") == 9
    assert asked == ["10-K"]


def test_fy_end_month_falls_back_to_december_when_no_annual_filing_exists(monkeypatch):
    """真的一份年報都沒有時才回 12。這條退路本身是對的——錯的是連問都問錯表單。"""
    import fetcher_gaap

    monkeypatch.setattr(fetcher_gaap, "_list_filings", lambda company, form: [])

    assert fetcher_gaap._probe_fy_end_month(object(), "20-F") == 12


# ── GUI 預覽（`preview_sheets`）────────────────────────────────────────────
#
# ⚠ 預覽寫死問 `"10-Q"`，沒有就整個 `return empty`。FPI 因此在 GUI 掃描時
# **拿不到最新期別／期末日／申報日**（三欄全空），但抓取本身是好的——使用者
# 看到空白會以為這家抓不到。2026-09-20 實測 ARM 三欄全空、AAPL 正常。

def test_preview_falls_back_to_6k_for_a_foreign_private_issuer(cache_dir, monkeypatch):
    """沒有 10-Q 就退到含財報的 6-K，跟 `_resolve_listings()` 同一個判準。"""
    import fetcher_gaap

    sixk = _fake_6k(ACC_WITH, r_files=68, filing_date="2026-07-29")
    sixk.period_of_report = "2026-06-30"

    company = MagicMock()
    company.get_filings.side_effect = lambda form, amendments=False: (
        [sixk] if form == "6-K" else [])
    company.fiscal_year_end = "0331"          # ARM：財年 3 月底結束

    monkeypatch.setattr(fetcher_gaap, "Company", lambda t: company)
    monkeypatch.setattr(fetcher_gaap, "set_identity", lambda i: None)

    result = fetcher_gaap.preview_sheets("ARM", "Test test@test.com")

    assert result["latest_period_end"] == "2026-06-30"
    assert result["filing_date"] == "2026-07-29"
    assert result["latest_label"] == "FY2027Q1"   # 財年 3 月底 → 4~6 月是 Q1


def test_preview_does_not_ask_for_6k_when_the_company_files_10q(cache_dir, monkeypatch):
    """214 家的常態：多問一次 6-K 對 Sony 那種公司是上千份的差別，不可以白問。"""
    import fetcher_gaap

    tenq = MagicMock()
    tenq.period_of_report = "2025-12-27"
    tenq.filing_date = "2026-02-01"
    tenq.accession_no = "0000320193-26-000001"

    asked = []

    def get_filings(form, amendments=False):
        asked.append(form)
        return [tenq] if form == "10-Q" else []

    company = MagicMock()
    company.get_filings.side_effect = get_filings
    company.fiscal_year_end = "0926"

    monkeypatch.setattr(fetcher_gaap, "Company", lambda t: company)
    monkeypatch.setattr(fetcher_gaap, "set_identity", lambda i: None)

    fetcher_gaap.preview_sheets("AAPL", "Test test@test.com")

    assert "6-K" not in asked


# ── 預覽的財年結束月要跟主路徑同源（J11）──────────────────────────────────
#
# 預覽原本走 `Company.fiscal_year_end` 屬性（免費，已在物件上），主路徑走
# `_detect_fy_end_month()`（解 10-K 全文）。**兩個依據，會分岔。**
#
# 全庫 215 家實測（基準＝解全文）：
#   `period_of_report` 月份 …… 215/215 = 100.0%
#   `fiscal_year_end` 屬性 …… 206/215 = 95.8%（ADI／DIS／HD／JNJ／LULU／
#                               MU／ONTO／SNDK／TGT 九家對不上）
#
# 其中 **JNJ 與 ONTO 的標籤真的錯了，而且差一整個財年**：
#   JNJ  預覽 FY2027Q2 vs 實抓 FY2026Q2
#   ONTO 預覽 FY2026Q2 vs 實抓 FY2027Q2
#
# 改用年報清單的 `period_of_report`：準確度跟解全文一樣（100%），成本只是一次
# filing 清單請求——預覽的成本上限是「不要解全文」，不是「不要問清單」。

def _annual(period_of_report: str):
    f = MagicMock()
    f.period_of_report = period_of_report
    return f


def test_preview_takes_the_fiscal_year_end_from_the_annual_filing_not_the_attribute(
        cache_dir, monkeypatch):
    """屬性跟年報期末日打架時要信年報——JNJ／ONTO 就是這樣差掉一整個財年。"""
    import fetcher_gaap

    tenq = MagicMock()
    tenq.period_of_report = "2026-06-28"
    tenq.filing_date = "2026-07-15"

    company = MagicMock()
    company.get_filings.side_effect = lambda form, amendments=False: (
        [tenq] if form == "10-Q" else [_annual("2025-12-28")] if form == "10-K" else [])
    company.fiscal_year_end = "0101"          # 屬性說 1 月（JNJ 實際就是這樣錯的）

    monkeypatch.setattr(fetcher_gaap, "Company", lambda t: company)
    monkeypatch.setattr(fetcher_gaap, "set_identity", lambda i: None)

    result = fetcher_gaap.preview_sheets("JNJ", "Test test@test.com")

    # 年報期末日 2025-12-28 → 財年 12 月底 → 曆年制 → 6/28 是 Q2
    assert result["latest_label"] == "FY2026Q2"


def test_preview_asks_the_annual_form_of_this_filer_type(cache_dir, monkeypatch):
    """FPI 要問 20-F，不是 10-K——問錯就拿不到期末日，白花一次請求又退回屬性。"""
    import fetcher_gaap

    sixk = _fake_6k(ACC_WITH, r_files=68, filing_date="2026-07-29")
    sixk.period_of_report = "2026-06-30"

    asked = []

    def get_filings(form, amendments=False):
        asked.append(form)
        return {"6-K": [sixk], "20-F": [_annual("2026-03-31")]}.get(form, [])

    company = MagicMock()
    company.get_filings.side_effect = get_filings
    company.fiscal_year_end = ""               # 屬性拿不到，只能靠年報

    monkeypatch.setattr(fetcher_gaap, "Company", lambda t: company)
    monkeypatch.setattr(fetcher_gaap, "set_identity", lambda i: None)

    result = fetcher_gaap.preview_sheets("ARM", "Test test@test.com")

    assert "20-F" in asked and "10-K" not in asked
    assert result["latest_label"] == "FY2027Q1"   # 財年 3 月底 → 4~6 月是 Q1


def test_preview_falls_back_to_the_attribute_when_no_annual_filing_exists(
        cache_dir, monkeypatch):
    """剛上市、只交過一份 10-Q 的公司沒有年報可問——退回屬性比直接用 12 月好。"""
    import fetcher_gaap

    tenq = MagicMock()
    tenq.period_of_report = "2026-06-27"
    tenq.filing_date = "2026-07-31"

    company = MagicMock()
    company.get_filings.side_effect = lambda form, amendments=False: (
        [tenq] if form == "10-Q" else [])
    company.fiscal_year_end = "0926"           # AAPL：財年 9 月底

    monkeypatch.setattr(fetcher_gaap, "Company", lambda t: company)
    monkeypatch.setattr(fetcher_gaap, "set_identity", lambda i: None)

    assert fetcher_gaap.preview_sheets("AAPL", "x")["latest_label"] == "FY2026Q3"


# ── 當期整期不在報表裡，要出聲（G13(a) 第一步：偵測）────────────────────
#
# edgartools 的期間篩選門檻會把當期整個丟掉（根因見 TODO G13(a)）。
# **現在的症狀是整季空白而且完全不報錯**——使用者不知道是公司沒報、SEC 沒有、
# 還是工具壞了。這一步不碰任何數字，只讓它出聲。
#
# 判準用 `period_of_report`：那是 SEC 官方認定「這份 filing 報的是哪一期」，
# 不是我們猜的。一個判斷涵蓋所有型：
#   ISRG 型 —— 挑到的是去年同期（`2017-03-31` != `2018-03-31`）
#   DELL 型 —— 報表裡只剩年度欄（`2023-02-03 (FY)` != `2023-05-05`）

def _filing_with_period(period_of_report):
    f = MagicMock()
    f.period_of_report = period_of_report
    f.accession_no = "0001035267-18-000048"
    f.filing_date = "2018-04-18"
    f.form = "10-Q"
    return f


def test_a_filing_whose_statement_lacks_its_own_period_is_recorded_as_a_gap():
    """ISRG 型：報表裡挑得到欄，但那是**去年同期**。"""
    import fetcher_gaap
    from fetch_ledger import FetchLedger

    ledger = FetchLedger(probe=lambda: True)
    token = fetcher_gaap._ledger_var.set(ledger)
    try:
        fetcher_gaap._note_missing_current_period(
            _filing_with_period("2018-03-31"), "2017-03-31 (Q1)")
    finally:
        fetcher_gaap._ledger_var.reset(token)

    assert len(ledger.gaps) == 1
    assert ledger.gaps[0].kind == "data"
    assert "2018-03-31" in ledger.gaps[0].where


def test_a_statement_with_only_an_annual_column_is_also_recorded():
    """DELL／SNOW 型：季報的損益表只剩年度欄。"""
    import fetcher_gaap
    from fetch_ledger import FetchLedger

    ledger = FetchLedger(probe=lambda: True)
    token = fetcher_gaap._ledger_var.set(ledger)
    try:
        fetcher_gaap._note_missing_current_period(
            _filing_with_period("2023-05-05"), "2023-02-03 (FY)")
    finally:
        fetcher_gaap._ledger_var.reset(token)

    assert len(ledger.gaps) == 1


def test_no_usable_column_at_all_is_recorded_too():
    import fetcher_gaap
    from fetch_ledger import FetchLedger

    ledger = FetchLedger(probe=lambda: True)
    token = fetcher_gaap._ledger_var.set(ledger)
    try:
        fetcher_gaap._note_missing_current_period(_filing_with_period("2018-03-31"), None)
    finally:
        fetcher_gaap._ledger_var.reset(token)

    assert len(ledger.gaps) == 1


def test_a_healthy_filing_is_not_recorded():
    """⚠ 最重要的一條：99.9% 的 filing 是好的，不可以誤報。"""
    import fetcher_gaap
    from fetch_ledger import FetchLedger

    ledger = FetchLedger(probe=lambda: True)
    token = fetcher_gaap._ledger_var.set(ledger)
    try:
        fetcher_gaap._note_missing_current_period(
            _filing_with_period("2018-03-31"), "2018-03-31 (Q1)")
    finally:
        fetcher_gaap._ledger_var.reset(token)

    assert ledger.gaps == []


def test_a_filing_without_a_period_of_report_is_not_recorded():
    """拿不到 `period_of_report` 就無從判斷——不可以把「不知道」報成「有缺漏」。"""
    import fetcher_gaap
    from fetch_ledger import FetchLedger

    ledger = FetchLedger(probe=lambda: True)
    token = fetcher_gaap._ledger_var.set(ledger)
    try:
        fetcher_gaap._note_missing_current_period(_filing_with_period(""), "2017-03-31 (Q1)")
    finally:
        fetcher_gaap._ledger_var.reset(token)

    assert ledger.gaps == []
