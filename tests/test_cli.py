"""cli.py 的測試（TODO B1）。

全部離線：網路那一層集中在 `cli._gaap_tables` / `cli._earnings_filings` /
`cli._press_release_html` 三個函式，測試把它們換掉。這也是刻意的設計——
CLI 只做參數解析與輸出格式化，抓取邏輯一律轉呼叫既有核心函式。
"""
from __future__ import annotations

import json
import sys

import pytest

import cli
from fetcher_gaap import StatementTable
from press_release_tables import PressTable


# ── --years 解析 ────────────────────────────────────────────────────────────

@pytest.mark.parametrize("text, expected", [
    ("2023-2026", (2023, 2026)),
    ("2024", (2024, 2024)),
    ("2020-", (2020, None)),
    ("-2020", (None, 2020)),
])
def test_parse_years(text, expected):
    assert cli.parse_years(text) == expected


def test_parse_years_none():
    assert cli.parse_years(None) == (None, None)


@pytest.mark.parametrize("bad", ["abc", "2023-2020", "20-26", "2023-2024-2025"])
def test_parse_years_rejects_garbage(bad):
    with pytest.raises(cli.CliError):
        cli.parse_years(bad)


# ── gaap 子指令 ─────────────────────────────────────────────────────────────

def _fake_table(sheet_name="Data_Financials(Q)") -> StatementTable:
    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=["FY2026Q1", "FY2026Q2"],
        filing_dates=["2026-05-07", "2026-08-06"],
        concepts=["Revenue", "Net Income"],
        values=[[100.0, 110.0], [10.0, 11.0]],
        labels=["Revenues", "NetIncomeLoss"],
        period_ends=["2026-03-29", "2026-06-28"],
    )


def test_gaap_writes_xlsx(tmp_path, monkeypatch):
    out = tmp_path / "AAPL.xlsx"
    monkeypatch.setattr(cli, "_gaap_tables", lambda **kw: [_fake_table()])
    rc = cli.main(["gaap", "AAPL", "--xlsx", str(out), "--identity", "T t@e.com"])
    assert rc == 0
    assert out.exists() and out.stat().st_size > 0


def test_gaap_passes_year_range_through(tmp_path, monkeypatch):
    seen = {}

    def _capture(**kw):
        seen.update(kw)
        return [_fake_table()]

    monkeypatch.setattr(cli, "_gaap_tables", _capture)
    cli.main(["gaap", "AAPL", "--years", "2023-2026",
              "--xlsx", str(tmp_path / "a.xlsx"), "--identity", "T t@e.com"])
    assert seen["start_year"] == 2023
    assert seen["end_year"] == 2026


def test_gaap_json_output(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(cli, "_gaap_tables", lambda **kw: [_fake_table()])
    rc = cli.main(["gaap", "AAPL", "--json", "-", "--identity", "T t@e.com"])
    assert rc == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["ticker"] == "AAPL"
    sheet = payload["sheets"][0]
    assert sheet["sheet_name"] == "Data_Financials(Q)"
    assert sheet["quarter_labels"] == ["FY2026Q1", "FY2026Q2"]
    assert sheet["period_ends"] == ["2026-03-29", "2026-06-28"]
    assert sheet["rows"][0] == {"concept": "Revenue",
                                "label": "Revenues",
                                "values": [100.0, 110.0]}


def test_gaap_needs_an_output(monkeypatch):
    """既不寫 xlsx 也不出 JSON 的話，抓半天沒有任何產出——直接擋掉。"""
    monkeypatch.setattr(cli, "_gaap_tables", lambda **kw: [_fake_table()])
    with pytest.raises(SystemExit):
        cli.main(["gaap", "AAPL", "--identity", "T t@e.com"])


def test_gaap_no_data_returns_nonzero(tmp_path, monkeypatch):
    monkeypatch.setattr(cli, "_gaap_tables", lambda **kw: [])
    rc = cli.main(["gaap", "AAPL", "--xlsx", str(tmp_path / "a.xlsx"),
                   "--identity", "T t@e.com"])
    assert rc != 0


# ── press-release 子指令 ────────────────────────────────────────────────────

class _FakeFiling:
    accession_no = "0001736946-26-000058"
    period_of_report = "2026-08-06"
    filing_date = "2026-08-06"

    def __init__(self, url="https://example.invalid/pr.htm"):
        self.url = url


_HTML = """
<html><body>
<table><tr><td>ARLO TECHNOLOGIES, INC.</td></tr>
       <tr><td>RECONCILIATIONS OF GAAP MEASURES</td></tr></table>
<table>
  <tr><td>GAAP net income</td><td>$</td><td>3028</td><td></td><td>$</td><td>3124</td></tr>
  <tr><td>Stock-based compensation</td><td>21710</td><td>21710</td><td></td><td>14983</td><td>14983</td></tr>
</table>
<table><tr><td>Segment</td><td>1</td></tr></table>
</body></html>
"""


# 真正的新聞稿表頭長這樣：本期、去年同期、半年累計三組日期都在同一張表上。
# 期末日必須挑「不晚於發布日的最新那個」——去年同期比較欄與財測的未來日期都不算。
_HTML_WITH_PERIODS = """
<html><body>
<table><tr><td>ARLO TECHNOLOGIES, INC.</td></tr>
       <tr><td>RECONCILIATIONS OF GAAP MEASURES</td></tr></table>
<table>
  <tr><td></td><td>Three Months Ended June 28, 2026</td><td></td>
      <td>Three Months Ended June 29, 2025</td></tr>
  <tr><td>GAAP net income</td><td>$</td><td>3028</td><td></td><td>$</td><td>3124</td></tr>
</table>
<table>
  <tr><td>Outlook for the quarter ending December 31, 2026</td><td>1</td></tr>
</table>
</body></html>
"""


@pytest.fixture
def fake_pr(monkeypatch):
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q3", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html", lambda filing: _HTML)
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "1231")


@pytest.fixture
def fake_pr_with_periods(monkeypatch):
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q3", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html", lambda filing: _HTML_WITH_PERIODS)
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "1231")


def test_press_release_json_has_filtered_tables(fake_pr, capsys):
    rc = cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    assert rc == 0
    payload = json.loads(capsys.readouterr().out)
    quarter = payload["quarters"][0]
    assert quarter["n_tables_total"] == 2      # 標題區塊已併入下一張表
    assert quarter["n_tables_kept"] == 1       # 只留調節表，segment 那張不留
    rows = quarter["tables"][0]["rows"]
    assert rows[0] == ["GAAP net income", "3028", "3124"]


def test_press_release_reports_the_known_label_offset(fake_pr, capsys):
    """退回舊算法時（label 不是零下載規則算的），仍要帶著 off-by-one 警告。

    這個 fixture 的 label 是寫死的 FY2026Q3，跟零下載規則對同一份申報算出的
    FY2026Q2 不一致 → cli 判定這個 label 不是新規則來的，警告照舊。
    見 docs/8k-period-off-by-one.md。
    """
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["period_of_report"] == "2026-08-06"
    assert quarter["label_source"] == "period_of_report"
    assert "off-by-one" in quarter["label_warning"]


def test_press_release_is_compact(fake_pr, capsys):
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    payload = json.loads(capsys.readouterr().out)
    assert payload["quarters"][0]["chars"] < 3000


def test_press_release_raw_keeps_everything(fake_pr, capsys):
    cli.main(["press-release", "ARLO", "--raw", "--json", "-",
              "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert "tables" not in quarter
    assert "GAAP net income" in quarter["text"]


def test_press_release_text_output(fake_pr, capsys):
    """不給 --json 時輸出人看的純文字。"""
    rc = cli.main(["press-release", "ARLO", "--identity", "T t@e.com"])
    assert rc == 0
    out = capsys.readouterr().out
    assert "FY2026Q3" in out
    assert "GAAP net income | 3028 | 3124" in out


# ── 期末日與正確財季（TODO D4 後半，方案 B+）────────────────────────────────

def test_press_release_reports_the_period_end_from_the_table_header(
        fake_pr_with_periods, capsys):
    """真正的財期結束日在表頭裡，不在 period_of_report。"""
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["period_end"] == "2026-06-28"


def test_press_release_period_end_ignores_prior_year_and_guidance_dates(
        fake_pr_with_periods, capsys):
    """去年同期（2025-06-29）比較欄與財測的未來日期（2026-12-31）都不是本期期末日。"""
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["period_end"] == "2026-06-28"
    assert quarter["fiscal_label"] == "FY2026Q2"


def test_press_release_period_end_survives_a_date_split_across_rows(
        monkeypatch, capsys):
    """NVDA 的表頭把日期拆成上下兩列：`April 26,` 一列、`2026` 下一列。

    只看單一儲存格會整家抓不到期末日（實測 NVDA 三季全空），所以每一欄
    還要**直向串起來**再找一次日期。
    """
    # 中間那個空欄是 Workiva 的期間間隔欄，真實版面就長這樣（少了它，
    # 兩個期間會被 clean_grid 併成同一欄）。
    html = """
    <html><body><table>
      <tr><td></td><td>April 26,</td><td></td><td>April 27,</td></tr>
      <tr><td></td><td>2026</td><td></td><td>2025</td></tr>
      <tr><td>Revenue</td><td>81615</td><td></td><td>44062</td></tr>
    </table></body></html>
    """
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q2", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html", lambda filing: html)
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "0131")  # NVDA
    cli.main(["press-release", "NVDA", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["period_end"] == "2026-04-26"
    assert quarter["fiscal_label"] == "FY2027Q1"


def test_press_release_period_end_beats_a_release_date_in_a_footnote(
        monkeypatch, capsys):
    """AMD 的安全港聲明裡有發布日（2026-08-04），比真正的期末日（06-27）晚。

    「不晚於申報日的最新日期」會挑到發布日，整家標錯一季。所以上限要再往前
    推 3 天——發布日永遠等於申報當天，而財報最快也要期末後兩週才發。
    """
    # 安全港聲明的原句就是 "as of August 4, 2026, and assumptions..."（AMD 實際文字）。
    # `as of` 不可當成期間引導詞，否則它會贏過真正的 `Ended June 27, 2026`。
    html = """
    <html><body>
    <table>
      <tr><td></td><td>Three Months Ended June 27, 2026</td><td></td>
          <td>Three Months Ended June 28, 2025</td></tr>
      <tr><td>Revenue</td><td>11536</td><td></td><td>7685</td></tr>
    </table>
    <table><tr><td>Forward-looking statements speak only as of August 4, 2026,
        and assumptions and estimates may change.</td><td>1</td></tr></table>
    </body></html>
    """
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q3", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html", lambda filing: html)
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "1231")
    cli.main(["press-release", "AMD", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["period_end"] == "2026-06-27"
    assert quarter["fiscal_label"] == "FY2026Q2"


def test_press_release_period_end_beats_a_balance_sheet_comparative_date(
        monkeypatch, capsys):
    """INTC 實際案例：本期期末日（`Jun 27, / 2026`）是**沒有引導詞的表頭**，
    而資產負債表的 `(as of December 27, 2025)` 反而帶關鍵字——那是去年年底。

    這是「優先採信 `ended` / `as of` 後面那個日期」被放棄的原因（實測 AMD／
    INTC／AVGO 三家因此標錯）。現在單純取最新，關鍵字完全不參與判斷。
    """
    html = """
    <html><body>
    <table>
      <tr><td></td><td>Jun 27,</td><td></td><td>Jun 28,</td></tr>
      <tr><td></td><td>2026</td><td></td><td>2025</td></tr>
      <tr><td>Revenue</td><td>13674</td><td></td><td>12833</td></tr>
    </table>
    <table>
      <tr><td>Accumulated other comprehensive income (as of December 27,
          2025)</td><td>416</td></tr>
    </table>
    </body></html>
    """
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q3", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html", lambda filing: html)
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "1231")
    cli.main(["press-release", "INTC", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["period_end"] == "2026-06-27"
    assert quarter["fiscal_label"] == "FY2026Q2"


def test_press_release_keeps_the_legacy_label_untouched(fake_pr_with_periods, capsys):
    """`label` 的值原樣吐出，cli 不會自己改寫列清單給的標籤。"""
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["label"] == "FY2026Q3"           # 發布日換算的舊標籤，仍偏一季
    assert quarter["label_source"] == "period_of_report"
    assert "off-by-one" in quarter["label_warning"]
    assert quarter["fiscal_label_source"] == "period_end"


def test_press_release_without_any_date_leaves_the_fiscal_label_empty(fake_pr, capsys):
    """抓不到期末日時留空，不可退回用發布日硬算——那正是要修的錯誤。"""
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["period_end"] == ""
    assert quarter["fiscal_label"] == ""


def test_press_release_payload_carries_fy_end_month(fake_pr_with_periods, capsys):
    """財季換算靠財年結束月，值要吐出來讓 skill 能自己複算。"""
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    assert json.loads(capsys.readouterr().out)["fy_end_month"] == 12


def test_press_release_fiscal_label_empty_when_fy_end_month_unknown(
        monkeypatch, capsys):
    """財年結束月查不到就不猜。預設 12 會讓非 12 月結算的公司整批標錯。"""
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q3", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html", lambda filing: _HTML_WITH_PERIODS)
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: None)
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    payload = json.loads(capsys.readouterr().out)
    assert payload["fy_end_month"] is None
    assert payload["quarters"][0]["period_end"] == "2026-06-28"   # 期末日照吐
    assert payload["quarters"][0]["fiscal_label"] == ""


def test_press_release_text_output_shows_the_fiscal_label(
        fake_pr_with_periods, capsys):
    rc = cli.main(["press-release", "ARLO", "--identity", "T t@e.com"])
    assert rc == 0
    out = capsys.readouterr().out
    assert "FY2026Q2" in out
    assert "2026-06-28" in out


# 財季換算：這四家是 docs/8k-period-off-by-one.md 實測標錯的案例，
# fy_end_month 取自 EDGAR submissions 的 fiscalYearEnd（實際值）。
@pytest.mark.parametrize("period_end,fy_end_month,expected", [
    ("2026-06-28", 12, "FY2026Q2"),   # ARLO：舊標籤 FY2026Q3
    ("2026-04-26",  1, "FY2027Q1"),   # NVDA：舊標籤 FY2026Q2，偏 −3 季
    ("2026-06-27",  4, "FY2027Q1"),   # QRVO：舊標籤 FY2026Q3，偏 −2 季
    ("2026-05-10",  8, "FY2026Q3"),   # COST：舊標籤 FY2026Q2
])
def test_fiscal_label_matches_the_company_reported_quarter(
        period_end, fy_end_month, expected):
    assert cli._fiscal_label(period_end, fy_end_month) == expected


def test_fiscal_label_survives_a_52_53_week_period_end():
    """WDC FY2026 Q2 結束在 2026-01-02。直接看月份會算成 Q3，整整差一季。

    `fiscal_input.fiscal_quarter_of()` 先把期末日往前推 15 天再取年月，
    52/53 週制浮動最多 6 天，推 15 天必定落回該季最後一個月。
    """
    assert cli._fiscal_label("2026-01-02", 7) == "FY2026Q2"


def test_press_release_skips_quarter_that_fails_to_download(monkeypatch, capsys):
    """一季下載失敗不能拖垮整趟——其餘季照常輸出。"""
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "1231")
    monkeypatch.setattr(cli, "_earnings_filings", lambda **kw: [
        ("FY2026Q3", _FakeFiling()), ("FY2026Q2", _FakeFiling())])

    calls = {"n": 0}

    def _flaky(filing):
        calls["n"] += 1
        if calls["n"] == 1:
            raise OSError("boom")
        return _HTML

    monkeypatch.setattr(cli, "_press_release_html", _flaky)
    rc = cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    payload = json.loads(capsys.readouterr().out)
    assert rc == 0
    assert len(payload["quarters"]) == 1
    assert payload["skipped"] == [{"label": "FY2026Q3", "error": "OSError"}]


def test_press_release_error_message_hides_exception_text(monkeypatch, capsys):
    """例外訊息會挾帶 URL/金鑰，只能記類型。"""
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "1231")
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q3", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html",
                        lambda f: (_ for _ in ()).throw(OSError("secret-key-12345")))
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    captured = capsys.readouterr()
    assert "secret-key-12345" not in captured.out + captured.err


# ── 共用 ────────────────────────────────────────────────────────────────────

def test_force_utf8_io_tolerates_streams_without_reconfigure():
    """pytest 的 capsys 換掉的 stdout 沒有 reconfigure，不可因此拋例外。"""
    import io

    class _NoReconfigure(io.StringIO):
        reconfigure = None

    saved = sys.stdout
    sys.stdout = _NoReconfigure()
    try:
        cli._force_utf8_io()          # 不該拋
    finally:
        sys.stdout = saved


def test_unknown_command_exits(capsys):
    with pytest.raises(SystemExit):
        cli.main(["nonsense", "AAPL"])


def test_identity_falls_back_to_config(monkeypatch, tmp_path):
    monkeypatch.setattr(cli, "load_config", lambda: {"identity": "Cfg c@e.com"})
    monkeypatch.setattr(cli, "_gaap_tables", lambda **kw: [_fake_table()])
    assert cli.resolve_identity(None) == "Cfg c@e.com"


def test_missing_identity_is_an_error(monkeypatch):
    monkeypatch.setattr(cli, "load_config", lambda: {})
    with pytest.raises(cli.CliError):
        cli.resolve_identity(None)


def test_gaap_checks_the_output_file_before_fetching(tmp_path, monkeypatch):
    """檔案被 Excel 開著時要在抓取**之前**就擋下來，不是白抓 24 秒才失敗。"""
    calls = []
    monkeypatch.setattr(cli, "_gaap_tables", lambda **kw: calls.append(kw) or [_fake_table()])
    monkeypatch.setattr(cli, "check_output_writable", lambda p: "檔案正被 Excel 開啟")
    rc = cli.main(["gaap", "AAPL", "--xlsx", str(tmp_path / "a.xlsx"),
                   "--identity", "T t@e.com"])
    assert rc != 0
    assert calls == [], "應該在抓取前就擋下來"


# ── B5：label 改由「發布日 + fiscal_year_end」算 ─────────────────────────────


def test_press_release_passes_the_fiscal_year_end_into_the_listing(monkeypatch, capsys):
    """財年結束日要在列清單**之前**查好並傳下去——`--years` 是在那裡篩的。"""
    seen = {}

    def _fake(**kw):
        seen.update(kw)
        return []

    monkeypatch.setattr(cli, "_earnings_filings", _fake)
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "0703")
    cli.main(["press-release", "WDC", "--json", "-", "--identity", "T t@e.com"])
    assert seen["fiscal_year_end"] == "0703"


def test_press_release_label_source_says_when_the_new_rule_produced_the_label(
        monkeypatch, capsys):
    """label 與零下載規則算出來的一致 → `label_source` 要講清楚是新規則。"""
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q2", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html", lambda filing: _HTML)
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "1231")
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["label"] == "FY2026Q2"
    assert quarter["label_source"] == "announcement+fiscal_year_end"
    assert "off-by-one" not in quarter["label_warning"]
    assert "fiscal_label" in quarter["label_warning"]


def test_press_release_flags_when_the_label_disagrees_with_the_fiscal_label(
        monkeypatch, capsys):
    """下載後的 `fiscal_label` 是最準的。兩邊不一致就標出來——這是唯一能偵測
    `fiscal_year_end` 隨時間改變（規則 C 最大的結構性風險）的訊號。"""
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q2", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html", lambda filing: _HTML_WITH_PERIODS)
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "1231")
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["fiscal_label"] == "FY2026Q2"       # 期末日 2026-06-28
    assert quarter["label_agrees_with_fiscal_label"] is True


def test_press_release_label_agreement_is_none_without_a_fiscal_label(
        fake_pr, capsys):
    """抓不到期末日就沒得比，要吐 null 而不是 false。"""
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["fiscal_label"] == ""
    assert quarter["label_agrees_with_fiscal_label"] is None


def test_press_release_label_agreement_false_when_they_differ(monkeypatch, capsys):
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q3", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html", lambda filing: _HTML_WITH_PERIODS)
    monkeypatch.setattr(cli, "_fiscal_year_end", lambda ticker, identity: "1231")
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    quarter = json.loads(capsys.readouterr().out)["quarters"][0]
    assert quarter["fiscal_label"] == "FY2026Q2"
    assert quarter["label_agrees_with_fiscal_label"] is False


@pytest.mark.parametrize("mmdd, expected", [
    ("1231", 12), ("0703", 7), ("0131", 1), ("0926", 9),
    ("", None), (None, None), ("13xx", None), ("1331", None), ("0000", None),
])
def test_fy_end_month_from_mmdd(mmdd, expected):
    """MMDD → 月份。查不到就回 None，不可以預設 12（非 12 月結算會整批標錯）。"""
    assert cli._fy_end_month_from_mmdd(mmdd) == expected


# ── update-db（TODO J3）─────────────────────────────────────────────────────
#
# 完全離線：抓取那一層是 `local_db.update_local_db()` 的注入點，這裡直接把
# 整個 `update_local_db` 換掉，只驗 CLI 的參數解析、名單維護與輸出格式。

@pytest.fixture
def db_cfg(tmp_path, monkeypatch):
    """把 config.json 與快取根目錄都導到 tmp_path。"""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    import config
    path = tmp_path / "config.json"
    monkeypatch.setattr(config, "CONFIG_PATH", path)
    monkeypatch.setattr(cli, "load_config", lambda: config.load_config(path))
    return path


def _write_cfg(path, data):
    import config
    config.save_config({**config.DEFAULT_CONFIG, **data}, path)


def test_update_db_list_prints_the_update_list(db_cfg, capsys):
    _write_cfg(db_cfg, {"local_db_tickers": ["AAPL", "NVDA"]})
    assert cli.main(["update-db", "--list", "--config-path", str(db_cfg)]) == 0
    assert "AAPL, NVDA" in capsys.readouterr().out


def test_update_db_import_watchlist_writes_config_and_does_not_fetch(db_cfg, capsys):
    """名單維護做完就結束，不順便發動幾小時的抓取——手滑的代價差太多。"""
    _write_cfg(db_cfg, {"watchlist": [{"ticker": "AAPL"}, {"ticker": "NVDA"}]})
    called = []
    import local_db
    _orig = local_db.update_local_db
    local_db.update_local_db = lambda *a, **k: called.append(1)
    try:
        assert cli.main(["update-db", "--import-watchlist",
                         "--config-path", str(db_cfg)]) == 0
    finally:
        local_db.update_local_db = _orig
    assert called == []
    assert json.loads(db_cfg.read_text(encoding="utf-8"))["local_db_tickers"] \
        == ["AAPL", "NVDA"]


def test_update_db_errors_when_the_list_is_empty(db_cfg, capsys):
    _write_cfg(db_cfg, {"identity": "T t@e.com"})
    assert cli.main(["update-db", "--config-path", str(db_cfg)]) == 2
    assert "更新名單是空的" in capsys.readouterr().err


def test_update_db_runs_the_list_and_reports(db_cfg, capsys, monkeypatch):
    _write_cfg(db_cfg, {"identity": "T t@e.com", "local_db_tickers": ["AAPL", "NVDA"]})
    import local_db
    seen = {}

    def fake_update(tickers, identity, **kw):
        seen["tickers"] = list(tickers)
        return local_db.UpdateReport(results=[
            local_db.TickerResult("AAPL", "skipped"),
            local_db.TickerResult("NVDA", "updated", new_filings=3, gaps=1),
        ])

    monkeypatch.setattr(local_db, "update_local_db", fake_update)
    assert cli.main(["update-db", "--config-path", str(db_cfg)]) == 0
    out = capsys.readouterr().out
    assert seen["tickers"] == ["AAPL", "NVDA"]
    assert "更新 1 家、跳過 1 家、失敗 0 家" in out
    assert "NVDA" in out          # 缺漏清單要點名（D11）


def test_update_db_positional_tickers_override_the_list(db_cfg, monkeypatch):
    _write_cfg(db_cfg, {"identity": "T t@e.com", "local_db_tickers": ["AAPL"]})
    import local_db
    seen = {}

    def fake_update(tickers, identity, **kw):
        seen["t"] = list(tickers)
        return local_db.UpdateReport()

    monkeypatch.setattr(local_db, "update_local_db", fake_update)
    assert cli.main(["update-db", "msft", "--config-path", str(db_cfg)]) == 0
    assert seen["t"] == ["MSFT"]


def test_update_db_returns_nonzero_when_a_company_failed(db_cfg, monkeypatch):
    _write_cfg(db_cfg, {"identity": "T t@e.com", "local_db_tickers": ["AAPL"]})
    import local_db
    monkeypatch.setattr(local_db, "update_local_db",
                        lambda *a, **k: local_db.UpdateReport(results=[
                            local_db.TickerResult("AAPL", "failed", error="X: boom")]))
    assert cli.main(["update-db", "--config-path", str(db_cfg)]) == 1
