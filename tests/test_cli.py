"""cli.py 的測試（TODO B1）。

全部離線：網路那一層集中在 `cli._gaap_tables` / `cli._earnings_filings` /
`cli._press_release_html` 三個函式，測試把它們換掉。這也是刻意的設計——
CLI 只做參數解析與輸出格式化，抓取邏輯一律轉呼叫既有核心函式。
"""
from __future__ import annotations

import json

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


@pytest.fixture
def fake_pr(monkeypatch):
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q3", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html", lambda filing: _HTML)


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
    """季度標籤來自 period_of_report（＝發布日），已知晚一季。

    下游 skill 不能無條件相信這個標籤，所以每一季都要帶著原始日期與警告。
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


def test_press_release_skips_quarter_that_fails_to_download(monkeypatch, capsys):
    """一季下載失敗不能拖垮整趟——其餘季照常輸出。"""
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
    monkeypatch.setattr(cli, "_earnings_filings",
                        lambda **kw: [("FY2026Q3", _FakeFiling())])
    monkeypatch.setattr(cli, "_press_release_html",
                        lambda f: (_ for _ in ()).throw(OSError("secret-key-12345")))
    cli.main(["press-release", "ARLO", "--json", "-", "--identity", "T t@e.com"])
    captured = capsys.readouterr()
    assert "secret-key-12345" not in captured.out + captured.err


# ── 共用 ────────────────────────────────────────────────────────────────────

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
