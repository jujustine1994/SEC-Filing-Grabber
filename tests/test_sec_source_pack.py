"""Observable source-pack invariants without SEC network requests."""
import importlib.util
import json
from pathlib import Path

import pytest


spec = importlib.util.spec_from_file_location(
    "sec_source_pack", Path(__file__).parents[1] / "skills/sec-revenue-breakdown/scripts/collect_sources.py"
)
pack = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pack)


def test_pinned_rerun_and_corruption_repair(tmp_path, monkeypatch):
    filing = dict(accessionNumber="0001045810-26-000001", filingDate="2026-01-01",
                  reportDate="2025-12-31", primaryDocument="report.htm", form="8-K")
    monkeypatch.setattr(pack, "inventory", lambda *args: dict(company_name="Example", tickers=["EX"], filings=[filing]))

    class Client:
        def __init__(self):
            self.calls = []

        def get(self, url):
            self.calls.append(url)
            if url.endswith("-index.html"):
                return b'<table><tr><td>2</td><td>CFO</td><td><a href="cfo.htm">cfo.htm</a></td><td>EX-99.2</td></tr></table>'
            return b"<html>Revenue evidence</html>"

    client = Client()
    first = pack.collect(client, 1045810, "2026-01-01", "2026-02-01", tmp_path)
    assert len(first["sources"]) == 3
    assert any(s["document_type"] == "EX-99.2" for s in first["sources"])
    client.calls.clear()
    second = pack.collect(client, 1045810, "2026-01-01", "2026-02-01", tmp_path)
    assert first == second
    assert client.calls == []
    source = first["sources"][-1]
    (tmp_path / source["filename"]).write_bytes(b"corrupt")
    repaired = pack.collect(client, 1045810, "2026-01-01", "2026-02-01", tmp_path)
    assert client.calls == [source["url"]]
    assert repaired["sources"][-1]["sha256"] == source["sha256"]
    with pytest.raises(ValueError, match="Pinned selection differs"):
        pack.collect(client, 1045810, "2025-01-01", "2026-02-01", tmp_path)


def test_exhibit_source_must_be_official_sec():
    with pytest.raises(ValueError, match="SEC official"):
        pack.exhibits('<tr><td><a href="https://example.com/cfo.htm">CFO</a></td><td>EX-99.2</td></tr>',
                      "https://www.sec.gov/Archives/index.html")


def test_selection_includes_amendments_and_earnings_only():
    data = {"name": "Example", "filings": {"recent": {
        "accessionNumber": ["a", "b", "c", "d"],
        "form": ["10-Q/A", "8-K", "8-K", "10-Q"],
        "items": ["", "2.02,9.01", "5.02", ""],
        "filingDate": ["2026-01-01"] * 3 + ["2025-01-01"],
    }}}

    class Client:
        def get(self, url):
            return json.dumps(data).encode()

    assert [f["accessionNumber"] for f in pack.inventory(Client(), 1, "2026-01-01", "2026-02-01")["filings"]] == ["a", "b"]


def test_local_context_keeps_dimensions_and_rejects_other_company(tmp_path, monkeypatch, isolated_database):
    monkeypatch.syspath_prepend(str(Path(__file__).parents[1] / "skills/sec-revenue-breakdown/scripts"))
    import local_context
    cache = isolated_database / "filings/EX"
    cache.mkdir(parents=True)
    entry = dict(cik=1, accession_no="0000000001-26-000001", form="10-Q", filing_date="2026-01-01",
                 dataframes={"income_statement": {"data": {
                     "columns": ["concept", "dimension_axis", "2025-12-31 (Q4)"],
                     "data": [["us-gaap_Revenues", "ProductAxis", 1000],
                              ["us-gaap_OperatingIncomeLoss", None, 100],
                              ["us-gaap_IncomeTaxExpenseBenefit", None, 10]]}}})
    (cache / "0000000001-26-000001.json").write_text(json.dumps(entry), encoding="utf-8")
    result = local_context.export(tmp_path, "EX", "2026-01-01", "2026-02-01", tmp_path / "context.json")
    extracted = result["local_filings"][0]["income_statement_rows"]
    assert len(extracted) == 2
    assert extracted[0]["dimension_axis"] == "ProductAxis"
    source = tmp_path / "source"
    source.mkdir()
    (source / "manifest.json").write_text(json.dumps(dict(status="complete", cik=2, sources=[])), encoding="utf-8")
    with pytest.raises(ValueError, match="CIK differs"):
        local_context.export(tmp_path, "EX", "2026-01-01", "2026-02-01", tmp_path / "context.json", source)
