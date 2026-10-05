"""Download official NVIDIA CFO commentaries for an isolated eight-quarter pilot."""
import hashlib
import json
from pathlib import Path

import requests
import argparse

ROOT = Path(__file__).resolve().parents[1] / "output" / "nvda_revenue_pilot"
PERIODS = [(2025, 3), (2025, 4), (2026, 1), (2026, 2),
           (2026, 3), (2026, 4), (2027, 1), (2027, 2)]


def download():
    ROOT.mkdir(parents=True, exist_ok=True)
    manifest = []
    session = requests.Session()
    for year, quarter in PERIODS:
        label = f"FY{year}Q{quarter}"
        name = f"Q{quarter}FY{str(year)[-2:]}-CFO-Commentary.pdf"
        url = (f"https://s201.q4cdn.com/141608511/files/doc_financials/"
               f"{year}/Q{quarter}{str(year)[-2:]}/{name}")
        path = ROOT / name
        if not path.exists():
            response = session.get(url, timeout=40)
            response.raise_for_status()
            if not response.content.startswith(b"%PDF"):
                raise ValueError(f"Not a PDF: {label}")
            path.write_bytes(response.content)
        data = path.read_bytes()
        manifest.append(dict(period=label, url=url, filename=name, bytes=len(data),
                             sha256=hashlib.sha256(data).hexdigest()))
        print(label, len(data), flush=True)
    (ROOT / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")


def download_filings():
    filings = [
        ("FY2025Q3", "000104581024000316", "20241027"),
        ("FY2025Q4", "000104581025000023", "20250126"),
        ("FY2026Q1", "000104581025000116", "20250427"),
        ("FY2026Q2", "000104581025000209", "20250727"),
        ("FY2026Q3", "000104581025000230", "20251026"),
        ("FY2026Q4", "000104581026000021", "20260125"),
        ("FY2027Q1", "000104581026000052", "20260426"),
        ("FY2027Q2", "000104581026000075", "20260726"),
    ]
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
    from config import load_config
    identity = load_config().get("identity", "")
    session = requests.Session()
    session.headers["User-Agent"] = identity or "SEC Financial Tools research"
    manifest = []
    for label, accession, end in filings:
        url = f"https://www.sec.gov/Archives/edgar/data/1045810/{accession}/nvda-{end}.htm"
        path = ROOT / f"{label}-filing.html"
        if not path.exists():
            response = session.get(url, timeout=40)
            print(label, response.status_code, flush=True)
            response.raise_for_status()
            path.write_bytes(response.content)
        manifest.append(dict(period=label, url=url, filename=path.name, period_end=end))
    (ROOT / "filings_manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--filings", action="store_true")
    args = parser.parse_args()
    download_filings() if args.filings else download()
