"""Export local SEC Financial Tools evidence for AI; no network or model calls."""
import argparse
import json
from pathlib import Path
import re

from collect_sources import atomic, sha, write_json


def export(repo, ticker, start, end, out, source_pack=None, supplements=()):
    repo = Path(repo).resolve()
    ticker = ticker.upper()
    if not re.fullmatch(r"[A-Z0-9][A-Z0-9.-]{0,19}", ticker):
        raise ValueError("Invalid ticker")
    out = Path(out).resolve()
    entries, warnings = [], []
    for path in sorted((repo / "local_db/filing_cache" / ticker).glob("*.json")):
        if not re.fullmatch(r"\d{10}-\d{2}-\d{6}", path.stem):
            continue
        data = path.read_bytes()
        entry = json.loads(data)
        if not start <= entry.get("filing_date", "") <= end:
            continue
        payload = (entry.get("dataframes") or {}).get("income_statement") or {}
        table = payload.get("data") or {}
        columns = table.get("columns", [])
        picked = []
        for values in table.get("data", []):
            row = dict(zip(columns, values))
            # Match actual concepts, not arbitrary dimension labels or growth rates.
            concept = str(row.get("concept", "")).lower()
            if any(term in concept for term in ("revenues", "salesrevenue", "revenuefromcontract", "operatingincome")):
                picked.append(row)
        cik = entry.get("cik")
        accession = entry.get("accession_no", path.stem)
        url = (f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/"
               f"{accession.replace('-', '')}/{accession}-index.html") if cik else None
        entries.append(dict(source_id=f"local-cache/{ticker}/{path.stem}", path=str(path),
                            sha256=sha(data), source_kind="derived_local_xbrl", sec_index_url=url,
                            accession=accession, cik=cik, form=entry.get("form"),
                            filing_date=entry.get("filing_date"), cached_at=entry.get("cached_at"),
                            schema_version=entry.get("schema_version"), parser_version=entry.get("edgartools_version"),
                            units="Verify XBRL unit before comparison; do not assume USD millions",
                            income_statement_rows=picked))
        if not picked:
            warnings.append(f"No revenue/OP candidates in {path.stem}")
    if len({e["cik"] for e in entries}) > 1:
        raise ValueError("Local cache contains different CIKs; resolve before combining")
    source_evidence = None
    if source_pack:
        manifest_path = Path(source_pack).resolve() / "manifest.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        if manifest.get("status") != "complete":
            raise ValueError("SEC source pack is incomplete")
        if entries and manifest.get("cik") != entries[0]["cik"]:
            raise ValueError("SEC source pack CIK differs from local cache")
        for source in manifest["sources"]:
            raw = (manifest_path.parent / source["filename"]).resolve()
            if not raw.is_relative_to(manifest_path.parent) or sha(raw.read_bytes()) != source["sha256"]:
                raise ValueError("SEC source pack path or hash verification failed")
        source_evidence = dict(path=str(manifest_path), sha256=sha(manifest_path.read_bytes()),
                               source_count=len(manifest["sources"]), status="hash_verified")
    supplemental = []
    for item in supplements:
        path = Path(item).resolve()
        if path == out:
            raise ValueError("Supplement must not be the output file")
        supplemental.append(dict(path=str(path), sha256=sha(path.read_bytes()),
                                 source_kind="supplemental_unverified",
                                 instruction="Verify company, period, provenance and basis before merging"))
    result = dict(schema_version="1.0", ticker=ticker, repo=str(repo),
                  requested_filing_dates=dict(start=start, end=end), local_filings=entries,
                  sec_source_pack=source_evidence, supplements=supplemental, warnings=warnings,
                  period_assignment="unverified",
                  merge_rule="Keep sources separate; reconcile units, financial periods, dimension and recast basis. "
                             "Resolve conflicts from original disclosure; never overwrite silently.")
    write_json(out, result)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, required=True)
    parser.add_argument("--ticker", required=True)
    parser.add_argument("--start", required=True)
    parser.add_argument("--end", required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--source-pack", type=Path)
    parser.add_argument("--supplement", type=Path, action="append", default=[])
    args = parser.parse_args()
    result = export(args.repo, args.ticker, args.start, args.end, args.out, args.source_pack, args.supplement)
    print(f"Exported {len(result['local_filings'])} local filing candidates; no AI API or network calls.")


if __name__ == "__main__":
    main()
