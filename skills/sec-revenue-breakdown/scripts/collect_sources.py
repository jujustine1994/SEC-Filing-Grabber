"""Create a pinned, hash-verified SEC source pack. Standard library only."""
import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import time
from html.parser import HTMLParser
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

VERSION = "1.1"


def sha(data):
    return hashlib.sha256(data).hexdigest()


def write_json(path, value):
    atomic(path, json.dumps(value, ensure_ascii=False, indent=2).encode("utf-8"))


def atomic(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_bytes(data)
    temp.replace(path)


def sec_url(url):
    parsed = urlparse(url)
    if parsed.scheme != "https" or parsed.hostname not in {"www.sec.gov", "data.sec.gov"}:
        raise ValueError("Only HTTPS SEC official URLs are accepted")
    return url


class Client:
    def __init__(self, identity):
        if "@" not in identity:
            raise ValueError("Configure SEC_IDENTITY or SEC Financial Tools identity with contact email")
        self.identity = identity
        self.last = 0

    def get(self, url):
        sec_url(url)
        for attempt in range(3):
            time.sleep(max(0, .25 - (time.monotonic() - self.last)))
            self.last = time.monotonic()
            try:
                with urlopen(Request(url, headers={"User-Agent": self.identity}), timeout=45) as response:
                    sec_url(response.url)
                    data = response.read()
                if not data or b"Your Request Originates from an Undeclared Automated Tool" in data:
                    raise ValueError("Empty or blocked SEC response")
                return data
            except HTTPError as exc:
                if exc.code not in {429, 500, 502, 503, 504} or attempt == 2:
                    raise RuntimeError(f"SEC HTTP {exc.code}: {url}") from None
            except (URLError, TimeoutError):
                if attempt == 2:
                    raise RuntimeError(f"SEC connection failed: {url}") from None
            time.sleep(2 ** (attempt + 1))


class IndexParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self.cells, self.links = [], [], []
        self.cell = None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self.cells, self.links = [], []
        elif tag == "td":
            self.cell = ""
        elif tag == "a" and self.cell is not None:
            href = dict(attrs).get("href")
            if href:
                self.links.append(href)

    def handle_data(self, data):
        if self.cell is not None:
            self.cell += data

    def handle_endtag(self, tag):
        if tag == "td" and self.cell is not None:
            self.cells.append(self.cell.strip())
            self.cell = None
        elif tag == "tr" and self.cells:
            self.rows.append((self.cells, self.links))


def exhibits(html, index_url):
    parser = IndexParser()
    parser.feed(html)
    result = []
    for cells, links in parser.rows:
        kind = next((cell for cell in cells if re.fullmatch(r"EX-99(?:\.\d+)?", cell)), None)
        if kind and links:
            url = sec_url(urljoin(index_url, links[0]))
            result.append((kind, url))
    return result


def rows(columnar):
    return [dict(zip(columnar, values)) for values in zip(*columnar.values())]


def inventory(client, cik, start, end):
    base = "https://data.sec.gov/submissions/"
    data = json.loads(client.get(f"{base}CIK{cik:010d}.json"))
    filings = rows(data["filings"]["recent"])
    for page in data["filings"].get("files", []):
        if page["filingTo"] >= start and page["filingFrom"] <= end:
            filings.extend(rows(json.loads(client.get(base + page["name"]))))
    chosen = {}
    for filing in filings:
        form = filing["form"]
        if start <= filing["filingDate"] <= end and (
            form in {"10-Q", "10-K", "10-Q/A", "10-K/A", "10-K405", "10-K405/A"}
            or (form in {"8-K", "8-K/A"} and (
                "2.02" in filing.get("items", "").split(",")
                # Legacy earnings releases used items 9/12, and metadata is
                # inconsistent. Retain all legacy 8-K as candidates, then inspect.
                or filing["filingDate"] < "2004-08-23"))
        ):
            chosen[filing["accessionNumber"]] = filing
    return dict(company_name=data["name"], tickers=data.get("tickers", []),
                filings=sorted(chosen.values(), key=lambda f: (f["filingDate"], f["accessionNumber"])))


def collect(client, cik, start, end, out, refresh=False):
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    selection = out / "selection.json"
    if selection.exists() and not refresh:
        selected = json.loads(selection.read_text(encoding="utf-8"))
        if (selected["cik"], selected["start"], selected["end"]) != (cik, start, end):
            raise ValueError("Pinned selection differs; use another output directory or --refresh")
    else:
        selected = dict(cik=cik, start=start, end=end, selected_at=dt.datetime.now(dt.timezone.utc).isoformat(),
                        **inventory(client, cik, start, end))
        if not selected["filings"]:
            raise ValueError("No relevant SEC filings in range")
        write_json(selection, selected)
    manifest_path = out / "manifest.json"
    old = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else {"sources": []}
    previous = {item["url"]: item for item in old["sources"]}
    sources = []

    def fetch(url, filing, kind):
        url = sec_url(url)
        name = urlparse(url).path.rsplit("/", 1)[1]
        if not re.fullmatch(r"[A-Za-z0-9_.-]+", name):
            raise ValueError("Unsafe SEC filename")
        relative = f"raw/{filing['accessionNumber']}/{name}"
        path = out / relative
        cached = previous.get(url)
        if cached and path.exists() and sha(path.read_bytes()) == cached["sha256"]:
            data = path.read_bytes()
            downloaded = cached["downloaded_at"]
        else:
            data = client.get(url)
            atomic(path, data)
            downloaded = dt.datetime.now(dt.timezone.utc).isoformat()
        sources.append(dict(source_id=f"{filing['accessionNumber']}/{name}", url=url, filename=relative,
                            accession=filing["accessionNumber"], form=filing["form"], document_type=kind,
                            filing_date=filing["filingDate"], report_date=filing.get("reportDate"),
                            period_assignment="unverified", sha256=sha(data), bytes=len(data), downloaded_at=downloaded))
        # Preserve validated cache entries even if a later download fails.
        checkpoint = {**previous, **{s["url"]: s for s in sources}}
        write_json(manifest_path, dict(schema_version="1.0", status="incomplete",
                                      sources=list(checkpoint.values())))
        return data

    for filing in selected["filings"]:
        acc = filing["accessionNumber"]
        base = f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc.replace('-', '')}/"
        index_url = base + acc + "-index.html"
        index = fetch(index_url, filing, "filing-index")
        # Legacy submissions omit primaryDocument. Their full submission text is
        # the authoritative source, not a guessed modern filename.
        primary = filing.get("primaryDocument") or acc + ".txt"
        try:
            fetch(base + primary, filing, filing["form"])
        except RuntimeError as exc:
            if not str(exc).startswith("SEC HTTP 404:") or primary == acc + ".txt":
                raise
            fetch(base + acc + ".txt", filing, filing["form"])
        for kind, url in exhibits(index.decode("utf-8", errors="replace"), index_url):
            fetch(url, filing, kind)
        print(f"Collected {filing['filingDate']} {filing['form']} {acc}", flush=True)
    result = dict(schema_version="1.0", status="complete", collector_version=VERSION, cik=cik,
                  company_name=selected["company_name"], start=start, end=end,
                  selection_sha256=sha(selection.read_bytes()), sources=sources,
                  warnings=["Financial period and requested coverage require document-level verification.",
                            "Legacy 8-K are candidates; their presence does not prove an earnings release."])
    write_json(manifest_path, result)
    lines = [f"# {selected['company_name']} SEC sources", "",
             f"CIK: {cik}. Filing dates: {start} through {end}.", "",
             "Financial periods and classification must be verified from the documents. "
             "An 8-K report date is not necessarily the quarter-end date.", "",
             "| Filed | Form | Document | SEC official source | Local original |",
             "|---|---|---|---|---|"]
    for source in sources:
        lines.append(f"| {source['filing_date']} | {source['form']} | {source['document_type']} | "
                     f"[{source['source_id']}]({source['url']}) | [raw]({source['filename']}) |")
    atomic(out / "SOURCES.md", ("\n".join(lines) + "\n").encode("utf-8"))
    return result


def identity():
    if os.environ.get("SEC_IDENTITY"):
        return os.environ["SEC_IDENTITY"]
    path = (Path(os.environ["APPDATA"]) / "SEC Financial Tools/config.json"
            if os.environ.get("APPDATA") else Path.home() / ".sec_financial_tools/config.json")
    return json.loads(path.read_text(encoding="utf-8")).get("identity", "") if path.exists() else ""


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cik", type=int, required=True)
    parser.add_argument("--start", required=True, help="Filing date YYYY-MM-DD")
    parser.add_argument("--end", required=True, help="Filing date YYYY-MM-DD")
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--refresh", action="store_true")
    args = parser.parse_args()
    if not 0 < args.cik < 10**10 or dt.date.fromisoformat(args.start) > dt.date.fromisoformat(args.end):
        parser.error("Invalid CIK or date range")
    result = collect(Client(identity()), args.cik, args.start, args.end, args.out, args.refresh)
    print(f"Saved {len(result['sources'])} SEC documents; period assignment requires verification.")


if __name__ == "__main__":
    main()
