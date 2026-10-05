"""Extract NVDA historical revenue evidence, preserve versions and reconcile groups."""
import collections
import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
import re

from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parents[1]
PACK = ROOT / "output/NVDA_history_sources"
OUT = ROOT / "output/NVDA_segment_history"
LOCAL = None
PILOT = Path.home() / "Documents/Code/SEC Financial Tools/output/nvda_revenue_pilot/dataset.json"
VERSION = "nvda-history-1.0"
DATE = re.compile(r"(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+\d{1,2},?\s+\d{4}", re.I)


def clean(text):
    return " ".join(str(text).replace("\u200b", "").split())


def date_value(text):
    text = text.replace(".", "").replace(",", "")
    for fmt in ("%B %d %Y", "%b %d %Y"):
        try:
            return dt.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def period(end, duration):
    d = dt.date.fromisoformat(end)
    fy = d.year if d.month <= 2 else d.year + 1
    if duration == "annual":
        return f"FY{fy}", fy, None
    if duration != "quarter":
        return f"FY{fy}YTD{duration}", fy, None
    quarter = {1: 4, 2: 4, 4: 1, 5: 1, 7: 2, 8: 2, 10: 3, 11: 3}.get(d.month)
    return (f"FY{fy}Q{quarter}", fy, quarter) if quarter else (None, fy, None)


def amounts(cells):
    result = []
    negative = False
    for cell in cells:
        text = cell.replace("$", "").replace(",", "").replace(" ", "").strip()
        if text == "(":
            negative = True
            continue
        if text == ")":
            negative = False
            continue
        if re.fullmatch(r"\(?-?\d+(?:\.\d+)?\)?", text):
            result.append(-float(text.strip("()")) if text.startswith("(") or negative else float(text))
            negative = False
        elif text in {"—", "–", "-"}:
            result.append(0.0)
    return result


def unit_of(table, document):
    text = clean(table.get_text(" ", strip=True)).lower()
    # Closest preceding explicit unit takes precedence over document defaults.
    before = []
    for item in table.previous_elements:
        if isinstance(item, str) and clean(item):
            before.append(clean(item))
        if len(before) >= 60:
            break
    nearby = text + " " + " ".join(before)
    if "in millions" in text:
        return 1.0
    if "in thousands" in text:
        return .001
    for item in before:
        if "in millions" in item.lower():
            return 1.0
        if "in thousands" in item.lower():
            return .001
    return .001 if document.lower().count("in thousands") > document.lower().count("in millions") else 1.0


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    manifest = json.loads((PACK / "manifest.json").read_text(encoding="utf-8"))
    if manifest.get("status") != "complete":
        raise ValueError("Source pack is incomplete")
    sources = {s["source_id"]: s for s in manifest["sources"]}
    for source in sources.values():
        path = PACK / source["filename"]
        if hashlib.sha256(path.read_bytes()).hexdigest() != source["sha256"]:
            raise ValueError(f"Source hash changed: {source['source_id']}")
    records, candidates, checks = [], [], []

    def add(source, dimension, node, value, end, duration, definition, locator, metric="revenue", parent=None, derivation=None):
        p, fy, fq = period(end, duration)
        if p is None:
            return
        records.append(dict(id=f"r{len(records)+1}", fiscal_year=fy, fiscal_quarter=fq, period=p,
                            period_start=None, period_end=end, duration=duration, metric=metric, value=round(value, 6),
                            currency="USD", unit="USD millions", basis="segment-reported" if dimension.startswith("reportable_segment") else "GAAP",
                            dimension=dimension, node_id=node, original_label=node, display_label=node, parent_id=parent,
                            classification_version=definition, presentation="as_reported" if source.get("report_date") == end else "comparative_as_disclosed",
                            published_in_accession=source.get("accession"), filing_date=source.get("filing_date"),
                            source=dict(source_id=source["source_id"], url=source["url"], locator=locator),
                            derivation=derivation or dict(method="direct"), status="provisional"))

    primary = [s for s in sources.values() if s["document_type"] in {"10-K", "10-K405", "10-K405/A", "10-Q", "10-Q/A", "10-K/A"}]
    for number, source in enumerate(primary):
        path = PACK / source["filename"]
        if hashlib.sha256(path.read_bytes()).hexdigest() != source["sha256"]:
            raise ValueError("Source hash changed")
        soup = BeautifulSoup(path.read_bytes(), "html.parser")
        full_text = clean(soup.get_text(" ", strip=True))
        for index, table in enumerate(soup.find_all("table")):
            text = clean(table.get_text(" ", strip=True))
            if len(text) > 14000 or "revenue" not in text.lower():
                continue
            if not any(word in text for word in ("GPU", "MCP", "Tegra", "Graphics", "Gaming", "Taiwan", "Asia Pacific", "Hyperscale", "Royalty", "Cost of revenue")):
                continue
            lines = [[clean(c.get_text(" ", strip=True)) for c in row.find_all(["td", "th"], recursive=False)] for row in table.find_all("tr")]
            candidates.append(dict(source_id=source["source_id"], table_index=index, text=text, rows=lines))
            if not lines and "Cost of revenue" in text and "Revenue" in text:
                original = table.get_text(" ", strip=True)
                financial = {}
                for match in re.finditer(r"(?m)(?:^|\n| )\s*(Product|Royalty|Total revenue|Revenue)\.+\s*([^\n]+)", original):
                    financial[match[1]] = amounts([v.replace("--", "-") for v in re.findall(r"\(?\$?[\d,]+\)?|--", match[2])])
                control_name = "Total revenue" if "Total revenue" in financial else "Revenue"
                if control_name in financial:
                    head = clean(original.split("Revenue", 1)[0])
                    years = re.findall(r"\b(?:19|20)\d{2}\b", head)
                    months = re.findall(r"(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?", head)
                    ends, durations, overrides = [], [], []
                    if "Three Months" in head and len(months) == len(years):
                        ends = [date_value(f"{m} {y}") for m,y in zip(months,years)]
                        durations = ["quarter"] * len(ends)
                        if len(ends) == 4 and ("Six Months" in head or "Nine Months" in head):
                            durations[2:] = ["9" if "Nine Months" in head else "6"] * 2
                    elif "Year" in head:
                        for year in years:
                            fy = int(year)
                            if fy <= 1997:
                                ends.append(f"{fy}-12-31");durations.append("annual")
                            elif fy == 1998 and "Month Ended" in head:
                                ends.append("1998-01-31");durations.append("transition")
                            else:
                                report = next((s for s in primary if s.get("report_date", "").startswith(str(fy)) and s["form"].startswith("10-K")), None)
                                ends.append(report["report_date"] if report else None);durations.append("annual")
                            overrides.append(fy)
                    if ends and all(ends) and all(len(v) == len(ends) for v in financial.values()):
                        factor = unit_of(table, full_text)
                        split = "Product" in financial and "Royalty" in financial
                        for col,end in enumerate(ends):
                            delta = financial["Product"][col]+financial["Royalty"][col]-financial[control_name][col] if split else 0
                            if abs(delta) > .01:
                                continue
                            start = len(records)
                            for name,vals in financial.items():
                                add(source, "consolidated" if name == control_name else "revenue_type",
                                    "Total revenue" if name == control_name else name, vals[col]*factor, end,
                                    durations[col], "revenue_type:Product/Royalty" if split else "consolidated:GAAP",
                                    f"SGML table {index}, {name}, column {col+1}")
                                if overrides:
                                    records[-1]["fiscal_year"] = overrides[col]
                                    records[-1]["period"] = f"FY{overrides[col]}" + ("Transition" if durations[col] == "transition" else "")
                                records[-1]["status"] = "verified"
                            checks.append(dict(source_id=source["source_id"], period=records[-1]["period"], metric="revenue_type" if split else "direct_consolidated",
                                               difference=delta*factor, status="pass", record_ids=[r["id"] for r in records[start:]]))
            # Geographic tables are vertical. Keep annual and transition periods
            # separate; early NVIDIA fiscal calendars differ from today's calendar.
            if any(word in text for word in ("Asia Pacific", "Taiwan", "U.S.")) and "Total revenue" in text:
                geo_names = {"U.S.", "United States", "U.S. and North America", "United States and other Americas", "Other Americas", "Asia Pacific", "Other Asia Pacific", "Europe", "China", "China (including Hong Kong)", "Taiwan", "Other countries", "Other", "Total revenue"}
                geo_data = {}
                ends, durations, fiscal_override = [], [], []
                if lines:
                    header_lines = []
                    for line in lines:
                        cells = [c for c in line if c]
                        if cells and cells[0] in geo_names:
                            vals = amounts(cells[1:])
                            if vals:
                                geo_data[cells[0]] = vals
                            if cells[0] == "Total revenue":
                                break
                        elif not geo_data:
                            header_lines.append(line)
                    dates = [date_value(m) for line in header_lines for c in line for m in DATE.findall(c)]
                    if not dates:
                        partial = [c for line in header_lines for c in line if re.fullmatch(r"[A-Za-z]+\.? \d{1,2},?", c)]
                        years = [c for line in header_lines for c in line if re.fullmatch(r"(?:19|20)\d{2}", c)]
                        if len(partial) == len(years):
                            dates = [date_value(f"{m} {y}") for m,y in zip(partial, years)]
                    ends = dates
                    header_text = " ".join(" ".join(line) for line in header_lines)
                    dur = "quarter" if "Three Months" in header_text else "9" if "Nine Months" in header_text else "6" if "Six Months" in header_text else "annual" if "Year" in header_text else None
                    durations = [dur] * len(ends)
                    if dur == "quarter" and len(ends) == 4:
                        second = "9" if "Nine Months" in header_text else "6" if "Six Months" in header_text else None
                        durations = ["quarter", "quarter", second, second]
                else:
                    original = table.get_text(" ", strip=True)
                    # Legacy SGML fixed-width tables have no tr/td elements.
                    for match in re.finditer(r"(U\.S\.|Europe|Asia Pacific|Total revenue)\.+\s*([^\n]+)", original):
                        geo_data[match[1]] = amounts([v.replace("--", "-") for v in re.findall(r"\$?[\d,]+|--", match[2])])
                    years = re.findall(r"\b(?:19|20)\d{2}\b", original.split("U.S.")[0])
                    for year in years:
                        fy = int(year)
                        if fy in {1996, 1997}:
                            ends.append(f"{fy}-12-31")
                            durations.append("annual")
                        elif fy == 1998:
                            ends.append("1998-01-31")
                            durations.append("transition")
                        else:
                            match_source = next((s for s in primary if s.get("report_date", "").startswith(str(fy)) and s["form"].startswith("10-K")), None)
                            ends.append(match_source["report_date"] if match_source else None)
                            durations.append("annual")
                        fiscal_override.append(fy)
                if geo_data and "Total revenue" in geo_data and ends and all(ends) and all(durations):
                    names = sorted(n for n in geo_data if n != "Total revenue")
                    definition = "geography:" + "/".join(names)
                    factor = unit_of(table, full_text)
                    for col, end in enumerate(ends):
                        if any(len(v) != len(ends) for v in geo_data.values()):
                            continue
                        delta = sum(geo_data[n][col] for n in names)-geo_data["Total revenue"][col]
                        if abs(delta) > .01:
                            continue
                        start = len(records)
                        for name, values in geo_data.items():
                            add(source, "consolidated" if name == "Total revenue" else "geography", name,
                                values[col]*factor, end, durations[col], definition, f"HTML/SGML table {index}, {name}, column {col+1}")
                            if fiscal_override:
                                records[-1]["fiscal_year"] = fiscal_override[col]
                                records[-1]["period"] = f"FY{fiscal_override[col]}" + ("Transition" if durations[col] == "transition" else "")
                            records[-1]["status"] = "verified"
                        checks.append(dict(source_id=source["source_id"], period=records[-1]["period"], metric="geography_revenue",
                                           difference=delta*factor, status="pass", record_ids=[r["id"] for r in records[start:]]))
            # Reportable segment tables place segments horizontally, metrics vertically.
            header = next(([c for c in line if c] for line in lines
                           if any(c in {"GPU", "Tegra Processor", "Compute & Networking"} for c in line)
                           and any(c in {"Consolidated", "Total"} for c in line)), None)
            if not header:
                continue
            header = [c for c in header if not re.fullmatch(r"\(?\d+\)?", c)]
            definition = "segments:" + "/".join(header)
            factor = unit_of(table, full_text)
            current = None
            for line in lines:
                joined = " ".join(line)
                dates = DATE.findall(joined)
                if dates:
                    if re.search(r"Three\s+Months?", joined, re.I):
                        duration = "quarter"
                    elif re.search(r"Nine\s+Months?", joined, re.I):
                        duration = "9"
                    elif re.search(r"Six\s+Months?", joined, re.I):
                        duration = "6"
                    elif re.search(r"Year\s+Ended", joined, re.I):
                        duration = "annual"
                    else:
                        current = None
                        continue
                    current = (date_value(dates[0]), duration)
                cells = [c for c in line if c]
                if not current or not cells:
                    continue
                metric = ("revenue" if cells[0].lower() in {"revenue", "revenues"} else
                          "operating_income" if cells[0].lower().startswith("operating income") else None)
                if not metric:
                    continue
                values = amounts(cells[1:])
                if len(values) != len(header):
                    continue
                start = len(records)
                for name, value in zip(header, values):
                    total = name in {"Consolidated", "Total"}
                    add(source, ("consolidated" if metric == "revenue" else "reportable_segment_total") if total else "reportable_segment",
                        ("Total revenue" if metric == "revenue" else "Total segment operating income") if total else name,
                        value * factor, *current, definition, f"HTML table {index}, row {joined[:150]}", metric)
                delta = round(sum(values[:-1]) - values[-1], 6)
                # Revenue must reconcile; segment OP total may have a different basis.
                check = dict(source_id=source["source_id"], period=period(*current)[0], metric=metric,
                             difference=delta * factor, status="pass" if abs(delta) < .01 else "fail",
                             record_ids=[r["id"] for r in records[start:]])
                checks.append(check)
                if check["status"] == "pass":
                    for record in records[start:]:
                        record["status"] = "verified"
        if number % 15 == 0:
            print(f"Parsed {number+1}/{len(primary)} primary filings", flush=True)

    # Reuse existing parsed revenue dimensions, including their comparative columns.
    # Every row stays linked to the immutable cache and its corresponding SEC filing.
    for path in sorted(LOCAL.glob("*.json")):
        if path.name.startswith("_"):
            continue
        entry = json.loads(path.read_text(encoding="utf-8"))
        source = next((s for s in primary if s["accession"] == path.stem), None)
        if not source:
            continue
        table = (((entry.get("dataframes") or {}).get("income_statement") or {}).get("data") or {})
        columns = table.get("columns", [])
        rows = [dict(zip(columns, values)) for values in table.get("data", [])]
        revenue_rows = [r for r in rows if any(term in str(r.get("concept", "")).lower()
                                               for term in ("revenues", "salesrevenue", "revenuefromcontract"))]
        for row in revenue_rows:
            axis = str(row.get("dimension_axis") or "")
            if not axis:
                dimension = "consolidated"
            elif "Geograph" in axis or "Geographic" in axis:
                dimension = "geography"
            elif any(term in axis for term in ("ProductOrService", "Market", "majormarket")):
                dimension = "market_platform"
            else:
                # Segment detail is obtained from the full original table above.
                continue
            name = "Total revenue" if dimension == "consolidated" else row.get("label")
            if name == "Datacenter":
                name = "Data Center"
            group_labels = sorted(r.get("label") for r in revenue_rows if r.get("dimension_axis") == row.get("dimension_axis"))
            definition = dimension + ":" + "/".join(group_labels)
            if dimension == "market_platform" and "Hyperscale" in group_labels:
                definition = "market:2027-Q2-recast" if source["filing_date"] >= "2026-08-26" else "market:2027-Q1"
            for col in columns:
                match = re.match(r"(\d{4}-\d{2}-\d{2}) \((Q[1-4]|YTD|FY|Annual)\)", col)
                value = row.get(col)
                if not match or not isinstance(value, (int, float)):
                    continue
                end, suffix = match.groups()
                duration = "quarter" if suffix.startswith("Q") else "annual" if suffix in {"FY", "Annual"} else "9" if period(end, "quarter")[2] == 3 else "6"
                add(source, dimension, name, value / 1e6, end, duration, definition,
                    f"Local XBRL income_statement: {row.get('concept')} / {axis} / {row.get('dimension_member')} / {col}; cache SHA256={hashlib.sha256(path.read_bytes()).hexdigest()}",
                    parent="Data Center" if name in {"Hyperscale", "AI Clouds, Industrial, & Enterprise"} else None)

    # Read the SEC CFO attachments directly, preserving every explicit comparison.
    period_ends = {r["period"]: r["period_end"] for r in records if r["duration"] == "quarter"}
    for source in primary:
        end = source.get("report_date")
        if end:
            p = period(end, "quarter")[0]
            if p:
                period_ends[p] = end
    for source in sources.values():
        if not source["document_type"].startswith("EX-99"):
            continue
        path = PACK / source["filename"]
        if path.suffix.lower() not in {".htm", ".html", ".txt"}:
            continue
        soup = BeautifulSoup(path.read_bytes(), "html.parser")
        for index, table in enumerate(soup.find_all("table")):
            text = clean(table.get_text(" ", strip=True))
            if not re.search(r"Revenue by (?:Market|Markets)", text, re.I) or len(text) > 6000:
                continue
            headers = re.findall(r"Q([1-4])\s*FY\s*(\d{2,4})", text)
            labels = [f"FY{int(y)+2000 if len(y)==2 else int(y)}Q{q}" for q,y in headers]
            if not labels or len(labels) != len(set(labels)):
                continue
            lines = [[clean(c.get_text(" ", strip=True)) for c in row.find_all(["td", "th"], recursive=False)] for row in table.find_all("tr")]
            data = {}
            known = {"Data Center", "Datacenter", "Gaming", "Professional Visualization", "Automotive", "OEM and Other", "OEM & Other", "OEM & IP", "OEM and IP", "Compute", "Networking", "Hyperscale", "AI Clouds, Industrial, & Enterprise", "Edge Computing", "Total"}
            for line in lines:
                cells = [c for c in line if c]
                if not cells or cells[0] not in known:
                    continue
                values = amounts(cells[1:])
                if len(values) >= len(labels):
                    data["Data Center" if cells[0] == "Datacenter" else cells[0]] = values[:len(labels)]
            if "Total" not in data:
                continue
            modern = "Hyperscale" in data
            definition = "market:2027-Q2-recast" if modern and source["filing_date"] >= "2026-08-26" else "market:2027-Q1" if modern else "market:" + "/".join(sorted(data))
            top = ["Data Center", "Edge Computing"] if modern else [n for n in data if n not in {"Compute", "Networking", "Total"}]
            for col, label in enumerate(labels):
                if label not in period_ends:
                    continue
                delta = sum(data[n][col] for n in top) - data["Total"][col]
                child_names = ["Hyperscale", "AI Clouds, Industrial, & Enterprise"] if modern else ["Compute", "Networking"]
                child_delta = sum(data[n][col] for n in child_names) - data["Data Center"][col] if all(n in data for n in child_names) else None
                if abs(delta) > .01 or (child_delta is not None and abs(child_delta) > .01):
                    checks.append(dict(source_id=source["source_id"], period=label, metric="market_revenue", difference=delta, status="fail", record_ids=[]))
                    continue
                start = len(records)
                for name, values in data.items():
                    add(source, "consolidated" if name == "Total" else "market_platform",
                        "Total revenue" if name == "Total" else name, values[col], period_ends[label], "quarter", definition,
                        f"HTML table {index}, Revenue by Market Platform, {label}, {name}",
                        parent="Data Center" if name in child_names else None)
                    records[-1]["status"] = "verified"
                    records[-1]["presentation"] = "as_reported" if col == 0 else "comparative_as_disclosed"
                checks.append(dict(source_id=source["source_id"], period=label, metric="market_revenue", difference=delta,
                                   children_difference=child_delta, status="pass", record_ids=[r["id"] for r in records[start:]]))

    # Audit every local dimension group against consolidated revenue from that filing.
    local_groups = collections.defaultdict(list)
    for record in records:
        if record["source"]["locator"].startswith("Local XBRL") and record["dimension"] in {"market_platform", "geography"}:
            local_groups[(record["published_in_accession"], record["period"], record["dimension"], record["classification_version"])].append(record)
    for key, group in local_groups.items():
        controls = [r for r in records if r["published_in_accession"] == key[0] and r["period"] == key[1]
                    and r["dimension"] == "consolidated" and r["metric"] == "revenue" and r["source"]["locator"].startswith("Local XBRL")]
        if not controls:
            continue
        names = {r["node_id"] for r in group}
        leaves = [r for r in group if not (r["node_id"] == "Data Center" and {"Hyperscale", "AI Clouds, Industrial, & Enterprise"} <= names)]
        delta = round(sum(r["value"] for r in leaves)-controls[0]["value"], 6)
        # A parsed statement can omit members; that is an incomplete baseline,
        # not evidence of missing disclosure or a license to invent Other.
        status = "pass" if abs(delta) <= .01 else "incomplete"
        checks.append(dict(source_id=group[0]["source"]["source_id"], period=key[1], metric=key[2]+"_revenue",
                           difference=delta, status=status, record_ids=[r["id"] for r in group]))
        if status == "pass":
            for record in group + controls:
                record["status"] = "verified"

    # Q4 = annual minus nine months, with exactly the same classification and source units.
    keyed = collections.defaultdict(list)
    for record in list(records):
        keyed[(record["fiscal_year"], record["dimension"], record["node_id"], record["metric"], record["classification_version"], record["duration"])].append(record)
    for key, annuals in list(keyed.items()):
        if key[-1] != "annual":
            continue
        ytds = keyed.get((*key[:-1], "9"), [])
        if not ytds:
            continue
        annual = min(annuals, key=lambda r: r["filing_date"])
        comparable_ytds = [r for r in ytds if r["filing_date"] <= annual["filing_date"]]
        if not comparable_ytds:
            continue
        ytd = max(comparable_ytds, key=lambda r: r["filing_date"])
        source = sources[annual["source"]["source_id"]]
        add(source, annual["dimension"], annual["node_id"], annual["value"]-ytd["value"], annual["period_end"],
            "quarter", annual["classification_version"], "Annual minus nine-month YTD", annual["metric"],
            derivation=dict(method="annual_minus_ytd", inputs=[annual["id"], ytd["id"]]))
        records[-1]["status"] = "verified" if annual["status"] == ytd["status"] == "verified" else "provisional"

    result = dict(schema_version="1.0", company=dict(ticker="NVDA", cik=1045810, name="NVIDIA Corporation"),
                  extracted_at=dt.datetime.now(dt.timezone.utc).isoformat(), extractor_version=VERSION,
                  classification_version="nvda-layout-1", sources_manifest_sha256=hashlib.sha256((PACK/"manifest.json").read_bytes()).hexdigest(),
                  records=records, classification_changes=[
                      dict(period="FY2027Q1", change="Market framework: Hyperscale/ACIE and Edge Computing replaces legacy presentation"),
                      dict(period="FY2027Q2", change="Hyperscale/ACIE customer reclassification; preserve originals and recast comparatives")])
    (OUT/"breakdown.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    (OUT/"candidate_tables.json").write_text(json.dumps(candidates, ensure_ascii=False, indent=2), encoding="utf-8")
    check_result = dict(reconciliations=checks, source_count=len(sources), primary_filing_count=len(primary),
                        record_count=len(records), status="provisional", unresolved=[
                            "Early plain-text tables and unparsed vertical tables remain in candidate_tables.json.",
                            "Verified marks numerical reconciliation and source integrity, not exhaustive disclosure coverage."])
    (OUT/"checks.json").write_text(json.dumps(check_result, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Records",len(records),"candidates",len(candidates),"checks",len(checks),"failed",sum(c['status']=='fail' for c in checks))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=ROOT)
    parser.add_argument("--cache-repo", type=Path)
    args = parser.parse_args()
    ROOT = args.repo.resolve()
    PACK = ROOT / "output/NVDA_history_sources"
    OUT = ROOT / "output/NVDA_segment_history"
    if args.cache_repo is None:
        args.cache_repo = Path(__file__).resolve().parents[3]
    if args.cache_repo:
        import sys
        sys.path.insert(0, str(args.cache_repo.resolve() / "src"))
        import filing_cache
        LOCAL = filing_cache.ticker_dir("NVDA")
    main()
