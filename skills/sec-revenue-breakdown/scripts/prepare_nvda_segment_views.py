"""Prepare latest-first and inactive historical views without rewriting evidence."""
import collections
import argparse
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output/NVDA_segment_history"


def select(records, latest=False):
    if not records:
        return None
    # Original presentation prefers current-period disclosures, then the earliest
    # available comparison. Latest recast prefers the newest source explicitly
    # disclosing that period under the selected definition.
    key = (lambda r: (r["filing_date"], "HTML table" in r["source"]["locator"])) if latest else (
        lambda r: (r["presentation"] != "as_reported", r["filing_date"], r["derivation"]["method"] != "direct"))
    return max(records, key=key) if latest else min(records, key=key)


def main():
    data = json.loads((OUT / "breakdown.json").read_text(encoding="utf-8"))
    records = [r for r in data["records"] if r["status"] == "verified"]
    quarter_records = [r for r in records if r["duration"] == "quarter"]
    observed_periods = sorted({r["period"] for r in quarter_records})
    periods = [f"FY{y}Q{q}" for y in range(1999, 2028) for q in range(1, 5) if f"FY{y}Q{q}" <= "FY2027Q2"]
    annuals = sorted({r["period"] for r in records if r["duration"] in {"annual", "transition"}})
    registry = []

    def view(dimension, nodes, predicate, status, title, latest=False, parents=()):
        matched = [r for r in records if r["dimension"] == dimension and predicate(r)]
        if not matched:
            return None
        version = f"view{len(registry)+1}"
        item = dict(version_id=version, dimension=dimension, title=title, status=status,
                    first_observed=min(r["period"] for r in matched), last_observed=max(r["period"] for r in matched),
                    note="Observed range is not the classification's formal effective date", nodes=nodes)
        registry.append(item)
        rows = []
        for node in nodes:
            row = dict(node=node, parent="Data Center" if node in parents else None, values={}, records={}, annual_values={}, annual_records={}, op_values={}, op_records={}, annual_op_values={}, annual_op_records={}, source_definitions=sorted({r["classification_version"] for r in matched if r["node_id"] == node}))
            for duration, labels, target in (("quarter", periods, "values"), ("annual", annuals, "annual_values")):
                for p in labels:
                    found = select([r for r in matched if r["node_id"] == node and r["metric"] == "revenue" and r["period"] == p
                                    and (r["duration"] == duration or (duration == "annual" and r["duration"] == "transition"))], latest)
                    if found:
                        row[target][p] = found["value"]
                        row["records" if duration == "quarter" else "annual_records"][p] = found["id"]
                        ops = [r for r in matched if r["node_id"] == node and r["metric"] == "operating_income" and r["period"] == p
                               and r["duration"] == found["duration"] and r["published_in_accession"] == found["published_in_accession"]]
                        op = select(ops, latest)
                        if op:
                            row["op_values" if duration == "quarter" else "annual_op_values"][p] = op["value"]
                            row["op_records" if duration == "quarter" else "annual_op_records"][p] = op["id"]
            rows.append(row)
        return dict(**item, rows=rows)

    blocks = []
    blocks.append(view("market_platform", ["Data Center", "Hyperscale", "AI Clouds, Industrial, & Enterprise", "Edge Computing"],
                       lambda r: r["node_id"] in {"Data Center", "Edge Computing"} or r["classification_version"] == "market:2027-Q2-recast",
                       "active", "最新市場分類（Hyperscale / ACIE 依 FY2027 Q2 重編）", True,
                       ["Hyperscale", "AI Clouds, Industrial, & Enterprise"]))
    blocks.append(view("market_platform", ["Hyperscale", "AI Clouds, Industrial, & Enterprise"],
                       lambda r: r["classification_version"] == "market:2027-Q1", "inactive",
                       "Inactive Segments — FY2027 Q1 原報分類", False,
                       ["Hyperscale", "AI Clouds, Industrial, & Enterprise"]))
    legacy_nodes = ["Compute", "Networking", "Gaming", "Professional Visualization", "Automotive", "OEM and Other", "OEM & Other", "OEM and IP", "OEM & IP"]
    legacy_nodes = [n for n in legacy_nodes if any(r["dimension"] == "market_platform" and r["node_id"] == n for r in records)]
    blocks.append(view("market_platform", legacy_nodes, lambda r: r["node_id"] in legacy_nodes,
                       "inactive", "Inactive Segments — 舊市場分類（OEM 名稱按原文分列）", False,
                       ["Compute", "Networking"]))
    blocks.append(view("revenue_type", ["Product", "Royalty"], lambda r: True, "inactive",
                       "Inactive Segments — 早期 Product / Royalty 營收類型（不是市場分類）"))
    blocks.append(view("reportable_segment", ["Compute & Networking", "Graphics"],
                       lambda r: r["node_id"] in {"Compute & Networking", "Graphics"}, "active", "最新報告部門分類", False))
    old_definitions = {r["classification_version"] for r in records if r["dimension"] == "reportable_segment"
                       and r["node_id"] not in {"Compute & Networking", "Graphics"}}
    old_definitions = sorted(old_definitions, key=lambda d: max(r["period"] for r in records if r["classification_version"] == d), reverse=True)
    for definition in old_definitions:
        nodes = list(dict.fromkeys(r["node_id"] for r in records if r["dimension"] == "reportable_segment" and r["classification_version"] == definition))
        nodes = [n for n in nodes if n not in {"Compute & Networking", "Graphics"}]
        blocks.append(view("reportable_segment", nodes, lambda r,d=definition: r["classification_version"] == d,
                           "inactive", "Inactive Segments — " + " / ".join(nodes), False))
    geo_definitions = {r["classification_version"] for r in records if r["dimension"] == "geography"}
    latest_geo = max(geo_definitions, key=lambda d: max((r["period"],r["filing_date"]) for r in records if r["classification_version"] == d))
    for definition in sorted(geo_definitions, key=lambda d: max(r["period"] for r in records if r["classification_version"] == d), reverse=True):
        nodes = list(dict.fromkeys(r["node_id"] for r in records if r["dimension"] == "geography" and r["classification_version"] == definition))
        active = definition == latest_geo
        blocks.append(view("geography", nodes, lambda r,d=definition: r["classification_version"] == d,
                           "active" if active else "inactive", "最新地域分類" if active else "Inactive Segments — 舊地域分類", active))
    # Move all active blocks above every inactive block; dimensions remain separate.
    blocks = [b for b in blocks if b]
    blocks.sort(key=lambda b: b["status"] != "active")
    controls = {}
    for p in periods + annuals:
        found = select([r for r in records if r["dimension"] == "consolidated" and r["metric"] == "revenue" and r["period"] == p], True)
        if found:
            controls[p] = found["value"]
    coverage = []
    # All periods expected from the first SEC 10-Q through the latest requested quarter.
    for year in range(1999, 2028):
        for q in range(1, 5):
            p = f"FY{year}Q{q}"
            if p > "FY2027Q2":
                continue
            coverage.append(dict(period=p, consolidated=p in controls,
                                 market=any(r["dimension"] == "market_platform" and r["period"] == p for r in quarter_records),
                                 revenue_type=any(r["dimension"] == "revenue_type" and r["period"] == p for r in quarter_records),
                                 reportable=any(r["dimension"] == "reportable_segment" and r["period"] == p for r in quarter_records),
                                 geography=any(r["dimension"] == "geography" and r["period"] == p for r in quarter_records)))
    data["classification_registry"] = registry
    (OUT / "breakdown.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    result = dict(periods=periods, recent_periods=periods[-8:], annual_periods=annuals, blocks=blocks, consolidated=controls,
                  coverage=coverage, source_record_count=len(data["records"]), verified_record_count=len(records))
    (OUT / "views.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Observed quarterly periods", len(observed_periods), "Expected",len(periods),"Annual/transition",len(annuals),"blocks",len(blocks),"verified records",len(records))
    print("Coverage counts",{dim:sum(r[dim] for r in coverage) for dim in ["market", "reportable", "geography"]})


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=ROOT)
    args = parser.parse_args()
    OUT = args.repo.resolve() / "output/NVDA_segment_history"
    main()
