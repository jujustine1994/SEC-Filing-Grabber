"""Extract and reconcile the isolated NVIDIA pilot; no model API calls."""
import json
import re
from pathlib import Path
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parents[1] / "output/nvda_revenue_pilot"
NUMBER = r"\$?([\d,]+)"
MARKETS = ["Data Center", "Compute", "Networking", "Gaming",
           "Professional Visualization", "Automotive", "OEM and Other",
           "Hyperscale", "AI Clouds, Industrial, & Enterprise", "Edge Computing"]


def numbers_after(line, label):
    rest = line[len(label):].strip()
    return [int(s.replace(",", "")) for s in re.findall(NUMBER, rest)[:3]]


def main():
    manifest = json.loads((ROOT / "manifest.json").read_text())
    filings = json.loads((ROOT / "filings_manifest.json").read_text())
    results = []
    all_chars = selected_chars = 0
    for source, filing in zip(manifest, filings):
        text = (ROOT / (source["period"] + ".txt")).read_text(encoding="utf-8")
        # Quarterly tables precede discussion and, in Q4, full-year tables.
        block = text.split("Revenue by Market Platform", 1)[1].split("Total", 1)[0]
        selected = text.split("Revenue by Market Platform", 1)[0] + block
        all_chars += len(text)
        selected_chars += len(selected)
        market = {}
        for label in MARKETS:
            # Wrapped ACIE label puts Enterprise after the numeric row.
            prefix = "AI Clouds, Industrial, &" if label.startswith("AI Clouds") else label
            match = re.search(r"^" + re.escape(prefix) + r"\s+(.+)$", block, re.M)
            if match:
                market[label] = numbers_after(match.group(0), prefix)
        revenue = int(re.search(r"^Revenue \$([\d,]+)", text, re.M)[1].replace(",", ""))
        ops = [int(s.replace(",", "")) for s in re.findall(r"^Operating income \$([\d,]+)", text, re.M)]
        report_block = text.split("Revenue by Reportable Segments", 1)[1].split("Revenue by Market Platform", 1)[0]
        reportable = {}
        for label in ["Compute & Networking", "Graphics"]:
            match = re.search(r"^" + re.escape(label) + r"\s+(.+)$", report_block, re.M)
            reportable[label] = numbers_after(match.group(0), label)[0]
        soup = BeautifulSoup((ROOT / filing["filename"]).read_bytes(), "html.parser")
        tables = [t for t in soup.find_all("table") if all(k in t.get_text(' ',strip=True)
                  for k in ["Compute", "Graphics", "Operating income"])
                  and len(t.get_text(' ',strip=True)) < 12000]
        assert len(tables) == 1, (source["period"], len(tables))
        table_text = tables[0].get_text(' ',strip=True)
        seg_ops = re.findall(r"Operating income(?: \(loss\))?\s+\$\s*([\d,]+)\s+\$\s*([\d,]+)", table_text)
        assert seg_ops, source["period"]
        first = [int(v.replace(',', '')) for v in seg_ops[0]]
        op_basis = "quarterly reported"
        if source["period"].endswith("Q4"):
            previous = results[-1]
            # 10-K annual segment OI minus current-year nine-month OI from Q3.
            prior_ops = previous["nine_month_segment_operating_income"]
            first = [a-b for a,b in zip(first, prior_ops)]
            op_basis = "annual minus nine months (derived Q4)"
        nine_month = None
        if source["period"].endswith("Q3"):
            assert len(seg_ops) >= 3
            nine_month = [int(v.replace(',', '')) for v in seg_ops[2]]
        top = ["Data Center", "Edge Computing"] if "Edge Computing" in market else [
            "Data Center", "Gaming", "Professional Visualization", "Automotive", "OEM and Other"]
        children = ["Hyperscale", "AI Clouds, Industrial, & Enterprise"] if "Hyperscale" in market else ["Compute", "Networking"]
        checks = dict(market_total=sum(market[k][0] for k in top)-revenue,
                      data_center_children=sum(market[k][0] for k in children)-market['Data Center'][0],
                      reportable_total=sum(reportable.values())-revenue)
        assert all(v == 0 for v in checks.values()), (source["period"], checks)
        results.append(dict(period=source['period'], period_end=filing['period_end'],
            revenue=revenue, gaap_operating_income=ops[0], nongaap_operating_income=ops[1],
            market=market, reportable_revenue=reportable,
            segment_operating_income=dict(zip(reportable, first)), segment_op_basis=op_basis,
            nine_month_segment_operating_income=nine_month,
            source=source, filing=filing, checks=checks))
    dataset = dict(unit='USD millions', basis='As originally reported in each quarter; source-comparable prior quarter used for QoQ',
                   quarters=results, full_text_characters=all_chars, candidate_table_characters=selected_chars,
                   external_model_api_calls=0)
    (ROOT/'dataset.json').write_text(json.dumps(dataset,ensure_ascii=False,indent=2),encoding='utf-8')
    print('Verified',len(results),'quarters;',len(results)*3,'revenue reconciliations passed')
    print('Characters',all_chars,'->',selected_chars)
    for r in results:
        print(r['period'],r['revenue'],r['segment_operating_income'],round(r['gaap_operating_income']/r['revenue']*100,2))


if __name__ == '__main__':
    main()
