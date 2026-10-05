"""Read-only cached-input audit of the actual GAAP pipeline. No network or DB writes.

--all launches a fresh process per ten companies. --output is a separate evidence
directory. Results represent cached data, not live SEC completeness. Shares from
companyfacts and AI diagnosis are deliberately excluded in both comparison runs.
"""
from __future__ import annotations
import argparse
from dataclasses import asdict
from datetime import date
import gc
import json
from pathlib import Path
import subprocess
import sys
from types import SimpleNamespace
from unittest.mock import patch

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'src'))
import filing_cache as fc
import fetcher_gaap as fg
from fiscal_audit import label_anomalies


def document_period(entry):
    frame=fc.payload_to_df((entry.get('dataframes') or {}).get('cover'))
    if frame is None or 'concept' not in frame:
        return ''
    rows=frame[frame['concept'].astype(str).str.endswith('DocumentPeriodEndDate')]
    dates=set()
    for _,row in rows.iterrows():
        for column in frame.columns:
            if fg._col_to_period_end(str(column)):
                value=str(row[column])
                if len(value)==10 and fg._col_to_period_end(value):
                    dates.add(value)
    return next(iter(dates)) if len(dates)==1 else ''


def audit(ticker, output):
    filings=[]
    evidence=output.parent/(ticker+'-sec-listing.json')
    official={r['accession_number']:r['reportDate'] for r in json.loads(evidence.read_text(encoding='utf-8'))} if evidence.exists() else {}
    for path in fc.ticker_dir(ticker).glob('*.json'):
        if not fc.ACCESSION_RE.fullmatch(path.stem): continue
        entry=json.loads(path.read_text(encoding='utf-8'))
        if entry.get('schema_version')!=fc.SCHEMA_VERSION or entry.get('edgartools_version')!=fc.edgartools_version():
            raise RuntimeError('Incompatible cached input; audit must not silently skip filings')
        obj=fc.cached_filing(entry)
        filings.append(SimpleNamespace(accession_no=path.stem,form=entry['form'],
            filing_date=date.fromisoformat(entry['filing_date'][:10]),
            period_of_report=official.get(path.stem) or document_period(entry),
            report_date=official.get(path.stem) or document_period(entry),obj=lambda obj=obj:obj))
    filings.sort(key=lambda f:f.filing_date, reverse=True)
    company=SimpleNamespace(cik=None,name=ticker)
    fg.reset_cf_fallbacks()
    with patch.object(fg,'Company',return_value=company), \
         patch.object(fg,'_bind_disk_cache'), \
         patch.object(fg,'_list_filings',side_effect=lambda company,form:[f for f in filings if f.form==form]), \
         patch.object(fg,'run_diagnosis',return_value={}), \
         patch.object(fg,'_fetch_shares_outstanding',return_value={}), \
         fg.collect_gaps() as ledger:
        tables=fg.fetch_gaap_statements(ticker,'Cached audit audit@example.com')
    quarterly=next(t for t in tables if t.sheet_name=='Data_Financials(Q)')
    result=dict(ticker=ticker,mode='cached-read-only',max_filings=80,max_annual_filings=20,
        anomalies=label_anomalies(quarterly.quarter_labels,quarterly.period_ends),
        cf_fallbacks=fg.cf_fallbacks(),gaps=[asdict(g) for g in ledger.gaps],
        tables=[asdict(t) for t in tables])
    (output/(ticker+'.json')).write_text(json.dumps(result,ensure_ascii=False),encoding='utf-8')


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('tickers',nargs='*')
    parser.add_argument('--all',action='store_true')
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args();args.output.mkdir(parents=True,exist_ok=True)
    targets=args.tickers or sorted(p.name for p in fc.cache_root().iterdir() if p.is_dir())
    if args.all:
        failed=0
        for start in range(0,len(targets),10):
            result=subprocess.run([sys.executable,__file__,'--output',str(args.output),*targets[start:start+10]])
            failed+=result.returncode!=0
        return int(bool(failed))
    failed=0
    for ticker in targets:
        try:
            audit(ticker,args.output)
            print(ticker+' OK',flush=True)
        except Exception as exc:
            failed+=1
            (args.output/(ticker+'.error.json')).write_text(json.dumps(dict(error=type(exc).__name__,message=str(exc))),encoding='utf-8')
            print(ticker+' FAILED '+type(exc).__name__,flush=True)
        gc.collect()
    return int(bool(failed))


if __name__=='__main__': sys.exit(main())
