"""Compare GUI preview identities with saved pipeline outputs, read-only/offline.

Scope is the saved filing set plus frozen official listing metadata, not unsaved
new SEC filings. Segments are excluded from this period-identity verification.
"""
import argparse
from datetime import date
import json
from pathlib import Path
import sys
from types import SimpleNamespace
from unittest.mock import patch

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'src'))
import filing_cache as fc
import fetcher_gaap as fg
import local_db


def verify(ticker,audits):
    cik=local_db.read_meta(ticker)['cik']
    official={r['accession_number']:r['reportDate'] for r in json.loads((audits.parent/(ticker+'-sec-listing.json')).read_text(encoding='utf8'))}
    filings=[]
    for path in fc.ticker_dir(ticker).glob('*.json'):
        if not fc.ACCESSION_RE.fullmatch(path.stem):continue
        entry=fc.load_filing(ticker,path.stem,cik)
        assert entry is not None,'Production cache gates rejected input'
        if date.fromisoformat(entry['filing_date'][:10])>=fg._XBRL_CUTOFF:
            assert path.stem in official,'Official metadata missing'
        end=official.get(path.stem,'')
        filings.append(SimpleNamespace(accession_no=path.stem,form=entry['form'],
                       filing_date=date.fromisoformat(entry['filing_date'][:10]),
                       period_of_report=end,report_date=end))
    filings.sort(key=lambda f:f.filing_date,reverse=True)
    company=SimpleNamespace(cik=cik,fiscal_year_end='')
    with patch.object(fg,'Company',return_value=company),patch.object(fg,'set_identity'), \
         patch.object(fg,'_list_filings',side_effect=lambda c,form:[f for f in filings if f.form==form]), \
         patch.object(fg,'_build_segment_tables',return_value=[]), \
         patch.object(fg,'_r_file_count',return_value=None), \
         patch.object(fc,'save_filing',side_effect=AssertionError('No source writes')), \
         patch.object(fc,'save_sixk_probe',side_effect=AssertionError('No probe writes')):
        preview=fg.preview_sheets(ticker,'Offline preview verification')
    data=json.loads((audits/(ticker+'.json')).read_text(encoding='utf8'))
    table=next(t for t in data['tables'] if t['sheet_name']=='Data_Financials(Q)')
    labels=sorted({label for label,end in zip(table['quarter_labels'],table['period_ends']) if end==preview['latest_period_end']})
    status=('no-output-current-period' if not labels else 'estimated' if preview['label_estimated']
            else 'pass' if labels==[preview['latest_label']] else 'FAIL')
    return dict(ticker=ticker,status=status,preview=preview,pipeline_labels=labels)


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--audits',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args()
    results=[]
    for path in sorted(args.audits.glob('*.json')):
        assert not path.name.endswith('.error.json')
        results.append(verify(path.stem,args.audits))
        if len(results)%25==0:print('Preview periods checked',len(results),flush=True)
    from collections import Counter
    report=dict(companies=len(results),status_counts=dict(Counter(r['status'] for r in results)),results=results)
    args.output.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    print(report['status_counts'],flush=True)
    assert not any(r['status']=='FAIL' for r in results),'Confirmed preview differs from pipeline'


if __name__=='__main__':main()
