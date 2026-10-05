"""Fetch official SEC listing metadata separately from read-only financial audits.

Create a fresh evidence directory; never overwrite a prior run or write to the
permanent database. Subsequent audits consume these frozen listing snapshots.
"""
import argparse
from pathlib import Path
import sys

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'src'))
from config import load_config
from database import database_root
from edgar import Company,set_identity
import filing_cache as fc
import local_db


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('tickers',nargs='*')
    parser.add_argument('--all',action='store_true')
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args()
    destination=args.output.resolve()
    if destination.is_relative_to(database_root().resolve()):
        parser.error('Audit evidence must be outside the permanent database')
    if args.all and args.tickers:
        parser.error('Use either --all or explicit tickers')
    targets=sorted(p.name for p in fc.cache_root().iterdir() if p.is_dir()) if args.all else args.tickers
    if not targets:
        parser.error('Specify --all or tickers')
    targets=list(dict.fromkeys(t.strip().upper() for t in targets))
    if any(not t.isalnum() for t in targets):
        parser.error('Invalid ticker')
    paths=[destination/(ticker+'-sec-listing.json') for ticker in targets]
    if any(p.exists() for p in paths):
        parser.error('Existing evidence is immutable; choose a fresh output directory')
    identity=load_config().get('identity')
    if not identity:
        parser.error('Configure an SEC identity before fetching metadata')
    set_identity(identity)
    destination.mkdir(parents=True,exist_ok=True)
    for ticker,path in zip(targets,paths):
        meta=local_db.read_meta(ticker) or {}
        cik=meta.get('cik')
        if not isinstance(cik,int) or isinstance(cik,bool) or cik<=0:
            raise RuntimeError(ticker+': missing valid CIK')
        filings=Company(cik).get_filings(form=['10-K','10-Q','20-F','6-K'],amendments=False)
        frame=filings.data.to_pandas()
        # Exclusive creation also protects against another concurrent run.
        with path.open('x',encoding='utf-8') as handle:
            handle.write(frame.to_json(orient='records',date_format='iso'))
        print(ticker,len(frame),'official filing records',flush=True)
    return 0


if __name__=='__main__':raise SystemExit(main())
