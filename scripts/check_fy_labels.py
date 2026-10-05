# -*- coding: utf-8 -*-
"""Cached, read-only fiscal-label check using the actual statement pipeline.

Without --verify, scan month variability as a descriptive feature only. A stable
month does not exclude week-based quarter problems (KR has a sixteen-week Q1).
--verify checks every requested company, including unclassified labels and
incomplete years. Incomplete years are inconclusive, never proof of correctness.
For all 215 companies use audit_fiscal_pipeline.py --all to bound process memory.
Python lives at ~/venvs/SEC Financial Tools/Scripts/python.exe, outside this repo.
"""
from __future__ import annotations

import collections
import json
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
import filing_cache

CACHE = filing_cache.cache_root()
FY_COL = re.compile(r'^(\d{4})-(\d{2})-(\d{2})\s+\(FY\)')


def scan(ticker):
    months = collections.Counter()
    for path in (CACHE / ticker).glob('*.json'):
        if not filing_cache.ACCESSION_RE.fullmatch(path.stem):
            continue
        entry = json.loads(path.read_text(encoding='utf-8'))
        payload = (entry.get('dataframes') or {}).get('income_statement')
        for column in payload['data']['columns'] if payload else []:
            match = FY_COL.match(str(column))
            if match:
                months[int(match[2])] += 1
    return dict(ticker=ticker, months_seen=dict(months), is_week_based=len(months)>1,
                skipped='No annual columns' if not months else '')


def verify(ticker, identity=None):
    # identity is retained for old callers; no SEC/AI calls or DB writes here.
    from audit_fiscal_pipeline import audit
    output = ROOT / 'output' / 'fy-label-check'
    output.mkdir(parents=True, exist_ok=True)
    return audit(ticker, output)['anomalies']


def main(argv):
    do_verify = '--verify' in argv
    targets = [a for a in argv if a != '--verify'] or sorted(d.name for d in CACHE.iterdir() if d.is_dir())
    failures = 0
    for ticker in targets:
        try:
            feature = scan(ticker)
            if not do_verify:
                print(ticker, feature, flush=True)
                continue
            issues = verify(ticker)
            bad = any(issues[k] for k in ('unclassified', 'duplicate_ends', 'collisions', 'misordered_years'))
            incomplete = issues['incomplete_years']
            failures += bool(bad)
            status = 'ANOMALIES' if bad else 'INCONCLUSIVE' if incomplete else 'CONSISTENT'
            print(ticker, status, issues, flush=True)
        except Exception as exc:
            failures += 1
            print(ticker, 'ERROR', type(exc).__name__, flush=True)
    if not do_verify:
        print('Month variability is not a verdict; add --verify. For a full run use audit_fiscal_pipeline.py --all.')
    return int(bool(failures))


if __name__ == '__main__':
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except (AttributeError, ValueError, OSError):
            pass
    sys.exit(main(sys.argv[1:]))
