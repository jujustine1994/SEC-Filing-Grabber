"""Print consolidated original concepts and calculation parents, read-only."""
import argparse
import audit_fiscal_pipeline as audit

p = argparse.ArgumentParser()
p.add_argument('ticker')
p.add_argument('year')
p.add_argument('--statement', default='income_statement')
p.add_argument('--pattern', default='cost|gross')
args = p.parse_args()
fc, fg = audit.fc, audit.fg
for path in fc.ticker_dir(args.ticker).glob('*.json'):
    if not fc.ACCESSION_RE.fullmatch(path.stem):
        continue
    entry = fc.load_filing(args.ticker, path.stem, fc.cached_cik(args.ticker))
    if entry is None or not entry['filing_date'].startswith(args.year):
        continue
    fin = fc.cached_filing(entry).financials
    stmt = getattr(fin, args.statement)() if fin else None
    if not stmt:
        continue
    df = stmt.to_dataframe()
    col = (fg._current_q_col(df) if args.statement == 'income_statement'
           else next((c for c in df.columns if c not in fg.META_COLS), None))
    rows = df[fg._consolidated_mask(df)]
    rows = rows[(rows['label'].astype(str) + ' ' + rows['concept'].astype(str)).str.contains(args.pattern, case=False)]
    print(path.stem, col)
    print(rows[[c for c in ('concept', 'label', 'standard_concept', 'parent_concept', 'weight', col)
                if c in rows]].to_json(orient='records', indent=1))
