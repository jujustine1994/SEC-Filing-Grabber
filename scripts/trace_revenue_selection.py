"""Trace actual cached-input Revenue selections without changing source data.

Uses the same read-only production adapter and frozen official metadata as
audit_fiscal_pipeline. Tags ephemeral dataframe copies with their accession so
diagnostics cannot silently substitute a later comparative filing.
"""
import argparse
import json
import re
from pathlib import Path
from unittest.mock import patch

import audit_fiscal_pipeline as audit


def trace(ticker, output):
    fc, fg = audit.fc, audit.fg
    original_cached = fc.cached_filing
    original_match = fg._match_revenue_row
    records = []

    def tagged(entry):
        obj = original_cached(entry)
        if obj.financials is not None:
            statement = obj.financials.income_statement()
            if statement is not None:
                original_dataframe = statement.to_dataframe

                def dataframe():
                    df = original_dataframe()
                    if df is not None:
                        df.attrs['trace_accession'] = entry['accession_no']
                        df.attrs['trace_form'] = entry['form']
                    return df

                statement.to_dataframe = dataframe
        return obj

    def matched(df, column):
        index, ambiguous = original_match(df, column)
        rows = []
        for i, row in df[fg._consolidated_mask(df)].iterrows():
            text = str(row.get('concept', '')) + ' ' + str(row.get('label', ''))
            if not re.search(r'revenue|sales|interest|fee|cost|provision|dispos', text, re.I):
                continue
            item = {name: fg._to_python_val(row.get(name)) for name in
                    ('concept', 'label', 'standard_concept', 'parent_concept',
                     'parent_abstract_concept', 'weight', 'level')}
            item.update(index=int(i), value=fg._to_python_val(row.get(column)))
            rows.append(item)
        records.append(dict(ticker=ticker, accession=df.attrs['trace_accession'],
                            form=df.attrs['trace_form'], column=column,
                            selected_index=None if index is None else int(index),
                            ambiguous=bool(ambiguous), rows=rows))
        return index, ambiguous

    audit_directory = output.with_name(output.name + '-audits')
    audit_directory.mkdir(parents=True, exist_ok=True)
    with patch.object(fc, 'cached_filing', side_effect=tagged), \
         patch.object(fg, '_match_revenue_row', side_effect=matched):
        audit.audit(ticker, audit_directory, require_official_metadata=True)
    path = output / (ticker + '.json')
    path.write_text(json.dumps(records, ensure_ascii=False,
                              default=lambda value: value.item() if hasattr(value, 'item') else str(value)),
                    encoding='utf-8')
    print(ticker, 'selections', len(records), 'ambiguous', sum(r['ambiguous'] for r in records), flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('tickers', nargs='+')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    for ticker in args.tickers:
        trace(ticker, args.output)


if __name__ == '__main__':
    main()
