"""Compare repaired financial Revenue cells with source-certified examples.

Offline and independent of Revenue selection code. Each example uses the actual
accession traced before repair; duplicate output periods are rejected, not picked.
"""
import argparse
import json
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('candidate', type=Path)
    args = parser.parse_args()
    evidence = Path(__file__).resolve().parents[1] / 'tests/fixtures/revenue-resolution-cases.json'
    cases = json.loads(evidence.read_text(encoding='utf-8'))
    failures = []
    for case in cases:
        data = json.loads((args.candidate / (case['ticker'] + '.json')).read_text(encoding='utf-8'))
        sheet = 'Data_Financials(Y)' if case['column'].endswith('(FY)') else 'Data_Financials(Q)'
        table = next(t for t in data['tables'] if t['sheet_name'] == sheet)
        end = case['column'][:10]
        columns = [i for i, date in enumerate(table['period_ends']) if date == end]
        values = [table['values'][table['concepts'].index('Revenue')][i] for i in columns]
        expected = case['expected_value']
        if len(values) != 1 or values[0] != expected:
            failures.append((case['ticker'], case['accession'], end, expected, values))
        print(case['ticker'], end, expected, values, 'OK' if len(values)==1 and values[0]==expected else 'FAILED')
    print(len(cases), 'source-certified Revenue cases;', len(failures), 'failures')
    return int(bool(failures))


if __name__ == '__main__':
    raise SystemExit(main())
