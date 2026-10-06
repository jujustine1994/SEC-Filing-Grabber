"""Check the fixed financial row contract in saved read-only pipeline outputs."""
import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from fetcher_gaap import IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE, SECTION_GAP


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('candidate', type=Path)
    args = parser.parse_args()
    fixed = ['Fiscal Quarter', 'Calendar Quarter', 'Period End', '', 'Income Statement']
    fixed += [row[0] for row in IS_TEMPLATE] + [''] * SECTION_GAP + ['Balance Sheet']
    fixed += [row[0] for row in BS_TEMPLATE] + [''] * SECTION_GAP + ['Cash Flow']
    fixed += [row[0] for row in CF_TEMPLATE]
    ng_fixed = ['Fiscal Quarter', 'Calendar Quarter', 'Period End', '', 'Income Statement']
    ng_fixed += [''] * SECTION_GAP + ['Balance Sheet'] + [''] * SECTION_GAP + ['Cash Flow']
    failures = []
    checked = 0
    files = list(args.candidate.glob('*.json'))
    assert files and not list(args.candidate.glob('*.error.json')), 'Missing or failed input'
    for path in files:
        data = json.loads(path.read_text(encoding='utf-8'))
        for table in data['tables']:
            name = table['sheet_name']
            if name in ('Data_Financials(Q)', 'Data_Financials(Y)'):
                expected = fixed
            elif name in ('Data_Financials_NG(Q)', 'Data_Financials_NG(Y)'):
                expected = ng_fixed
            else:
                continue
            checked += 1
            if table['concepts'][:len(expected)] != expected:
                failures.append((data['ticker'], name))
    print(len(files), 'companies;', checked, 'financial tables;', len(failures), 'fixed-layout failures')
    if failures:
        print(failures)
    return int(bool(failures))


if __name__ == '__main__':
    raise SystemExit(main())
