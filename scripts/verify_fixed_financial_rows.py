"""Check the fixed financial row contract in saved read-only pipeline outputs."""
import argparse
import ast
import json
from pathlib import Path


def source_contract():
    """Read literal template names without initializing edgartools or its cache."""
    path = Path(__file__).resolve().parents[1] / 'src/fetcher_gaap.py'
    tree = ast.parse(path.read_text(encoding='utf-8'))
    result = {}
    for node in tree.body:
        if isinstance(node, ast.Assign):
            targets, value = node.targets, node.value
        elif isinstance(node, ast.AnnAssign):
            targets, value = [node.target], node.value
        else:
            continue
        for target in targets:
            if not isinstance(target, ast.Name):
                continue
            if target.id in ('IS_TEMPLATE', 'BS_TEMPLATE', 'CF_TEMPLATE'):
                result[target.id] = [ast.literal_eval(row.elts[0]) for row in value.elts]
            elif target.id == 'SECTION_GAP':
                result[target.id] = ast.literal_eval(value)
    assert len(result) == 4, 'Template contract is not statically readable'
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('candidate', type=Path)
    args = parser.parse_args()
    contract = source_contract()
    gap = [''] * contract['SECTION_GAP']
    fixed = ['Fiscal Quarter', 'Calendar Quarter', 'Period End', '', 'Income Statement']
    fixed += contract['IS_TEMPLATE'] + gap + ['Balance Sheet']
    fixed += contract['BS_TEMPLATE'] + gap + ['Cash Flow']
    fixed += contract['CF_TEMPLATE']
    ng_fixed = ['Fiscal Quarter', 'Calendar Quarter', 'Period End', '', 'Income Statement']
    ng_fixed += gap + ['Balance Sheet'] + gap + ['Cash Flow']
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
