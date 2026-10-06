"""Exhaustively compare scoped Revenue rule changes on saved income frames.

Only supports changes inside Other Income groups and percentage classification;
the unchanged outer selector is checked by AST. Compiles pure production helpers
without edgartools/config initialization. Reads current filing JSON, never writes
the database. Includes ALL date columns, not just the runtime-selected period.
"""
import argparse
import ast
import hashlib
import json
from pathlib import Path
import re
import time
from typing import Any
import unicodedata

import pandas as pd

FUNCTIONS = {'_consolidated_mask', '_to_python_val', '_revenue_label',
             '_revenue_concept_words', '_revenue_competitors', '_operating_revenue_group',
             '_uncovered_revenue_component', '_revenue_total_rows', '_revenue_components',
             '_match_revenue_row'}


def load_rules(path):
    source = path.read_text(encoding='utf-8-sig')
    tree = ast.parse(source)
    nodes = []
    functions = {}
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in FUNCTIONS:
            nodes.append(node)
            functions[node.name] = node
        elif isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            if any(isinstance(t, ast.Name) and t.id in ('META_COLS', '_REVENUE_CONCEPT_TIERS') for t in targets):
                nodes.append(node)
    assert set(functions) == FUNCTIONS, 'Unsupported selector version'
    namespace = {'pd': pd, 're': re, 'unicodedata': unicodedata, 'Any': Any}
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(path), 'exec'), namespace)
    return namespace, functions, hashlib.sha256(source.encode('utf-8')).hexdigest()


def frame(payload):
    raw = payload['data']
    df = pd.DataFrame(raw['data'], index=raw['index'], columns=raw['columns'])
    for column, dtype in payload.get('dtypes', {}).items():
        if column in df and (dtype in ('object', 'bool') or dtype.startswith(('int', 'uint', 'float'))):
            try:
                df[column] = df[column].astype(dtype)
            except (TypeError, ValueError):
                pass
    return df


def selection(rules, df, column):
    index, ambiguous = rules['_match_revenue_row'](df, column)
    components, conflict = rules['_revenue_components'](df, column) if index is None and not ambiguous else ([], False)
    value = (rules['_to_python_val'](df.loc[index, column]) if index is not None
             else sum(rules['_to_python_val'](df.loc[i, column]) for i in components) if components else None)
    return {'index': None if index is None else int(index), 'ambiguous': bool(ambiguous or conflict),
            'components': [int(i) for i in components], 'value': value}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--before', type=Path, required=True)
    parser.add_argument('--after', type=Path, required=True)
    parser.add_argument('--database', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    before, old_functions, old_sha = load_rules(args.before)
    after, new_functions, new_sha = load_rules(args.after)
    # Scope of the optimization: the outer selector, subtotal coverage, label
    # normalization and dimension/value filters must be exactly unchanged.
    for name in ('_match_revenue_row', '_uncovered_revenue_component', '_revenue_label',
                 '_revenue_concept_words', '_consolidated_mask', '_to_python_val'):
        assert ast.dump(old_functions[name]) == ast.dump(new_functions[name]), name
    # Only the casefold of PercentToSales is allowed to change competitors.
    old_comp = ast.dump(old_functions['_revenue_competitors'])
    stripped = ast.parse(ast.unparse(new_functions['_revenue_competitors']).replace(
        ".str.casefold().str.endswith('percenttosales')", ".str.endswith('PercentToSales')")).body[0]
    assert old_comp == ast.dump(stripped), 'Unbounded competitor change'
    started = time.monotonic()
    scanned = frames = relevant = periods = filings = 0
    tickers = set()
    changes = []
    for path in sorted((args.database / 'filings').glob('*/*.json')):
        data = json.loads(path.read_text(encoding='utf-8'))
        if isinstance(data.get('dataframes'), dict):
            filings += 1
        payload = (data.get('dataframes') or {}).get('income_statement')
        if payload is None:
            continue
        scanned += 1
        tickers.add(path.parent.name)
        raw = payload['data']
        columns = raw['columns']
        labels = [row[columns.index('label')] for row in raw['data']]
        concepts = [str(row[columns.index('concept')]) for row in raw['data']]
        other_income = any(re.fullmatch(r'(?:total )?revenues? and other income', before['_revenue_label'](s)) for s in labels)
        percentage_change = any(s.casefold().endswith('percenttosales') and not s.endswith('PercentToSales') for s in concepts)
        frames += 1
        if not other_income and not percentage_change:
            continue
        relevant += 1
        df = frame(payload)
        for column in df.columns:
            if not re.match(r'^\d{4}-\d{2}-\d{2}', str(column)):
                continue
            periods += 1
            a = selection(before, df, column)
            b = selection(after, df, column)
            if a != b:
                changes.append({'ticker': path.parent.name, 'file': path.name, 'column': str(column), 'before': a, 'after': b})
        if relevant % 100 == 0:
            print('Relevant frames', relevant, 'scanned', scanned, 'changes', len(changes), flush=True)
    result = {'before_sha256_normalized_text': old_sha, 'after_sha256_normalized_text': new_sha,
              'filings': filings, 'income_filings': scanned, 'companies': len(tickers), 'frames': frames,
              'relevant_frames': relevant, 'periods_compared': periods,
              'affected_companies': sorted({c['ticker'] for c in changes}),
              'elapsed_seconds': time.monotonic()-started, 'changes': changes}
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k!='changes'}), flush=True)


if __name__ == '__main__':
    main()
