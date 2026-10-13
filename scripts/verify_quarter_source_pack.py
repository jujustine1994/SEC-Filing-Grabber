"""Check production raw-table parser against independently selected official cells."""
import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from quarterly_source import read_quarter_source_pack

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('pack', type=Path)
parser.add_argument('evidence', type=Path)
parser.add_argument('--cik', type=int, required=True)
args = parser.parse_args()
records = read_quarter_source_pack(args.pack, args.cik)
evidence = json.loads(args.evidence.read_text(encoding='utf-8'))
keys = {'basic_shares': 'Basic Shares', 'diluted_shares': 'Diluted Shares',
        'basic_eps': 'Basic EPS', 'diluted_eps': 'Diluted EPS'}
failures = []
for item in evidence['selected']:
    record = records.get(item['period_end'])
    if record is None:
        failures.append((item['period'], 'not parsed'))
        continue
    for name, key in keys.items():
        expected = item['values'][name] * (1 if key.endswith('EPS') else 1_000_000)
        if abs(expected - record['values'][key]) > 1e-5:
            failures.append((item['period'], key, expected, record['values'][key]))
print(json.dumps(dict(raw_periods=len(records), checked_periods=len(evidence['selected']), failures=failures)))
raise SystemExit(bool(failures))
