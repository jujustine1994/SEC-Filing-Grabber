"""Verify exported Revenue against independently extracted original SEC facts."""
import argparse,json
from pathlib import Path


def main():
    parser=argparse.ArgumentParser();parser.add_argument('candidate',type=Path)
    args=parser.parse_args()
    evidence=Path(__file__).resolve().parents[1]/'docs/superpowers/evidence/revenue-source-examples.json'
    cases=json.loads(evidence.read_text(encoding='utf-8'));failed=0
    for case in cases:
        data=json.loads((args.candidate/(case['ticker']+'.json')).read_text(encoding='utf-8'))
        table=next(t for t in data['tables'] if t['sheet_name']==case['sheet'])
        columns=[i for i,end in enumerate(table['period_ends']) if end==case['end']]
        values=[table['values'][table['concepts'].index('Revenue')][i] for i in columns]
        ok=len(values)==1 and values[0]==case['value'];failed+=not ok
        print(case['ticker'],case['end'],'expected',case['value'],'actual',values,'OK' if ok else 'FAILED')
    print('Source examples',len(cases),'failures',failed)
    return int(bool(failed))


if __name__=='__main__':raise SystemExit(main())
