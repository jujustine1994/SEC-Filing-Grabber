"""Verify raw overflow value bags survive moves between GAAP and NG sheets.

Use original source concept, quarter/year sheet kind and exact end date; retain
duplicate-period values and multiplicities. Never compare display row numbers.
"""
import argparse
from collections import Counter,defaultdict
import json
from pathlib import Path


def values(result):
    out=defaultdict(Counter)
    for table in result['tables']:
        if not table['sheet_name'].startswith('Data_Financials'):continue
        concepts=table['concepts'];sources=table['labels']
        if 'Other (as reported)' not in concepts:continue
        kind='quarter' if table['sheet_name'].endswith('(Q)') else 'annual'
        for row in range(concepts.index('Other (as reported)')+1,len(concepts)):
            source=sources[row]
            assert source,'Overflow must retain its raw source identity'
            for col,value in enumerate(table['values'][row]):
                if value is None:continue
                end=table['period_ends'][col]
                period=('end',end) if end else ('label',table['quarter_labels'][col])
                out[(kind,source,*period)][json.dumps(value,sort_keys=True)]+=1
    return out


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('before',type=Path);parser.add_argument('after',type=Path)
    parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--require-complete',action='store_true')
    args=parser.parse_args();results={}
    before={p.name:p for p in args.before.glob('*.json')}
    after={p.name:p for p in args.after.glob('*.json')}
    if args.require_complete:assert len(before)==len(after)==215 and before.keys()==after.keys()
    for name,path in sorted(after.items()):
        assert name in before and not name.endswith('.error.json')
        a=values(json.loads(before[name].read_text(encoding='utf8')))
        b=values(json.loads(path.read_text(encoding='utf8')))
        keys=a.keys()|b.keys()
        changes=[dict(identity=key,before=dict(a.get(key,{})),after=dict(b.get(key,{}))) for key in sorted(keys) if a.get(key)!=b.get(key)]
        results[name[:-5]]=dict(keys_checked=len(keys),changes=changes)
    report=dict(companies=len(results),complete=len(results)==len(before)==215,
                changed_companies=[t for t,r in results.items() if r['changes']],results=results)
    args.output.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps({k:v for k,v in report.items() if k!='results'}),flush=True)
    assert not report['changed_companies'],'Overflow values changed during reclassification'


if __name__=='__main__':main()
