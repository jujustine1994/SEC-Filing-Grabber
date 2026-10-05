"""Compare cached pipeline outputs by actual period and row identity, not FY labels.

Records every changed nonempty cell, including added/removed periods. No financial
materiality threshold and no rounding. Meta/header changes are reported separately.
"""
from collections import Counter
import argparse
import json
from pathlib import Path


def cells(result):
    values, labels = {}, {}
    source_ends={}
    for table in result['tables']:
        if table['sheet_name'] not in ('Data_Financials(Q)','Data_Financials(Y)'): continue
        for label,end in zip(table['quarter_labels'],table.get('period_ends',[])):
            if end: source_ends[label]=end
    for table in result['tables']:
        sheet=table['sheet_name']
        occurrences=Counter()
        for row,concept in enumerate(table['concepts']):
            occurrence=occurrences[concept];occurrences[concept]+=1
            for col,label in enumerate(table['quarter_labels']):
                ends=table.get('period_ends',[])
                dates=table.get('filing_dates',[])
                end=ends[col] if col<len(ends) and ends[col] else ''
                end=end or source_ends.get(label,'')
                period=('end',end) if end else ('filed',dates[col]) if col<len(dates) and dates[col] else ('label',label)
                key=(sheet,concept,occurrence,*period)
                if key in values:
                    raise ValueError('Ambiguous duplicate period key; comparison cannot silently collapse cells')
                values[key]=table['values'][row][col]
                labels[(sheet,*period)]=label
    return values, labels


def compare(before, after):
    old,old_labels=cells(before);new,new_labels=cells(after)
    changes=[];headers=[];metadata=[]
    for key in sorted(old.keys()|new.keys()):
        a=old.get(key);b=new.get(key)
        if a==b:continue
        detail=dict(sheet=key[0],concept=key[1],occurrence=key[2],period_kind=key[3],period=key[4],before=a,after=b)
        target=metadata if key[0]=='Data_Meta' else headers if key[1] in ('Fiscal Quarter','Calendar Quarter','Period End') else changes
        target.append(detail)
    label_changes=[dict(sheet=k[0],period_kind=k[1],period=k[2],before=old_labels.get(k),after=new_labels.get(k))
                   for k in sorted(old_labels.keys()|new_labels.keys()) if old_labels.get(k)!=new_labels.get(k)]
    return dict(value_changes=changes,label_changes=label_changes,header_changes=headers,metadata_changes=metadata)


def main():
    parser=argparse.ArgumentParser();parser.add_argument('before',type=Path);parser.add_argument('after',type=Path);parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args();args.output.mkdir(parents=True,exist_ok=True)
    old={p.stem for p in args.before.glob('*.json') if not p.name.endswith('.error.json')}
    new={p.stem for p in args.after.glob('*.json') if not p.name.endswith('.error.json')}
    summary=dict(before_count=len(old),after_count=len(new),missing_before=sorted(new-old),missing_after=sorted(old-new),companies={})
    for ticker in sorted(old&new):
        try:
            a=json.loads((args.before/(ticker+'.json')).read_text(encoding='utf-8'))
            b=json.loads((args.after/(ticker+'.json')).read_text(encoding='utf-8'))
            change=compare(a,b)
            (args.output/(ticker+'.json')).write_text(json.dumps(change,ensure_ascii=False),encoding='utf-8')
            summary['companies'][ticker]={k:len(v) for k,v in change.items()}
        except Exception as exc:
            summary['companies'][ticker]=dict(error=type(exc).__name__)
    (args.output/'summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    print('Compared',len(old&new),'companies; missing',len(old^new),flush=True)
    return int(bool(old^new or any('error' in v for v in summary['companies'].values())))


if __name__=='__main__': raise SystemExit(main())
