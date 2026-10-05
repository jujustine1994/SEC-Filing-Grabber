"""Compare saved pipeline outputs with independently inspected original SEC facts."""
import argparse
import json
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('audit_directory',type=Path)
    args=parser.parse_args()
    evidence=json.loads((ROOT/'docs/superpowers/evidence/fiscal-source-examples.json').read_text(encoding='utf-8'))
    failures=[];checked=0
    for case in evidence['cases']:
        result=json.loads((args.audit_directory/(case['ticker']+'.json')).read_text(encoding='utf-8'))
        table=next(t for t in result['tables'] if t['sheet_name']=='Data_Financials(Q)')
        columns=[i for i,end in enumerate(table['period_ends']) if end==case['end']]
        if len(columns)!=1:
            failures.append((case['ticker'],case['end'],'Expected one quarter at this end date'));continue
        col=columns[0]
        label=f"FY{case['fiscal_year']}Q4"
        if table['quarter_labels'][col]!=label:
            failures.append((case['ticker'],case['end'],'Wrong fiscal label',table['quarter_labels'][col],label))
        ocf=case['annual_ocf']-case['nine_month_ocf']
        capex=case['annual_capex']-case['nine_month_capex']
        expected={'Revenue':case['annual_revenue']-case['nine_month_revenue'],
                  'Operating Cash Flow':ocf,'Capex':capex,
                  'Free Cash Flow':ocf-abs(capex),'Ending Cash':case['ending_cash']}
        for concept,value in expected.items():
            rows=[i for i,c in enumerate(table['concepts']) if c==concept]
            actual=table['values'][rows[0]][col] if len(rows)==1 else None
            # SEC payment facts are positive outflow magnitudes. Presentation
            # tables may apply a negative preferred sign; FCF uses magnitude.
            if concept=='Capex' and isinstance(actual,(int,float)):actual=abs(actual)
            checked+=1
            if actual!=value:failures.append((case['ticker'],case['end'],concept,actual,value))
        print(case['ticker'],case['end'],label,expected)
    print('Checked',checked,'source-derived financial cells; failures',len(failures))
    for failure in failures:print(failure)
    return int(bool(failures))


if __name__=='__main__':raise SystemExit(main())
