"""Own Excel verification only: save/recalculate isolated new B, never originals."""
from pathlib import Path
import sys,json,hashlib,datetime
import xlwings as xw
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from validate_workbook_format import validate_with_plan
file=Path(sys.argv[1]).resolve();plan=json.loads(Path(sys.argv[2]).read_text(encoding='utf-8'))
for existing_app in xw.apps:
 for existing_book in existing_app.books:
  if Path(existing_book.fullname).resolve()==file:
   raise RuntimeError('Verification requires a closed isolated workbook; close the test copy first.')
report=validate_with_plan(file,plan);report['native']={};app=None;book=None
try:
 app=xw.App(visible=False,add_book=False);app.api.DisplayAlerts=False;app.api.EnableEvents=False;app.api.AutomationSecurity=3
 book=app.books.open(str(file),update_links=False)
 book.sheets['Sources'].api.Calculate()
 for n in ['Quarterly History','Annual History','Coverage']:book.sheets[n].api.Calculate()
 errors=[];numeric=0;blanks=0;max_diff=0
 observed={}
 raw={c['address']:c.get('value') for c in plan['sheets']['Sources']['cells']}
 ids={x['value']:int(x['address'][1:]) for x in plan['sheets']['Sources']['cells'] if x['address'].startswith('A') and x.get('value')}
 for n in ['Quarterly History','Annual History']:
  ws=book.sheets[n];area=ws.used_range;values=area.api.Value2;start_row=area.row;start_col=area.column
  for c in plan['sheets'][n]['cells']:
   if 'formula' not in c:continue
   import re
   coord=re.fullmatch(r'([A-Z]+)(\d+)',c['address']);col=0
   for ch in coord[1]:col=col*26+ord(ch)-64
   val=values[int(coord[2])-start_row][col-start_col]
   observed[(n,c.get('kind'),c.get('period'),c.get('record_id'))]=val
   if isinstance(val,int) and val<=-2146826245 and val>=-2146826288:errors.append(n+'!'+c['address']+': error '+str(val));continue
   if isinstance(val,(int,float)):numeric+=1
   elif val in (None,''):blanks+=1
   if c.get('kind') in {'revenue','operating_income','net_income','total'}:
    import re
    refs=re.findall(r'Sources!\$O\$(\d+)',c['formula'])
    # Direct source formula expected exact; derived compare to recorded normalized result via ID.
    expected=raw.get('O'+str(ids[c['record_id']]))
    if expected is not None:
     if not isinstance(val,(int,float)):errors.append('Missing display '+c['record_id'])
     else:
      diff=abs(val-expected);max_diff=max(max_diff,diff)
      if diff>1e-7:errors.append('Value mismatch '+c['record_id'])
 assert not errors,errors[:10]
 if plan['company']['ticker']=='TEST':
  expected={('share','FY2025Q1'):0.1,('qoq','FY2025Q1'):-1/3,('yoy','FY2025Q1'):1,('op_margin','FY2025Q1'):0.25,('share','FY2024Q2'):0}
  for (kind,period),value in expected.items():
   actual=observed[('Quarterly History',kind,period,None)];assert abs(actual-value)<1e-10,(kind,actual,value)
  report['synthetic_arithmetic_checks']=len(expected)
 report['native']=dict(status='PASS',numeric_formula_results=numeric,blank_formula_results=blanks,max_source_difference=max_diff,error_count=len(errors))
 book.save()
 report.update(validate_with_plan(file,plan));report['sha256']=hashlib.sha256(file.read_bytes()).hexdigest()
finally:
 if book is not None:book.close()
 if app is not None:app.quit()
out=file.with_suffix('.verification.json');out.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(report))
