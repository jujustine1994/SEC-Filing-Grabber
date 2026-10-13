"""Plan four-sheet B workbooks from explicit evidence selections, not AI cells.

Existing breakdown + views remain authoritative inputs. A view chooses record
IDs; this module validates their identity and generates auditable Excel formulas.
It does not classify companies, calculate displayed metrics, or change inputs.
"""
import argparse, datetime, hashlib, json, math, re
from pathlib import Path
from urllib.parse import urlparse

SHEETS=['Quarterly History','Annual History','Coverage','Sources']
TEMPLATES={'flat','hierarchy','segment-op'}
RAW_FIELDS=['Record ID','Fiscal period','Period start','Period end','Duration','Metric','Original label','Parent','Dimension','Classification version','Presentation','Original value','Original unit','Currency','Normalized value','Normalized unit','Basis / geography definition','Filed / accession','Source ID','Source locator','Official URL','Method','Input record IDs','Verification','Original footnote / definition','Display cell']

def column(n):
    s=''
    while n:n,r=divmod(n-1,26);s=chr(65+r)+s
    return s

def prior(period,lag):
    m=re.fullmatch(r'FY(\d{4})Q([1-4])',period)
    if m:
        q=int(m[1])*4+int(m[2])-1-lag
        y,n=divmod(q,4);return f'FY{y}Q{n+1}'
    m=re.fullmatch(r'FY(\d{4})',period)
    return f'FY{int(m[1])-1}' if m else None

def compatible(a,b,version=True,basis=True):
    keys=['unit','currency']+(['basis'] if basis else [])
    if version:keys+=['dimension','node_id','classification_version']
    return all(a.get(k) is not None and a.get(k)==b.get(k) for k in keys)

def year_of(r):
    m=re.match(r'^FY(\d{4})',r['period'])
    if not m:raise ValueError('Invalid fiscal period: '+r['id'])
    year=int(m[1])
    if r.get('fiscal_year') is not None and r['fiscal_year']!=year:raise ValueError('Fiscal year/period mismatch: '+r['id'])
    return year

def presentation_comparable(a,b):
    recast=lambda r:r.get('presentation')=='recast'
    if recast(a)==recast(b):return True
    acc=a.get('published_in_accession')
    return bool(acc and acc==b.get('published_in_accession'))

def same_interval(a,b):
    if a['duration']!=b['duration'] or a['period']!=b['period']:return False
    return all(not (a.get(k) and b.get(k)) or a[k]==b[k] for k in ['period_start','period_end'])

def normalized_cik(value):
    if isinstance(value,bool):raise ValueError('CIK must be a positive SEC company identifier')
    if isinstance(value,str) and re.fullmatch(r'[0-9]{1,10}',value.strip()):value=int(value.strip())
    if not isinstance(value,int) or not 0<value<10**10:
        raise ValueError('Missing or invalid CIK; confirm the company before building B')
    return value

def make_plan(data,views,manifest,template_id='hierarchy'):
    if template_id not in TEMPLATES:raise ValueError('Unknown template')
    company=data['company'];records=data['records'];lookup={}
    cik=normalized_cik(company.get('cik'))
    if cik!=normalized_cik(manifest.get('cik')):
        raise ValueError('Company CIK differs from the source manifest CIK')
    sources={s['source_id']:s for s in manifest['sources']}
    for r in records:
        if r['id'] in lookup:raise ValueError('Duplicate record ID: '+r['id'])
        if r['source']['source_id'] not in sources:raise ValueError('Missing source: '+r['id'])
        if r['source'].get('url')!=sources[r['source']['source_id']].get('url'):raise ValueError('Source URL mismatch: '+r['id'])
        url=urlparse(sources[r['source']['source_id']].get('url',''))
        if url.hostname in {'www.sec.gov','data.sec.gov'}:
            identity=re.match(r'^/Archives/edgar/data/([0-9]+)/',url.path)
            if identity and normalized_cik(identity[1])!=cik:
                raise ValueError('SEC source URL CIK differs from the company CIK: '+r['id'])
        value=r.get('value')
        if value is not None and (not isinstance(value,(int,float)) or isinstance(value,bool) or not math.isfinite(value)):raise ValueError('Non-finite/non-numeric value: '+r['id'])
        lookup[r['id']]=r
    plan=dict(schema_version='b-formula-1',template_id=template_id,template_version=1,
              company=company,sheets={n:dict(cells=[],bands=[],last_row=1,last_col=2) for n in SHEETS},
              unresolved=[],display_records=[])
    def put(sheet,row,col,value=None,formula=None,**meta):
        obj=dict(address=f'{column(col)}{row}',**meta)
        if formula is not None:obj['formula']=formula
        else:obj['value']=value
        plan['sheets'][sheet]['cells'].append(obj)
        s=plan['sheets'][sheet];s['last_row']=max(s['last_row'],row);s['last_col']=max(s['last_col'],col)
        return obj
    for c,v in enumerate(['Source ID','Form','Document type','Filed','SEC official URL','SHA-256'],1):put('Sources',1,c,v)
    for row,s in enumerate(manifest['sources'],2):
        for col,v in enumerate([s['source_id'],s.get('form'),s.get('document_type'),s.get('filing_date'),s.get('url'),s.get('sha256')],1):put('Sources',row,col,v)
    raw_header=len(manifest['sources'])+5
    plan['raw_header']=raw_header
    for c,v in enumerate(RAW_FIELDS,1):put('Sources',raw_header,c,v)
    source_rows={};source_display_cells={}
    for row,r in enumerate(records,raw_header+1):
        source_rows[r['id']]=row
        values=[r['id'],r['period'],r.get('period_start'),r.get('period_end'),r['duration'],r['metric'],r.get('original_label'),r.get('parent_id'),r['dimension'],r['classification_version'],r.get('presentation'),r.get('original_value'),r.get('original_unit'),r.get('currency'),r.get('value'),r.get('unit'),r.get('basis'),str(r.get('filing_date',''))+' / '+str(r.get('published_in_accession','')),r['source']['source_id'],r['source'].get('locator'),r['source'].get('url'),r.get('derivation',{}).get('method'),json.dumps(r.get('derivation',{}).get('inputs',[]),ensure_ascii=False),r.get('status'),r.get('original_footnote'),None]
        for c,v in enumerate(values,1):
            cell=put('Sources',row,c,v)
            if c==26:source_display_cells[r['id']]=cell
    def record_formula(r,stack=()):
        rid=r['id'];method=r.get('derivation',{}).get('method','direct')
        if rid in stack:raise ValueError('Cyclic derivation: '+rid)
        if method=='direct':
            start=raw_header+1;end=raw_header+len(records);escaped=rid.replace('~','~~').replace('*','~*').replace('?','~?').replace('"','""')
            addr=f'INDEX(Sources!$O${start}:$O${end},MATCH("{escaped}",Sources!$A${start}:$A${end},0))'
            # COUNT on the INDEX reference distinguishes an empty source cell
            # from an officially reported zero; exact ID survives source sorting.
            return f'IF(COUNT({addr})=1,{addr},"")'
        if method in {'annual_minus_ytd','annual_minus_9m'}:
            inputs=r.get('derivation',{}).get('inputs',[])
            ids=[i if isinstance(i,str) else i.get('id',i.get('record_id')) for i in inputs]
            if len(ids)!=2 or any(i not in lookup for i in ids):raise ValueError('Derived Q4 lacks two inputs: '+rid)
            a,b=(lookup[i] for i in ids)
            nine_months=b['duration']=='9' or (b['duration']=='ytd' and re.fullmatch(r'FY\d{4}YTD9',b['period']))
            year=year_of(r)
            dates_ok=not (a.get('period_end') and r.get('period_end')) or a['period_end']==r['period_end']
            if b.get('period_end') and a.get('period_end'):dates_ok=dates_ok and b['period_end']<a['period_end']
            if a.get('period_start') and b.get('period_start'):dates_ok=dates_ok and a['period_start']==b['period_start']
            if r.get('period_start') and b.get('period_end'):
                dates_ok=dates_ok and datetime.date.fromisoformat(r['period_start'])==datetime.date.fromisoformat(b['period_end'])+datetime.timedelta(days=1)
            if a['duration']!='annual' or not nine_months or not r['period'].endswith('Q4') or a['period']!=f'FY{year}' or b['period']!=f'FY{year}YTD9' or year_of(a)!=year or year_of(b)!=year or not dates_ok or not presentation_comparable(a,b) or not compatible(a,b) or not compatible(a,r) or a.get('metric')!=r.get('metric') or b.get('metric')!=r.get('metric') or any(x.get('status')!='verified' for x in (a,b)):
                raise ValueError('Incompatible Q4 derivation: '+rid)
            fa,fb=record_formula(a,stack+(rid,)),record_formula(b,stack+(rid,))
            return f'IF(COUNT({fa},{fb})=2,({fa})-({fb}),"")'
        raise ValueError('Unsupported derivation: '+method+' '+rid)
    # Select total revenue only when evidence is unique or explicitly selected.
    totals={}
    specified=views.get('consolidated_records',{})
    for p in views['periods']+views.get('annual_periods',[]):
        duration='quarter' if re.fullmatch(r'FY\d{4}Q[1-4]',p) else ('annual' if re.fullmatch(r'FY\d{4}',p) else 'transition')
        candidates=[r for r in records if r['period']==p and r['duration']==duration and r['dimension']=='consolidated' and r['metric']=='revenue' and r['status']=='verified' and r['value'] is not None and r['source'].get('locator')]
        if p in specified:
            candidates=[r for r in candidates if r['id']==specified[p]]
            if len(candidates)!=1:raise ValueError('Invalid total selection: '+p)
        if candidates:
            contexts={(r['value'],r.get('unit'),r.get('currency'),r.get('basis'),r.get('duration')) for r in candidates}
            if len(contexts)>1:
                plan['unresolved'].append('Conflicting consolidated total: '+p);continue
            totals[p]=sorted(candidates,key=lambda r:(r.get('filing_date',''),r['id']))[-1]
    display_refs={}
    def select(ids,p,block,node,metric,annual):
        rid=ids.get(p)
        if not rid:return None
        if rid not in lookup:raise ValueError('Unknown selected record: '+rid)
        r=lookup[rid]
        if r['period']!=p or r['dimension']!=block['dimension'] or r['node_id']!=node or r['metric']!=metric:raise ValueError('View identity mismatch: '+rid)
        if r['duration'] not in ({'annual','transition'} if annual else {'quarter'}):raise ValueError('View duration mismatch: '+rid)
        if r.get('status')!='verified' or r.get('value') is None:
            plan['unresolved'].append('Selected record not verified/non-null: '+rid);return None
        if not r['source'].get('locator'):raise ValueError('Missing source locator: '+rid)
        return r
    for annual,sheet in enumerate(SHEETS[:2]):
        periods=views.get('annual_periods',[]) if annual else views['periods']
        if len(periods)!=len(set(periods)) or periods!=sorted(periods):raise ValueError('Duplicate/unordered periods')
        put(sheet,2,2,f"{company['ticker']} {company.get('name','')} 營收拆分 — {'年度' if annual else '季度'}")
        units=sorted({str(r.get('unit')) for r in records if r.get('value') is not None and r.get('metric') in {'revenue','operating_income','net_income'}})
        unit_note=' / '.join(units) if len(units)<=3 else '多種原始單位，逐筆見 Sources'
        put(sheet,3,2,unit_note+'；占公司總營收。空白不等於零；不同維度／版本不可相加。')
        put(sheet,5,2,'分類／指標')
        for c,p in enumerate(periods,3):put(sheet,5,c,p)
        total_row=6
        put(sheet,total_row,2,'公司總營收（官方揭露）')
        for c,p in enumerate(periods,3):
            if p in totals:
                r=totals[p];put(sheet,total_row,c,formula='='+record_formula(r),kind='total',period=p,record_id=r['id']);display_refs.setdefault(r['id'],[]).append(f'{sheet}!{column(c)}{total_row}')
        row=8
        blocks=views['blocks']
        if any(b.get('status') not in {'active','inactive',None} for b in blocks):raise ValueError('Invalid lifecycle status')
        blocks=sorted(blocks,key=lambda b:{'active':0,None:1,'inactive':2}[b.get('status')])
        for block in blocks:
            block=dict(block,rows=[dict(r) for r in block['rows']])
            present={r['node'] for r in block['rows']}
            contexts=[]
            for nr in block['rows']:
                parent=nr.get('parent')
                if parent and parent not in present:
                    ids=nr.get('records',{})|nr.get('annual_records',{})
                    rr=[lookup[rid] for rid in ids.values() if rid in lookup]
                    if not rr or any(r.get('parent_id')!=parent or r['dimension']!=block['dimension'] for r in rr):raise ValueError('Missing supported display parent: '+parent)
                    contexts.append(dict(node=parent,parent=None,context_only=True));present.add(parent)
            block['rows']=contexts+block['rows']
            nodes=[r['node'] for r in block['rows']]
            if len(nodes)!=len(set(nodes)):raise ValueError('Duplicate display node')
            parent_map={r['node']:r.get('parent') for r in block['rows']}
            depths={}
            def depth(n,trail=()):
                if n in trail:raise ValueError('Cyclic display hierarchy')
                parent=parent_map[n]
                if parent is None:return 0
                if parent not in parent_map:raise ValueError('Missing display parent: '+parent)
                return 1+depth(parent,trail+(n,))
            for n in nodes:depths[n]=depth(n)
            if template_id=='flat' and any(depths.values()):raise ValueError('Flat template cannot flatten disclosed hierarchy')
            # Keep explicit existing section identities for C's local mappings.
            # A section title is not a node ID; do not silently rename it.
            prefix=('Inactive Segments — ' if block.get('status')=='inactive' and not block['title'].startswith('Inactive Segments') else '') if block.get('status') is not None else '狀態待確認 — '
            put(sheet,row,2,prefix+block['title']);plan['sheets'][sheet]['bands'].append(dict(row=row,status=block.get('status')));row+=1
            for nodeview in block['rows']:
                node=nodeview['node'];parent=nodeview.get('parent');dep=depths[node];label='    '*dep+node
                if nodeview.get('context_only'):
                    put(sheet,row,2,label+'（階層參考，無本版數值）');row+=1;continue
                put(sheet,row,2,label);revrow=row
                ids=nodeview.get('annual_records' if annual else 'records',{});selected={}
                for c,p in enumerate(periods,3):
                    r=select(ids,p,block,node,'revenue',bool(annual))
                    if not r:continue
                    # Existing NVDA display aggregates versions; each record still owns its actual identity.
                    if parent and r.get('parent_id')!=parent:raise ValueError('Display parent lacks selected record support: '+r['id'])
                    selected[p]=r
                    put(sheet,row,c,formula='='+record_formula(r),kind='revenue',period=p,record_id=r['id'])
                    display_refs.setdefault(r['id'],[]).append(f'{sheet}!{column(c)}{row}')
                row+=1
                metrics=[('share','占公司總營收',None),*([('qoq','QoQ',1)] if not annual else []),('yoy','YoY',4 if not annual else 1)]
                for kind,title,lag in metrics:
                    put(sheet,row,2,'    '*(dep+1)+title)
                    for c,p in enumerate(periods,3):
                        r=selected.get(p)
                        if not r or r['duration']=='transition':continue
                        curr=f'{column(c)}{revrow}';formula=None
                        if kind=='share':
                            total=totals.get(p)
                            if total and same_interval(r,total) and presentation_comparable(r,total) and compatible(r,total,False,False):
                                den=f'{column(c)}${total_row}';formula=f'=IF(COUNT({curr},{den})=2,IF({den}>0,{curr}/{den},""),"")'
                        else:
                            pp=prior(p,lag);prev=selected.get(pp)
                            dates_ok=not (r.get('period_end') and prev and prev.get('period_end')) or r['period_end']>prev['period_end']
                            if prev and prev['duration']==r['duration'] and compatible(r,prev) and presentation_comparable(r,prev) and dates_ok:
                                old=f'{column(periods.index(pp)+3)}{revrow}';formula=f'=IF(COUNT({curr},{old})=2,IF({old}>0,{curr}/{old}-1,""),"")'
                        if formula:put(sheet,row,c,formula=formula,kind=kind,period=p)
                    row+=1
                for metric,key,title,margin in [('operating_income','op_records','Operating income','op_margin'),('net_income','net_records','Net income','net_margin')]:
                    mids=nodeview.get(('annual_'+key) if annual else key,{})
                    if not mids:continue
                    put(sheet,row,2,'    '*(dep+1)+title);income_row=row;profits={}
                    for c,p in enumerate(periods,3):
                        r=select(mids,p,block,node,metric,bool(annual))
                        if not r:continue
                        profits[p]=r;put(sheet,row,c,formula='='+record_formula(r),kind=metric,period=p,record_id=r['id']);display_refs.setdefault(r['id'],[]).append(f'{sheet}!{column(c)}{row}')
                    row+=1;put(sheet,row,2,'    '*(dep+1)+('OP%' if margin=='op_margin' else 'Net margin'))
                    for c,p in enumerate(periods,3):
                        r=selected.get(p);op=profits.get(p)
                        if r and op and same_interval(r,op) and presentation_comparable(r,op) and compatible(r,op):
                            rev=f'{column(c)}{revrow}';profit=f'{column(c)}{income_row}'
                            put(sheet,row,c,formula=f'=IF(COUNT({rev},{profit})=2,IF({rev}>0,{profit}/{rev},""),"")',kind=margin,period=p)
                    row+=1
                # No blank separator inside a classification subtree: existing
                # C treats unindented blanks as its parent-boundary terminator.
            row+=1
    for rid,refs in display_refs.items():
        source_display_cells[rid]['value']='; '.join(refs)
    coverage=[['B template',f'{template_id} / formula-v1'],['Company',company['ticker']],['Classification','按官方來源與既有 registry；latest/inactive 只是顯示，缺季不代表停用。'],['Calculations','主表來源公式；比重／QoQ／YoY／利益率由 Excel 計算。前期或分母≤0留空。'],['Growth comparability','同分類版本／節點／單位／幣別／口徑；過渡期間不算成長，缺期不跳期。'],['Evidence limitations','沿用來源狀態；verified 不代表全部歷史完整。無來源關係／狀態仍待確認。'],['Original raw strings','原始字串／註腳未擷取時留空，不冒充標準化數字為原文。'],['Selected records',len(display_refs)],['Unresolved count',len(plan['unresolved'])]]
    coverage+= [['Unresolved',m] for m in plan['unresolved']]
    for row,vals in enumerate(coverage,2):
        for c,v in enumerate(vals,2):put('Coverage',row,c,v)
    rr=len(coverage)+4
    for c,v in enumerate(['Version ID','Dimension','Status','Original definition','First observed','Last observed','Evidence / limitations'],2):put('Coverage',rr,c,v)
    for i,b in enumerate(views['blocks'],rr+1):
        for c,v in enumerate([b['version_id'],b['dimension'],b.get('status'),b['title'],b.get('first_observed'),b.get('last_observed'),b.get('note','Display selection from views; observed dates are not effective dates')],2):put('Coverage',i,c,v)
    plan['display_records']=[dict(record_id=k,cells=v) for k,v in display_refs.items()]
    return plan

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--breakdown',required=True,type=Path);p.add_argument('--views',required=True,type=Path);p.add_argument('--manifest',required=True,type=Path);p.add_argument('--template-id',choices=sorted(TEMPLATES),default='hierarchy');p.add_argument('--out',required=True,type=Path)
    args=p.parse_args()
    load=lambda f:json.loads(f.read_text(encoding='utf-8'))
    plan=make_plan(load(args.breakdown),load(args.views),load(args.manifest),args.template_id)
    plan['input_sha256']={str(f):hashlib.sha256(f.read_bytes()).hexdigest() for f in (args.breakdown,args.views,args.manifest)}
    args.out.parent.mkdir(parents=True,exist_ok=True)
    if args.out.exists():raise FileExistsError(args.out)
    args.out.write_text(json.dumps(plan,ensure_ascii=False,indent=2),encoding='utf-8')
    print('PLAN',args.out,'formulas',sum('formula' in c for s in plan['sheets'].values() for c in s['cells']))

if __name__=='__main__':main()
