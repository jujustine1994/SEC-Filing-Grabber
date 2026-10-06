import argparse,json,re,sys
from pathlib import Path
from openpyxl import load_workbook
parser=argparse.ArgumentParser();parser.add_argument('folder',type=Path)
args=parser.parse_args()
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'src'))
from fiscal_input import fiscal_quarter_of,calendar_quarter_of
from fetcher_gaap import StatementTable
from excel_writer import _excel_sheet_names
from excel_formatter import unit_format_for,SECTION_HEADERS
source=json.loads((args.folder/'source-bundle.json').read_text(encoding='utf-8'))
overrides=json.loads((args.folder/'excel-recalculation.json').read_text(encoding='utf-8-sig'))
results={}
def equal(a,b):
    if a==b or a is None and b=='':return True
    return isinstance(a,(int,float)) and isinstance(b,(int,float)) and abs(a-b)<=max(1e-8,abs(b)*1e-14)
for ticker,tables in source.items():
    wb=load_workbook(args.folder/(ticker+'.xlsx'),read_only=True,data_only=True)
    errors=[];inconclusive=[];checked=0;ratios=0;headers=0
    export_tables = _excel_sheet_names([StatementTable(**t) for t in tables], [])
    for table, exported in zip(tables, export_tables):
        matrix=list(wb[exported.sheet_name].values)
        financial=table['sheet_name'] in ('Data_Financials(Q)','Data_Financials(Y)')
        annual=table['sheet_name']=='Data_Financials(Y)'
        for col,label in enumerate(table['quarter_labels']):
            checked+=1;headers+=1
            if matrix[0][col+3]!=label:errors.append((table['sheet_name'],'label',col,matrix[0][col+3],label))
        for row,concept in enumerate(table['concepts']):
            for col,expected in enumerate(table['values'][row]):
                if financial and concept=='Calendar Quarter':
                    end=table['period_ends'][col]
                    if end:expected=calendar_quarter_of(end,basis='end')
                    headers+=1
                elif financial and concept=='Fiscal Quarter' and not annual:
                    end=table['period_ends'][col]
                    if end:expected=re.sub(r'Q([1-4])$',r'FQ\1',table['quarter_labels'][col])
                    headers+=1
                elif table['sheet_name']!='Data_Meta' and concept.strip() not in SECTION_HEADERS and concept.strip() and isinstance(expected,(int,float)):
                    expected/=unit_format_for(concept.strip())[1]
                actual=matrix[row+2][col+3];checked+=1
                ratios+=table['sheet_name']=='Data_Ratios'
                if not equal(actual,expected):errors.append((table['sheet_name'],concept,col,actual,expected))
    wb.close()
    override=next(r for r in overrides if r['ticker']==ticker)
    for cell in override['overrideHeaders']:
        if not cell['end']:continue
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', str(cell['end'])):
            inconclusive.append(cell)
            continue
        label=fiscal_quarter_of(cell['end'],int(override['overrideMonth']))
        annual=cell['sheet']=='Data_Financials(Y)'
        if annual:label=label.split('Q')[0]
        expected={'label':label,'calendar':calendar_quarter_of(cell['end'],basis='end')}
        if not annual:expected['fiscal']=label.replace('Q','FQ')
        for key,value in expected.items():
            checked+=1;headers+=1
            if cell[key]!=value:errors.append((cell['sheet'],'override '+key,cell['column'],cell[key],value))
    results[ticker]=dict(cells_checked=checked,ratio_cells=ratios,header_cells=headers,errors=errors,inconclusive_override_headers=inconclusive)
    print(ticker,checked,'cells;',ratios,'ratio cells;',headers,'headers;',len(errors),'errors;',len(inconclusive),'inconclusive partial-date columns')
(args.folder/'verification.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
raise SystemExit(int(any(r['errors'] for r in results.values())))




