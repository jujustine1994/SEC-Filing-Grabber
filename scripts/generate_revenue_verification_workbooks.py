import argparse,json,sys
from dataclasses import asdict
from pathlib import Path
parser=argparse.ArgumentParser()
parser.add_argument('--code',type=Path,required=True)
parser.add_argument('--source',type=Path,required=True)
parser.add_argument('--output',type=Path,required=True)
parser.add_argument('--tickers',nargs='+',default=['MAR','HLT','AMZN','MSFT','AFL','C','COST'])
args=parser.parse_args();sys.path.insert(0,str(args.code/'src'))
from fetcher_gaap import StatementTable
from excel_writer import write_statements
from ratios import build_ratio_table
args.output.mkdir(parents=True,exist_ok=True)
bundle={}
for ticker in args.tickers:
    path=args.output/(ticker+'.xlsx')
    if path.exists():raise RuntimeError('Creates only new owned verification workbooks')
    data=json.loads((args.source/(ticker+'.json')).read_text(encoding='utf-8'))
    tables=[StatementTable(**t) for t in data['tables']]
    ratio=build_ratio_table(next(t for t in tables if t.sheet_name=='Data_Financials(Q)'))
    if ratio is not None:tables.append(ratio)
    write_statements(tables,path)
    bundle[ticker]=[asdict(t) for t in tables]
    print(ticker,'written including ratios',flush=True)
(args.output/'source-bundle.json').write_text(json.dumps(bundle,ensure_ascii=False),encoding='utf-8')

