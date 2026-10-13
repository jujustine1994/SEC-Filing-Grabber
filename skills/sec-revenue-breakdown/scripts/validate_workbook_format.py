"""Check the shared revenue-breakdown XLSX sheet contract without dependencies."""
import argparse
import xml.etree.ElementTree as ET
import zipfile
import json

EXPECTED = ['Quarterly History', 'Annual History', 'Coverage', 'Sources']

def validate(path):
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read('xl/workbook.xml'))
    actual = [s.attrib['name'] for s in root.findall('{*}sheets/{*}sheet')]
    if actual != EXPECTED:
        raise ValueError(f'Expected {EXPECTED}; got {actual}')
    return actual

def validate_with_plan(path,plan):
    """Strict gate for formula-v1 outputs. Numeric truth remains source QA's job."""
    validate(path)
    import openpyxl
    wb=openpyxl.load_workbook(path,data_only=False)
    errors=[];count=0
    try:
        for name,spec in plan['sheets'].items():
            ws=wb[name];expected={c['address']:c for c in spec['cells']}
            if ws.sheet_state!='visible':errors.append(name+': hidden sheet')
            if any(d.hidden for d in ws.row_dimensions.values()) or any(d.hidden for d in ws.column_dimensions.values()):errors.append(name+': hidden rows/columns')
            if ws.sheet_view.showGridLines is False:errors.append(name+': gridlines disabled')
            for addr,item in expected.items():
                cell=ws[addr];want=item.get('formula',item.get('value'))
                if cell.value!=want and not (want=='' and cell.value is None):errors.append(name+'!'+addr+': plan mismatch')
                if 'formula' in item:
                    count+=1
                    if cell.data_type!='f':errors.append(name+'!'+addr+': formula replaced with value')
                if cell.value is not None and (cell.font.name!='Microsoft JhengHei' or cell.font.sz is None or cell.font.sz<12):errors.append(name+'!'+addr+': typography')
                if item.get('kind') in {'share','qoq','yoy','op_margin','net_margin'} and '%' not in cell.number_format:errors.append(name+'!'+addr+': percent format')
            for row in ws:
                for cell in row:
                    if cell.value is not None and cell.coordinate not in expected:errors.append(name+'!'+cell.coordinate+': unplanned content')
        if wb['Sources']['A'+str(plan['raw_header'])].value!='Record ID':errors.append('Raw source header missing')
        if wb._external_links:errors.append('External links')
    finally:wb.close()
    if errors:raise ValueError('; '.join(errors[:20])+f' ({len(errors)} errors)')
    return dict(status='PASS',formula_count=count,strict_plan=True,sheet_count=4)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('xlsx')
    parser.add_argument('--plan',help='Validated workbook-plan.json; enable strict formula/layout checks')
    args = parser.parse_args()
    if args.plan:
        with open(args.plan,encoding='utf-8') as f:plan=json.load(f)
        print(json.dumps(validate_with_plan(args.xlsx,plan)))
    else:print('PASS:', ', '.join(validate(args.xlsx)))
