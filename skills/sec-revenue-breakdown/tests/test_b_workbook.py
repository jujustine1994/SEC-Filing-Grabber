"""Behavior tests for reusable B plans, using explicitly synthetic disclosures."""
import copy, importlib.util, sys, unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from b_workbook import make_plan

def fixture():
    records=[]
    def add(rid,period,value,node='Cloud',metric='revenue',version='v1',basis='segment-reported'):
        records.append(dict(id=rid,period=period,duration='quarter',metric=metric,value=value,
          currency='USD',unit='USD millions',basis=basis,dimension='product',node_id=node,
          parent_id=None,original_label=node,classification_version=version,presentation='as_reported',
          source=dict(source_id='s1',url='https://www.sec.gov/example',locator='synthetic test table'),
          status='verified',derivation=dict(method='direct')))
    for p,val,total in [('FY2024Q1',10,100),('FY2024Q2',0,100),('FY2024Q4',30,100),('FY2025Q1',20,200)]:
        add('r'+p,p,val);add('t'+p,p,total,'Total',basis='GAAP')
        records[-1]['dimension']='consolidated'
    add('op1','FY2025Q1',5,metric='operating_income')
    data=dict(company=dict(ticker='TEST',name='Synthetic fixture',cik=1),records=records)
    row=dict(node='Cloud',parent=None,records={r['period']:r['id'] for r in records if r['node_id']=='Cloud' and r['metric']=='revenue'},op_records={'FY2025Q1':'op1'},annual_records={})
    views=dict(periods=['FY2024Q1','FY2024Q2','FY2024Q3','FY2024Q4','FY2025Q1'],annual_periods=[],blocks=[dict(dimension='product',version_id='display-v1',status='active',title='Products',rows=[row])])
    manifest=dict(cik=1,sources=[dict(source_id='s1',url='https://www.sec.gov/example',sha256='a'*64)])
    return data,views,manifest

class PlanTests(unittest.TestCase):
    def plan(self):return make_plan(*fixture(),template_id='hierarchy')
    def items(self,p,kind):return [c for c in p['sheets']['Quarterly History']['cells'] if c.get('kind')==kind]
    def test_source_formula_and_zero_not_blank(self):
        p=self.plan();self.assertTrue(self.items(p,'revenue'));self.assertIn('Sources!',self.items(p,'revenue')[0]['formula'])
        self.assertIn(0,[c.get('value') for c in p['sheets']['Sources']['cells']])
    def test_source_lookup_uses_record_id_not_row_position(self):
        p=self.plan();formula=self.items(p,'revenue')[0]['formula']
        self.assertIn('MATCH("rFY2024Q1"',formula)
        self.assertIn('COUNT(INDEX(',formula)
    def test_wildcard_record_id_is_escaped(self):
        d,v,m=fixture();d['records'][0]['id']='r?*~';v['blocks'][0]['rows'][0]['records']['FY2024Q1']='r?*~'
        p=make_plan(d,v,m);self.assertIn('MATCH("r~?~*~~"',self.items(p,'revenue')[0]['formula'])
    def test_ytd_cannot_be_quarterly_total(self):
        d,v,m=fixture();r=next(r for r in d['records'] if r['id']=='tFY2025Q1');r['duration']='ytd';r['source']['locator']=''
        p=make_plan(d,v,m);self.assertFalse(any(c['period']=='FY2025Q1' for c in self.items(p,'share')))
    def test_recast_requires_comparable_source(self):
        d,v,m=fixture();r=next(r for r in d['records'] if r['id']=='rFY2025Q1');r['presentation']='recast'
        p=make_plan(d,v,m);self.assertFalse(any(c['period']=='FY2025Q1' for c in self.items(p,'yoy')))
        r['published_in_accession']='official-recast';next(r for r in d['records'] if r['id']=='rFY2024Q1')['published_in_accession']='official-recast'
        p=make_plan(d,v,m);self.assertTrue(any(c['period']=='FY2025Q1' for c in self.items(p,'yoy')))
    def test_gap_does_not_use_previous_column(self):
        p=self.plan();qoq=self.items(p,'qoq');self.assertFalse(any(c['period']=='FY2024Q4' for c in qoq));self.assertTrue(any(c['period']=='FY2025Q1' for c in qoq))
    def test_changed_version_suppresses_growth(self):
        d,v,m=fixture();next(r for r in d['records'] if r['id']=='rFY2025Q1')['classification_version']='v2'
        p=make_plan(d,v,m);self.assertFalse(any(c['period']=='FY2025Q1' for c in self.items(p,'yoy')))
    def test_wrong_unit_suppresses_share(self):
        d,v,m=fixture();next(r for r in d['records'] if r['id']=='rFY2025Q1')['unit']='USD thousands'
        p=make_plan(d,v,m);self.assertFalse(any(c['period']=='FY2025Q1' for c in self.items(p,'share')))
    def test_wrong_basis_suppresses_margin(self):
        d,v,m=fixture();d['records'][-1]['basis']='non-GAAP'
        p=make_plan(d,v,m);self.assertFalse(any(c['period']=='FY2025Q1' for c in self.items(p,'op_margin')))
    def test_duplicate_record_rejected(self):
        d,v,m=fixture();d['records'].append(copy.deepcopy(d['records'][0]))
        with self.assertRaises(ValueError):make_plan(d,v,m)
    def test_missing_source_rejected(self):
        d,v,m=fixture();m['sources']=[]
        with self.assertRaises(ValueError):make_plan(d,v,m)
    def test_source_url_mismatch_rejected(self):
        d,v,m=fixture();d['records'][0]['source']['url']='https://unverified.example/'
        with self.assertRaises(ValueError):make_plan(d,v,m)
    def test_missing_parent_context_needs_record_evidence(self):
        d,v,m=fixture();v['blocks'][0]['rows'][0]['parent']='Official parent'
        for r in d['records']:
            if r['node_id']=='Cloud':r['parent_id']='Official parent'
        p=make_plan(d,v,m)
        self.assertTrue(any('階層參考' in str(c.get('value','')) for c in p['sheets']['Quarterly History']['cells']))
    def test_unsupported_parent_rejected(self):
        d,v,m=fixture();v['blocks'][0]['rows'][0]['parent']='Invented'
        with self.assertRaises(ValueError):make_plan(d,v,m)
    def test_unknown_template_rejected(self):
        with self.assertRaises(ValueError):make_plan(*fixture(),template_id='guess')
    def test_provisional_not_promoted(self):
        d,v,m=fixture();d['records'][0]['status']='provisional'
        p=make_plan(d,v,m);self.assertFalse(any(c['record_id']=='rFY2024Q1' for c in self.items(p,'revenue')))
    def test_negative_base_growth_blank(self):
        d,v,m=fixture();d['records'][0]['value']=-10
        p=make_plan(d,v,m);self.assertTrue(all('>0' in c['formula'] for c in self.items(p,'yoy')+self.items(p,'qoq')))
    def test_visible_fixed_four_sheets(self):
        self.assertEqual(list(self.plan()['sheets']),['Quarterly History','Annual History','Coverage','Sources'])
    def test_q4_legacy_nine_month_inputs_use_formula(self):
        d,v,m=fixture(); base=copy.deepcopy(d['records'][0]); base.update(id='annual',period='FY2024',duration='annual',value=100,fiscal_year=2024)
        ytd=copy.deepcopy(base); ytd.update(id='nine',period='FY2024YTD9',duration='9',value=70)
        d['records'] += [base,ytd]
        q4=next(r for r in d['records'] if r['id']=='rFY2024Q4');q4.update(fiscal_year=2024,derivation=dict(method='annual_minus_ytd',inputs=['annual','nine']))
        p=make_plan(d,v,m); formula=next(c['formula'] for c in self.items(p,'revenue') if c['period']=='FY2024Q4')
        self.assertIn(')-(',formula)
        base['period_start']='2024-01-01';base['period_end']='2024-12-31'
        ytd['period_start']='2023-01-01';ytd['period_end']='2024-09-30';q4['period_start']='2024-09-01';q4['period_end']='2024-12-31'
        with self.assertRaises(ValueError):make_plan(d,v,m)
        ytd['period_start']='2024-01-01';q4['period_start']='2024-10-01'
        make_plan(d,v,m)
        ytd['duration']='6'
        with self.assertRaises(ValueError):make_plan(d,v,m)
        ytd['duration']='9';ytd['period']='FY2022YTD9';base['period']='FY2023'
        for r in [base,ytd,q4]:r.pop('fiscal_year',None)
        with self.assertRaises(ValueError):make_plan(d,v,m)

    def test_company_and_manifest_cik_mismatch_rejected(self):
        d,v,m=fixture();m['cik']=2488
        with self.assertRaisesRegex(ValueError,'CIK'):make_plan(d,v,m)

    def test_missing_company_or_manifest_cik_rejected(self):
        for missing in ('company','manifest'):
            with self.subTest(missing=missing):
                d,v,m=fixture()
                (d['company'] if missing=='company' else m).pop('cik')
                with self.assertRaisesRegex(ValueError,'CIK'):make_plan(d,v,m)

    def test_invalid_cik_rejected(self):
        for value in (None,True,False,0,-1,10**10,1.0,'1e0','1.0','abc','１２３'):
            with self.subTest(value=value):
                d,v,m=fixture();m['cik']=value
                with self.assertRaisesRegex(ValueError,'CIK'):make_plan(d,v,m)

    def test_zero_padded_cik_matches_without_rewriting_inputs(self):
        d,v,m=fixture();m['cik']='0000000001'
        before=copy.deepcopy((d,v,m))
        make_plan(d,v,m)
        self.assertEqual((d,v,m),before)

    def test_used_sec_url_cik_must_match_manifest(self):
        d,v,m=fixture();url='https://www.sec.gov/Archives/edgar/data/2488/000000248826000123/example.htm'
        m['sources'][0]['url']=url
        for r in d['records']:r['source']['url']=url
        with self.assertRaisesRegex(ValueError,'CIK'):make_plan(d,v,m)

if __name__=='__main__':unittest.main()
