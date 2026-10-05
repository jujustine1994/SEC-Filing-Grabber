import json
import sys
from pathlib import Path
import pytest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import audit_fiscal_pipeline as audit
import local_db


def test_full_audit_can_require_official_listing_instead_of_cover_substitution(tmp_path,monkeypatch):
    monkeypatch.setattr(local_db,'read_meta',lambda ticker:dict(cik=1))
    with pytest.raises(RuntimeError,match='Official SEC filing metadata required'):
        audit.audit('TEST',tmp_path,require_official_metadata=True)


@pytest.mark.parametrize('listing,message',[
    ([], 'Accession missing'),
    ([dict(accession_number='0000000001-25-000001',reportDate='')], 'Invalid official'),
])
def test_required_official_metadata_never_falls_back_to_cover(tmp_path,monkeypatch,listing,message):
    accession='0000000001-25-000001'
    (tmp_path/(accession+'.json')).write_text('{}',encoding='utf-8')
    (tmp_path.parent/'TEST-sec-listing.json').write_text(json.dumps(listing),encoding='utf-8')
    monkeypatch.setattr(local_db,'read_meta',lambda ticker:dict(cik=1))
    monkeypatch.setattr(audit.fc,'ticker_dir',lambda ticker:tmp_path)
    monkeypatch.setattr(audit.fc,'load_filing',lambda *a:dict(form='10-Q',filing_date='2025-04-30'))
    with pytest.raises(RuntimeError,match=message):
        audit.audit('TEST',tmp_path,require_official_metadata=True)


@pytest.mark.parametrize('cik,keys',[(2,list(audit.fc.STATEMENT_KEYS)),(1,['income_statement'])])
def test_audit_rejects_inputs_the_real_cache_gate_rejects(tmp_path,monkeypatch,cik,keys):
    accession='0000000001-25-000001'
    entry=dict(schema_version=audit.fc.SCHEMA_VERSION,edgartools_version='test',
               cik=cik,accession_no=accession,form='10-Q',filing_date='2025-04-30',
               has_financials=True,fetched_keys=keys,dataframes={})
    (tmp_path/(accession+'.json')).write_text(json.dumps(entry),encoding='utf-8')
    monkeypatch.setattr(audit.fc,'ticker_dir',lambda ticker:tmp_path)
    monkeypatch.setattr(audit.fc,'filing_path',lambda ticker,acc:tmp_path/(acc+'.json'))
    monkeypatch.setattr(audit.fc,'edgartools_version',lambda:'test')
    monkeypatch.setattr(local_db,'read_meta',lambda ticker:dict(cik=1))
    table=audit.fg.StatementTable(sheet_name='Data_Financials(Q)',quarter_labels=[],filing_dates=[],concepts=[],values=[])
    monkeypatch.setattr(audit.fg,'fetch_gaap_statements',lambda *a,**kw:[table])
    with pytest.raises(RuntimeError):
        audit.audit('TEST',tmp_path)
