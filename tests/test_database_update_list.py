import json

import pytest
import local_db
from database import DatabaseError


def test_empty_list_never_reimports_old_config(isolated_database):
    local_db.set_update_list({}, [])
    assert local_db.get_update_list({'local_db_tickers': ['NVDA']}) == []


def test_list_survives_switching_app_config(isolated_database):
    local_db.add_tickers({}, ['NVDA', 'NO-DATA'])
    assert local_db.get_update_list({}) == ['NVDA', 'NO-DATA']
    assert local_db.remove_ticker({}, 'NVDA')
    assert local_db.get_update_list({'local_db_tickers': ['NVDA']}) == ['NO-DATA']


def test_invalid_uuid_in_list_is_not_used(isolated_database):
    path = isolated_database / 'metadata/update_list.json'
    data = json.loads(path.read_text())
    data['database_id'] = 'wrong'
    path.write_text(json.dumps(data))
    with pytest.raises(DatabaseError):
        local_db.get_update_list({})


def test_missing_list_does_not_repopulate_legacy_config(isolated_database):
    (isolated_database / 'metadata/update_list.json').unlink()
    with pytest.raises(DatabaseError):
        local_db.get_update_list({'local_db_tickers': ['NVDA']})
