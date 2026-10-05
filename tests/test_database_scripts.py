import importlib.util
import os
from pathlib import Path
import subprocess

import pytest

ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.skipif(os.name != 'nt', reason='Windows launcher')
@pytest.mark.parametrize('name', ['overnight_update_db.ps1', 'pack.ps1'])
def test_powershell_script_parses_and_overnight_binds_parameters(name):
    script = ROOT / 'scripts' / name
    assert not script.read_bytes().startswith(b'\xef\xbb\xbf\xef\xbb\xbf')
    command = "$e=$null;$t=$null;$a=[System.Management.Automation.Language.Parser]::ParseFile($args[0],[ref]$t,[ref]$e);if($e.Count -or -not $a.ParamBlock){exit 1}"
    # Argument passing through -Command is awkward in Windows PowerShell;
    # use a literal single-quoted path with PowerShell's quote escaping.
    command = command.replace('$args[0]', "'" + str(script).replace("'", "''") + "'")
    if name == 'pack.ps1':
        command = command.replace('$e.Count -or -not $a.ParamBlock', '$e.Count')
    result = subprocess.run(['powershell','-NoProfile','-Command',command], capture_output=True)
    assert result.returncode == 0, result.stderr.decode(errors='replace')


def test_health_check_uses_connected_inventory(monkeypatch, isolated_database, tmp_path):
    spec = importlib.util.spec_from_file_location('health_check_probe', ROOT / 'scripts/health_check.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    (isolated_database / 'filings' / 'NEWCO').mkdir()
    seen = []
    def check(ticker, identity):
        seen.append(ticker)
        return dict(ticker=ticker, sheets={}, elapsed=0, gaps=[])
    monkeypatch.setattr(module, 'check_one', check)
    monkeypatch.setattr(module, 'OUT', tmp_path / 'health')
    monkeypatch.setattr(module, 'build_report', lambda: 0)
    module.main([])
    assert seen == ['NEWCO']
