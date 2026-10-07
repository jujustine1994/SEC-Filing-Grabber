"""Exercise the real Tk scan/queue/ticker flow with isolated data and no SEC I/O."""
import json
import os
from pathlib import Path
import sys
import tempfile
import threading
import time
from unittest.mock import patch

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'src'))
from database import create_database


def main():
    import tkinter as tk
    with tempfile.TemporaryDirectory(prefix='sec-preview-probe-') as tmp:
        db=Path(tmp)/'database'
        create_database(db,test_mode=True)
        os.environ['SEC_LOCAL_DB_ROOT']=str(db)
        os.environ['SEC_CONFIG_PATH']=str(Path(tmp)/'config.json')
        import main as gui
        import fetcher_gaap as fg
        import i18n
        root=tk.Tk();root.withdraw()
        results={}
        release=threading.Event()
        try:
            with patch.object(gui,'_write_log_header'),patch.object(gui,'_write_log'):
                app=gui.SECFetcherApp(root)
                app.cfg['identity']='Probe probe@example.com'
                app._ensure_database=lambda:True
                app._confirm_company=lambda:None
                i18n.set_lang('zh_tw')
                def result(name,estimated=False):
                    return dict(sheets=['Data_Financials(Q)','Data_Financials(Y)','Data_Meta','Data_Seg_'+name],
                                latest_label='FY2025Q1',latest_period_end='2025-05-03',
                                filing_date='2025-06-01',label_estimated=estimated)
                def pump():
                    until=time.monotonic()+10
                    while app._scan_running and time.monotonic()<until:
                        root.update();time.sleep(.02)
                    assert not app._scan_running,'worker/queue did not complete'
                    root.update()
                app.ticker_var.set('OLD')
                app._show_preview_result('OLD',result('OLD'))
                old=app._sheet_check_vars['Data_Seg_OLD'];old.set(False)
                app.ticker_var.set('NEW')
                results['ticker_change_clears_exclusions']=not app._sheet_check_vars and app._sheet_panel_ticker is None
                def slow(ticker,*args,**kwargs):
                    assert release.wait(10)
                    return result(ticker)
                app.ticker_var.set('OLD')
                with patch.object(fg,'preview_sheets',side_effect=slow):
                    app._run_preview_scan()
                    results['button_starts_worker_and_locks_scan']=app._scan_running and str(app._scan_btn['state'])=='disabled'
                    app.ticker_var.set('NEW')
                    release.set();pump()
                results['stale_worker_result_discarded']=not app._sheet_check_vars and app._sheet_panel_ticker is None
                with patch.object(fg,'preview_sheets',return_value=result('NEW',True)):
                    app._run_preview_scan();pump()
                results['current_result_applied']=app._sheet_panel_ticker=='NEW' and 'Data_Seg_NEW' in app._sheet_check_vars
                results['estimated_identity_visible']='估計' in str(app._sheet_panel_frame['text'])
                results['button_restored_after_queue']=str(app._scan_btn['state'])=='normal'
                app.is_running=True
                app._finish_preview_scan()
                results['active_fetch_lock_retained']=str(app._scan_btn['state'])=='disabled'
                app.is_running=False
                app._finish_preview_scan()
                assert all(results.values()),results
        finally:
            release.set()
            root.destroy()
        output=ROOT/'output/gui-rules-tk-probe.json'
        output.write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf8')
        print(json.dumps(results,ensure_ascii=False))


if __name__=='__main__':main()
