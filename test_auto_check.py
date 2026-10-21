import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch

from auto_check import previous_months, inherit_settings, check_history, show_auto_check
from core import Store, load_workbook, origin
from drive_archive import build_snapshot
from test_core import rec, fx
from test_drive_archive import source


class AutoTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.old = Store(Path(self.tmp.name)/'old.db')
        self.new = Store(Path(self.tmp.name)/'new.db')
    def tearDown(self):
        self.tmp.cleanup()
    def snapshot(self, raw, months):
        return build_snapshot(raw, 'source.xlsx', self.old, months, rate_provider=fx)
    def test_calendar_rollover(self):
        self.assertEqual(previous_months(date(2026,1,3)), ['2025-10','2025-11','2025-12'])
    def test_assignment_survives_amount_and_date_changes_but_alerts(self):
        r = rec(dates='2026-07-01, 2026-08-01')
        self.old.save_decision(r,'2026-07')
        old = self.snapshot(source(r), ['2026-07'])
        inherit_settings(self.new,[old])
        for changed in [rec(margin='120',dates='2026-07-01, 2026-08-01'),
                        rec(dates='2026-08-01')]:
            raw = source(changed)
            loaded = load_workbook(raw)[1][0]
            self.assertEqual(origin(loaded,self.new),'2026-07')
            result = check_history(raw,'new.xlsx',self.new,[old],['2026-07'],fx)[0]['result']
            self.assertEqual(len(result['changes']),1)
            self.assertNotIn('Hónapbesorolás', result['changes'][0]['Ok'])
        unchanged = check_history(source(r),'same.xlsx',self.new,[old],['2026-07'],fx)
        self.assertFalse(unchanged[0]['result']['changes'])
    def test_separate_latest_baseline_per_month_and_missing(self):
        july = self.snapshot(source(rec()),['2026-07'])
        august = self.snapshot(source(rec('aug', dates='2026-08-02')),['2026-08'])
        inherit_settings(self.new,[july,august])
        raw = source(rec(),rec('aug', dates='2026-08-02'),rec('new',dates='2026-07-05'))
        results=check_history(raw,'new.xlsx',self.new,[july,august],['2026-07','2026-08','2026-09'],fx)
        self.assertEqual(results[0]['result']['changes'][0]['Fuvarszám'],'new')
        self.assertFalse(results[1]['result']['changes'])
        self.assertTrue(results[2]['missing'])
    def test_latest_archived_choice_and_local_choice_priority(self):
        r=rec(dates='2026-07-01, 2026-08-01')
        self.old.save_decision(r,'2026-07')
        first=self.snapshot(source(r),['2026-07'])
        self.old.save_decision(r,'2026-08')
        second=self.snapshot(source(r),['2026-08'])
        inherit_settings(self.new,[first,second])
        self.assertEqual(origin(r,self.new),'2026-08')
        self.new.save_decision(r,'2026-07')
        inherit_settings(self.new,[second])
        self.assertEqual(origin(r,self.new),'2026-07')
    def test_unresolved_one_month_does_not_hide_other_results(self):
        old=self.snapshot(source(rec()),['2026-07','2026-08'])
        inherit_settings(self.new,[old])
        raw=source(rec(),rec('cross', dates='2026-08-01, 2026-09-01'))
        results=check_history(raw,'new.xlsx',self.new,[old],['2026-07','2026-08'],fx)
        self.assertFalse(results[0]['result']['changes'])
        self.assertIn('error',results[1])
    def test_upload_refresh_loads_once_and_never_saves(self):
        old=self.snapshot(source(rec()),['2026-07'])
        st=MagicMock(); st.session_state={}; st.button.return_value=False
        client=Mock(folder_id='folder'); client.list_snapshots.return_value=[{'id':'saved'}]
        client.load.return_value=(old,source(rec()))
        with patch('archive_ui.archive_client',return_value=client), patch('auto_check.today',return_value=date(2026,10,21)):
            show_auto_check(st,self.new,source(rec()),'a.xlsx')
            show_auto_check(st,self.new,source(rec()),'a.xlsx')
        self.assertEqual(client.list_snapshots.call_count,1)
        client.save.assert_not_called()

    def test_real_streamlit_auto_check_and_source_replacement(self):
        from streamlit.testing.v1 import AppTest
        r=rec(currency='HUF', dates='2026-07-01, 2026-08-01')
        self.old.save_decision(r,'2026-07')
        old=self.snapshot(source(r),['2026-07','2026-08','2026-09'])
        client=Mock(folder_id='folder')
        client.list_snapshots.return_value=[{'id':'saved'}]
        client.load.return_value=(old,source(r))
        script='''
import streamlit as st
from core import Store
from auto_check import show_auto_check
show_auto_check(st, Store(st.session_state['db']), st.session_state['raw'], 'new.xlsx')
'''
        app=AppTest.from_string(script)
        app.session_state['db']=str(self.new.path)
        app.session_state['raw']=source(r)
        with patch('archive_ui.archive_client',return_value=client), patch('auto_check.today',return_value=date(2026,10,21)):
            app.run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.success),3)
            app.session_state['raw']=source(rec(currency='HUF',margin='200', dates='2026-07-01, 2026-08-01'))
            app.run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.warning),1)
            self.assertIn('Jutalékeltérés: 10 Ft', app.warning[0].value)
            self.assertEqual(origin(r,self.new),'2026-07')
            self.assertEqual(client.list_snapshots.call_count,2)
        client.save.assert_not_called()
