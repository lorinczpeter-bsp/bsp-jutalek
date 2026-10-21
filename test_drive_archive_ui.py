"""Real Streamlit rerun tests, with an in-memory Drive transport (no credentials)."""
import importlib.util
import unittest


@unittest.skipUnless(importlib.util.find_spec('streamlit'), 'Streamlit is optional for local UI tests')
class ArchiveUITests(unittest.TestCase):
    def app(self):
        from streamlit.testing.v1 import AppTest
        code = '''
import copy, tempfile
from pathlib import Path
from unittest.mock import patch
import streamlit as st
from core import Store, load_workbook
from drive_archive import build_snapshot, encode_snapshot, decode_snapshot
from test_core import rec
from test_drive_archive import source
import archive_ui
if 'test_db' not in st.session_state:
    st.session_state.test_db = str(Path(tempfile.mkdtemp()) / 'store.db')
store = Store(st.session_state.test_db)
if 'fake_archives' not in st.session_state:
    old_raw = source(rec(currency='HUF', margin='1000'))
    old = build_snapshot(old_raw, 'old.xlsx', store, ['2026-07'])
    st.session_state.fake_archives = {'first': (old, old_raw)}
class Client:
    folder_id = 'fake-folder'
    def check_folder(self):return {'driveId': 'shared'}
    def list_snapshots(self):
        return [{'id': k, 'name': k + '.zip'} for k in reversed(st.session_state.fake_archives)]
    def load(self, key):return copy.deepcopy(st.session_state.fake_archives[key])
    def save(self, snapshot, raw):
        saved, saved_raw = decode_snapshot(encode_snapshot(snapshot, raw))
        st.session_state.fake_archives[snapshot['id']] = (saved, saved_raw)
        return {'id': snapshot['id'], 'name': 'saved.zip'}
if 'current_raw' not in st.session_state:
    st.session_state.current_raw = source(rec(currency='HUF', margin='1500'))
raw = st.session_state.current_raw
headers, records, errors = load_workbook(raw)
with patch('archive_ui.archive_client', lambda st: Client()):
    archive_ui.show_archive(st, store, ('hash', headers, records, errors), raw, 'current.xlsx')
'''
        at = AppTest.from_string(code, default_timeout=20).run()
        at.checkbox[0].check().run()
        return at

    def test_comparison_confirmation_save_and_fresh_baseline(self):
        at = self.app()
        self.assertFalse(at.exception)
        self.assertIn('50 Ft', [m.value for m in at.metric])
        complete = next(c for c in at.checkbox if c.label.startswith('Ellenőriztem:'))
        save = next(b for b in at.button if b.label == 'Ellenőrzött állapot mentése')
        self.assertTrue(save.disabled)
        complete.check().run()
        next(b for b in at.button if b.label == 'Ellenőrzött állapot mentése').click().run()
        self.assertFalse(at.exception)
        self.assertEqual(len(at.session_state['fake_archives']), 2)
        self.assertIn('0 Ft', [m.value for m in at.metric])
        self.assertTrue(any('nincs eltérés' in s.value for s in at.success))

    def test_concurrent_new_archive_requires_review_before_save(self):
        at = self.app()
        next(c for c in at.checkbox if c.label.startswith('Ellenőriztem:')).check().run()
        entries = at.session_state['fake_archives']
        entries['another-user'] = entries['first']
        next(b for b in at.button if b.label == 'Ellenőrzött állapot mentése').click().run()
        self.assertFalse(at.exception)
        self.assertEqual(len(at.session_state['fake_archives']), 2)
        self.assertTrue(any('közben megváltozott' in w.value for w in at.warning))

    def test_main_app_without_drive_still_loads(self):
        from pathlib import Path
        from streamlit.testing.v1 import AppTest
        at = AppTest.from_file(str(Path(__file__).with_name('app.py')), default_timeout=20).run()
        self.assertFalse(at.exception)
        self.assertIn('Drive-archívum', [tab.label for tab in at.tabs])
        self.assertTrue(any('még nincs összekapcsolva' in x.value for x in at.info))


if __name__ == '__main__':unittest.main()
