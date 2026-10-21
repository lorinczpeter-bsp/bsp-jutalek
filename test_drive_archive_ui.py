import unittest
from pathlib import Path
from streamlit.testing.v1 import AppTest

class MainUITests(unittest.TestCase):
    def test_removed_tabs_and_manual_controls_absent(self):
        app=AppTest.from_file(str(Path(__file__).with_name('app.py')),default_timeout=20).run()
        self.assertFalse(app.exception)
        self.assertEqual([t.label for t in app.tabs],['Kimutatás készítése','Fuvar átsorolása','Jutalékelemzés','Használat és beállítások'])
        self.assertFalse(any('Kifizet' in b.label or 'Ellenőrzött állapot' in b.label for b in app.button))
