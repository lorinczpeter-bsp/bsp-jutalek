import json,os,tempfile,unittest
from pathlib import Path
from unittest.mock import patch
import pandas as pd
from core import Store,origin
from settlement import reassign,import_payments,payments
from test_core import rec
from ui_features import money_display,commission_chart_spec,show_analytics
from test_features_ui import FeatureApp
from web_session import session_store,settings_backup,restore_settings

class WebTests(unittest.TestCase):
    def test_hungarian_money_format_does_not_change_numeric_source(self):
        source=pd.DataFrame({'Hónap':['2026-09'],'Jutalék (Ft)':[4799741],'Havi átlag (Ft)':[479974.1]})
        view=money_display(source)
        self.assertEqual(view.iloc[0]['Jutalék (Ft)'],'4 799 741 Ft')
        self.assertEqual(view.iloc[0]['Havi átlag (Ft)'],'479 974,10 Ft')
        self.assertEqual(source.iloc[0]['Jutalék (Ft)'],4799741)
        self.assertIn('Ft',commission_chart_spec('bar')['encoding']['y']['axis']['labelExpr'])
        self.assertEqual(commission_chart_spec('line')['encoding']['tooltip'][-1]['field'],'Jutalék')
    def test_default_cloud_sessions_do_not_share_state(self):
        with tempfile.TemporaryDirectory() as root,patch.dict(os.environ,{},clear=True):
            states=[{},{}]
            with patch('web_session.tempfile.mkdtemp',side_effect=[str(Path(root)/'a'),str(Path(root)/'b')]):
                first=session_store(states[0]);again=session_store(states[0]);second=session_store(states[1])
            self.assertEqual(first.path,again.path);self.assertNotEqual(first.path,second.path)
            reassign(first,rec(),'2026-08')
            self.assertEqual(origin(rec(),first),'2026-08')
            self.assertEqual(origin(rec(),second),'2026-07')
    def test_portable_settings_omit_payment_state(self):
        with tempfile.TemporaryDirectory() as root:
            a=Store(Path(root)/'a.sqlite3');b=Store(Path(root)/'b.sqlite3')
            reassign(a,rec(),'2026-08')
            a.save_rate('2026-07',{'rate':'365.0531818181818181818','source':'teszt'})
            import_payments(a,[{'freight':'1-1','contact':'Peti','amount':'100'}],'paid.xlsx','hash','2026-08-01')
            payload=settings_backup(a)
            self.assertNotIn('payments',json.loads(payload))
            restore_settings(b,payload)
            self.assertEqual(origin(rec(),b),'2026-08')
            self.assertEqual(b.rate('2026-07')['rate'],'365.0531818181818181818')
            self.assertFalse(payments(b))
            broken=json.loads(payload);broken['overrides'][0]['month']='2026-99'
            with self.assertRaises(ValueError):restore_settings(b,json.dumps(broken).encode())
            self.assertEqual(origin(rec(),b),'2026-08')
    def test_analytics_never_shows_payment_columns(self):
        with tempfile.TemporaryDirectory() as root:
            store=Store(Path(root)/'a.sqlite3')
            record=rec('1-1',currency='HUF',margin='1000')
            import_payments(store,[{'freight':'1-1','contact':'Peti','amount':'100'}],'paid.xlsx','hash','2026-08-01')
            ui=FeatureApp();show_analytics(ui,store,('x',list(record.original),[record],[]))
            self.assertIn(('Kiválasztott hónapok összege','100 Ft'),ui.metrics)
            self.assertNotIn('Még fizetendő',[x[0] for x in ui.metrics])
            for table in ui.tables:
                self.assertFalse(any('fizet' in c.lower() for c in table.columns))
            self.assertEqual(ui.tables[0]['Számított jutalék (Ft)'].iloc[0],'100 Ft')

if __name__=='__main__':unittest.main()
