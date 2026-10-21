import copy
import hashlib
import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
from datetime import date

import openpyxl

from core import Store, origin, build_plan, load_workbook
from drive_archive import ArchiveError, build_snapshot, encode_snapshot, decode_snapshot
from persistence import ExportJob, restore_history, apply_event, persist_choice
from auto_check import check_history
from settlement import reassign, clear_reassignment
from test_core import rec, fx, HEADERS
from test_drive_archive import source
from web_session import settings_backup


def event(action='override', month='2026-08', key='1-1', eid='a'):
    return {'format':'babolna-month-choice-v1','id':eid*32,'created':'2026-10-21T12:00:00+02:00',
            'freight':key,'signature':'signature','action':action,'month':month,'note':''}


class PersistenceTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.store=Store(Path(self.tmp.name)/'store.db')
        self.client=Mock()
    def tearDown(self):self.tmp.cleanup()
    def snapshot(self, records, contacts=None):
        return build_snapshot(source(*records),'source.xlsx',self.store,['2026-07'],rate_provider=fx,contacts=contacts)
    def job(self, records=None):
        records=records or [rec(currency='HUF')]
        raw=source(*records)
        headers, loaded, errors=load_workbook(raw)
        p=build_plan(loaded,headers,self.store,'2026-07','Peti',hashlib.sha256(raw).hexdigest(),snapshot=True,rate_provider=fx)
        return ExportJob(self.client,raw,'source.xlsx',settings_backup(self.store),p)
    def test_no_upload_on_render_and_save_before_bytes_for_both_formats(self):
        job=self.job()
        self.client.save.assert_not_called()
        x=job.download('xlsx');p=job.download('pdf')
        self.assertTrue(x.startswith(b'PK'));self.assertTrue(p.startswith(b'%PDF'))
        self.assertEqual(self.client.save.call_count,1)
        snapshot,raw,reports=self.client.save.call_args.args
        restored,_=decode_snapshot(encode_snapshot(snapshot,raw,reports))
        self.assertEqual(restored,snapshot)
        self.assertEqual(snapshot['contacts'],['Peti'])
        self.assertEqual(reports['report.xlsx'],x)
        self.assertEqual(reports['report.pdf'],p)
    def test_drive_failure_blocks_download_and_retry_keeps_same_id(self):
        self.client.save.side_effect=[ArchiveError('Drive átmeneti hiba'),{'id':'saved'}]
        job=self.job()
        with self.assertRaisesRegex(RuntimeError,'nem indult el'):
            job.download('xlsx')
        first=job.snapshot['id'];self.assertIsNone(job.saved)
        self.assertTrue(job.error)
        self.assertTrue(job.download('pdf').startswith(b'%PDF'))
        self.assertEqual(first,job.snapshot['id'])
        self.assertEqual(self.client.save.call_count,2)
    def test_incomplete_report_cannot_be_promoted_to_ready(self):
        job=self.job();job.plan['blockers']=['Hiányzó árfolyam']
        with self.assertRaises(RuntimeError):job.download('xlsx')
        self.client.save.assert_not_called()
    def test_changed_settings_do_not_mutate_frozen_download(self):
        job=self.job()
        reassign(self.store,rec(),'2026-08')
        job.download('xlsx')
        self.assertEqual(job.snapshot['records']['1-1']['month'],'2026-07')
    def test_selected_person_only_and_other_person_pending_does_not_block_export(self):
        r=rec('other',owner='Anna',dates='2026-07-01, 2026-08-01')
        job=self.job([rec(currency='HUF'),r])
        job.download('pdf')
        snapshot=job.snapshot
        self.assertIsNone(snapshot['records']['other']['calculation'])
        decode_snapshot(encode_snapshot(snapshot,job.raw,job.reports))
    def test_export_hash_tampering_is_rejected(self):
        job=self.job();job.download('xlsx')
        with self.assertRaises(ArchiveError):
            decode_snapshot(encode_snapshot(job.snapshot,job.raw,{'report.xlsx':b'bad','report.pdf':b'bad'}))
    def test_removed_override_is_not_resurrected_from_old_snapshot(self):
        r=rec();reassign(self.store,r,'2026-08')
        a=self.snapshot([r]);clear_reassignment(self.store,r.key);b=self.snapshot([r])
        restore_history(self.store,[a,b],[])
        self.assertEqual(origin(r,self.store),'2026-07')
    def test_event_clear_survives_restart_and_old_stale_snapshot(self):
        r=rec();reassign(self.store,r,'2026-08');old=self.snapshot([r])
        restore_history(self.store,[old],[event(),event('clear_override',None,eid='b')])
        self.assertEqual(origin(r,self.store),'2026-07')
        restore_history(self.store,[old],[event(),event('clear_override',None,eid='b')])
        self.assertEqual(origin(r,self.store),'2026-07')
    def test_decision_event_survives_restart_and_amount_change(self):
        r=rec(dates='2026-07-01, 2026-08-01')
        restore_history(self.store,[],[event('decision','2026-07')])
        self.assertEqual(origin(r,self.store),'2026-07')
        self.assertEqual(origin(rec(margin='200',dates='2026-07-01, 2026-08-01'),self.store),'2026-07')
    def test_choice_write_failure_does_not_claim_local_success_and_retry_reuses_uuid(self):
        st=Mock();st.session_state={}
        self.client.save_setting_event.side_effect=[ArchiveError('Drive hiba'),{'id':'saved'}]
        r=rec(dates='2026-07-01, 2026-08-01')
        with patch('archive_ui.archive_client',return_value=self.client):
            with self.assertRaises(ArchiveError):persist_choice(st,self.store,r,'decision','2026-07')
            self.assertIsNone(origin(r,self.store))
            first=st.session_state['pending_choice'][1]['id']
            persist_choice(st,self.store,r,'decision','2026-07')
        self.assertEqual(origin(r,self.store),'2026-07')
        self.assertEqual(first,self.client.save_setting_event.call_args.args[0]['id'])
    def test_invalid_history_does_not_erase_active_decisions(self):
        self.store.save_decision(rec(dates='2026-07-01, 2026-08-01'),'2026-07')
        with self.assertRaises(ArchiveError):restore_history(self.store,[],[{}])
        self.assertEqual(origin(rec(),self.store),'2026-07')
    def test_latest_ready_baseline_per_person_not_whole_month(self):
        records=[rec(currency='HUF'),rec('anna',currency='HUF',owner='Anna')]
        a=self.snapshot(records,['Peti'])
        records[1]=rec('anna',currency='HUF',margin='200',owner='Anna')
        b=self.snapshot(records,['Anna'])
        records[0]=rec(currency='HUF',margin='150')
        result=check_history(source(*records),'new.xlsx',self.store,[a,b],['2026-07'],fx)[0]['result']
        self.assertEqual(len(result['changes']),1)
        self.assertEqual(result['changes'][0]['Fuvarszám'],'1-1')
        self.assertEqual(result['difference'],5)
    def test_new_person_without_ready_report_is_unknown_not_no_change(self):
        a=self.snapshot([rec(currency='HUF')],['Peti'])
        result=check_history(source(rec(currency='HUF'),rec('new',currency='HUF',owner='Anna')),
                             'new.xlsx',self.store,[a],['2026-07'],fx)[0]
        self.assertEqual(result['unknown_contacts'],['Anna'])

    def test_live_streamlit_upload_registers_deferred_downloads_without_saving(self):
        from streamlit.testing.v1 import AppTest
        raw=source(rec(currency='HUF'))
        client=Mock(folder_id='folder')
        client.list_snapshots.return_value=[]
        client.list_setting_events.return_value=[]
        client.save.return_value={'id':'saved','name':'saved.zip'}
        script='''
import io, runpy
from unittest.mock import patch
import streamlit as st
uploaded=io.BytesIO(st.session_state['test_raw'])
uploaded.name='source.xlsx'
with patch('streamlit.file_uploader',return_value=uploaded), patch('archive_ui.archive_client',return_value=st.session_state['test_client']):
    runpy.run_path(st.session_state['test_app'],run_name='__main__')
'''
        app=AppTest.from_string(script,default_timeout=20)
        app.session_state['test_raw']=raw
        app.session_state['test_client']=client
        app.session_state['test_app']=str(Path(__file__).with_name('app.py'))
        app.run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.session_state['export_jobs']),1)
        client.save.assert_not_called()
        job=next(iter(app.session_state['export_jobs'].values()))
        self.assertTrue(job.download('xlsx').startswith(b'PK'))
        client.save.assert_called_once()

    def test_setting_event_transport_create_readback_and_idempotence(self):
        from drive_archive import DriveArchive
        client=DriveArchive.__new__(DriveArchive);client.folder_id='folder'
        e=event();client.check_folder=Mock()
        client.list_setting_events=Mock(side_effect=[[],[{'id':'saved','appProperties':{'event_id':e['id']}}]])
        response=Mock();response.json.return_value={'id':'saved'}
        client._request=Mock(return_value=response)
        client.load_setting_event=Mock(return_value=e)
        client.save_setting_event(e);client.save_setting_event(e)
        self.assertEqual(client._request.call_count,1)
        self.assertEqual(client._request.call_args.args[0],'POST')
        self.assertEqual(client.load_setting_event.call_count,2)
    def test_setting_event_uncertain_post_recovers_without_reupload(self):
        from drive_archive import DriveArchive, DriveRequestError
        client=DriveArchive.__new__(DriveArchive);client.folder_id='folder'
        e=event();client.check_folder=Mock()
        client.list_setting_events=Mock(side_effect=[[],[{'id':'saved','appProperties':{'event_id':e['id']}}]])
        client._request=Mock(side_effect=DriveRequestError('temporary',status=500,transient=True))
        client.load_setting_event=Mock(return_value=e)
        self.assertEqual(client.save_setting_event(e)['id'],'saved')
        self.assertEqual(client._request.call_count,1)

    def test_concurrent_format_requests_share_single_ready_backup(self):
        from concurrent.futures import ThreadPoolExecutor
        job=self.job()
        with ThreadPoolExecutor(max_workers=2) as pool:
            x=pool.submit(job.download,'xlsx');p=pool.submit(job.download,'pdf')
            self.assertTrue(x.result().startswith(b'PK'))
            self.assertTrue(p.result().startswith(b'%PDF'))
        self.client.save.assert_called_once()
    def test_confirmed_choice_not_lost_if_drive_listing_lags(self):
        from auto_check import show_auto_check
        from unittest.mock import MagicMock
        st=MagicMock();st.session_state={'confirmed_choice_events':{'a'*32:event('decision','2026-07')}}
        client=Mock(folder_id='folder');client.list_snapshots.return_value=[];client.list_setting_events.return_value=[]
        with patch('archive_ui.archive_client',return_value=client):
            show_auto_check(st,self.store,source(rec(dates='2026-07-01, 2026-08-01')),'new.xlsx')
        self.assertEqual(origin(rec(dates='2026-07-01, 2026-08-01'),self.store),'2026-07')
        self.assertTrue(st.session_state['history_ready'])
