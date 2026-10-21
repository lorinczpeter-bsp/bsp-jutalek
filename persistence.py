"""Automatic choice persistence and export-before-download backups."""
import copy
import hashlib
import json
import tempfile
import threading
import uuid
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from core import Store
from drive_archive import ArchiveError, build_snapshot, validate_setting_event
from exports import make_excel, make_pdf
from web_session import restore_settings, settings_backup


def apply_event_connection(c, event):
    validate_setting_event(event)
    action, key = event['action'], event['freight']
    if action == 'decision':
        c.execute('INSERT OR REPLACE INTO decisions VALUES(?,?,?,?)',
                  (key, event['signature'], event['month'], event['note']))
    elif action == 'override':
        c.execute('INSERT OR REPLACE INTO overrides VALUES(?,?,?,?)',
                  (key, event['month'], event['note'], event['created']))
    else:
        c.execute('DELETE FROM overrides WHERE freight=?', (key,))


def apply_event(store, event):
    with store.connect() as c:
        apply_event_connection(c, event)
        store.bump(c)


def restore_history(store, snapshots, events):
    # Validate and replay in isolation before changing the active session.
    latest = max(snapshots, key=lambda s: s['created'], default=None)
    with tempfile.TemporaryDirectory(prefix='jutalek_restore_') as tmp:
        staging = Store(Path(tmp)/'settings.db')
        if latest:
            restore_settings(staging, json.dumps(latest['settings']).encode())
        with staging.connect() as connection:
            for event in events:
                apply_event_connection(connection, event)
            staging.bump(connection)
        payload = json.loads(settings_backup(staging))
    with store.connect() as c:
        for table in ('decisions','overrides'):
            c.execute('DELETE FROM '+table)
        for table in ('decisions','overrides','rates'):
            for row in payload[table]:
                columns = list(row)
                c.execute('INSERT OR REPLACE INTO '+table+' ('+','.join(columns)+') VALUES ('+
                          ','.join('?' for _ in columns)+')',tuple(row[k] for k in columns))
        store.bump(c)


def persist_choice(st, store, record, action, month=None, note=''):
    from archive_ui import archive_client
    if st.session_state.get('history_ready') is False:
        raise ArchiveError('Előbb frissítsd az automatikus ellenőrzést, hogy a korábbi hónapdöntések beolvasása sikerüljön.')
    client = archive_client(st)
    if client is None:
        raise ArchiveError('A hónapdöntéshez működő Drive-kapcsolat szükséges.')
    if action == 'decision' and month not in record.months:
        raise ArchiveError('A felsorolt hónapok egyikét válaszd ki.')
    identity = (action, record.key, month, note)
    pending = st.session_state.get('pending_choice')
    if not pending or pending[0] != identity:
        event = {'format': 'babolna-month-choice-v1', 'id': uuid.uuid4().hex,
                 'created': datetime.now(ZoneInfo('Europe/Budapest')).isoformat(),
                 'freight': record.key, 'signature': record.signature,
                 'action': action, 'month': month, 'note': note.strip()[:10000]}
        st.session_state['pending_choice'] = (identity, event)
    event = st.session_state['pending_choice'][1]
    client.save_setting_event(event)  # No local success before verified cloud write.
    apply_event(store, event)
    st.session_state.setdefault('confirmed_choice_events', {})[event['id']] = event
    st.session_state.pop('pending_choice', None)
    st.session_state.pop('auto_history_source', None)


class ExportJob:
    """Frozen download request. Streamlit calls this on its download thread.

    Never calls Streamlit in that thread; the UI polls the small status object.
    Failed saves retain their UUID for safe retries. No bytes leave before save.
    """
    def __init__(self, client, raw, name, settings, plan):
        self.client, self.raw, self.name = client, raw, name
        self.settings, self.plan = settings, copy.deepcopy(plan)
        self.lock = threading.Lock()
        self.snapshot = None
        self.reports = None
        self.saved = None
        self.status = ''
        self.error = ''
        self.finished = False
        self.completed_count = 0

    def download(self, kind):
        with self.lock:
            self.error = ''; self.finished = False
            try:
                if not self.saved:
                    self.status = 'Kimutatás készítése és automatikus Drive-mentés…'
                    if self.reports is None:
                        if self.plan['blockers']:
                            raise ArchiveError('A kimutatás hibás vagy hiányos; előbb rendezd a jelzett tételeket.')
                        self.reports = {'report.xlsx': make_excel(self.plan), 'report.pdf': make_pdf(self.plan)}
                    if self.snapshot is None:
                        with tempfile.TemporaryDirectory(prefix='jutalek_export_') as tmp:
                            store = Store(Path(tmp)/'frozen.db')
                            restore_settings(store, self.settings)
                            snapshot = build_snapshot(self.raw, self.name, store, [self.plan['month']],
                                                      contacts=[self.plan['contact']])
                        total = sum(r['calculation']['commission'] for r in snapshot['records'].values()
                                    if r['calculation'] is not None)
                        if total != self.plan['total'] or snapshot['source_sha256'] != self.plan['source_hash']:
                            raise ArchiveError('A letöltés és a mentés számítása eltér. Frissítsd a kimutatást.')
                        snapshot['reports'] = {k: hashlib.sha256(v).hexdigest() for k,v in self.reports.items()}
                        snapshot['export_month'] = self.plan['month']
                        snapshot['export_contact'] = self.plan['contact']
                        self.snapshot = snapshot
                    self.saved = self.client.save(self.snapshot, self.raw, self.reports)
                self.status = 'A kész kimutatás Drive-mentése sikeres és visszaellenőrzött.'
                self.finished = True
                return self.reports['report.' + kind]
            except Exception as exc:
                self.error = str(exc) if isinstance(exc, ValueError) else 'A mentés nem sikerült. Próbáld újra a letöltést.'
                self.status = ''; self.finished = True
                raise RuntimeError('A letöltés nem indult el: ' + self.error) from exc
            finally:
                self.completed_count += 1
