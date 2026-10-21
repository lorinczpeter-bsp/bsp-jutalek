"""Automatic checks against completed downloads and durable month choices."""
import hashlib
import json
from datetime import timedelta

import pandas as pd

from core import today
from drive_archive import ArchiveError, build_snapshot, compare_snapshots
from web_session import settings_backup
from persistence import restore_history


def previous_months(day, count=3):
    cursor = day.replace(day=1)
    result = []
    for _ in range(count):
        cursor = (cursor - timedelta(days=1)).replace(day=1)
        result.append(cursor.strftime('%Y-%m'))
    return sorted(result)


def inherit_settings(store, snapshots):
    # Compatibility helper: use the latest full settings, never union old removals.
    restore_history(store, snapshots, [])


def check_history(raw, name, store, snapshots, months, rate_provider=None):
    from core import load_workbook, origin
    results = []
    records = load_workbook(raw)[1]
    ordered = sorted(snapshots, key=lambda s: s['created'], reverse=True)
    for month in months:
        candidates = [s for s in ordered if month in s['months']]
        if not candidates:
            results.append({'month': month, 'missing': True}); continue
        by_contact = {}
        for snapshot in candidates:
            scope = snapshot.get('contacts')
            contacts = scope if scope is not None else {r['contact'] for r in snapshot['records'].values()}
            for contact in contacts:
                by_contact.setdefault(contact, snapshot)
        old_rows, row_dates = {}, {}
        for contact, snapshot in by_contact.items():
            for key, row in snapshot['records'].items():
                if row['contact'] == contact and snapshot['created'] > row_dates.get(key, ''):
                    old_rows[key] = row
                    row_dates[key] = snapshot['created']
        active_keys = {k for k,r in old_rows.items() if r['month'] == month}
        current_contacts = set(by_contact) | {r.contact for r in records if r.key in active_keys}
        unknown = sorted({r.contact for r in records if origin(r,store) == month} - current_contacts)
        baseline = dict(candidates[0], months=[month], records=old_rows)
        try:
            current = build_snapshot(raw, name, store, [month], rate_provider=rate_provider,
                                     contacts=sorted(current_contacts))
            current['records'] = {k:r for k,r in current['records'].items()
                                  if r['contact'] in by_contact or k in active_keys}
            result = compare_snapshots(baseline, current, [month])
            results.append({'month': month, 'baseline': baseline, 'result': result,
                            'unknown_contacts': unknown})
        except (ValueError, ArchiveError) as exc:
            results.append({'month': month, 'error': str(exc)})
    return results


def show_auto_check(st, store, raw, name):
    from archive_ui import archive_client
    from ui_features import money_display
    st.session_state['history_ready'] = False
    st.subheader('Előző 3 hónap automatikus visszaellenőrzése')
    months = previous_months(today())
    st.caption('Vizsgált naptári hónapok: ' + ', '.join(months) +
               '. Hónaponként és fuvarszervezőnként a legutóbbi kész kimutatáshoz hasonlítunk.')
    try:
        client = archive_client(st)
        if client is None:
            st.info('Az automatikus ellenőrzéshez állítsd be a Drive-archívumot.'); return
        source_key = (client.folder_id, hashlib.sha256(raw).hexdigest())
        refresh = st.button('Automatikus ellenőrzés frissítése', key='auto_check_refresh')
        if refresh or st.session_state.get('auto_history_source') != source_key:
            with st.spinner('Korábbi mentések és hónapbesorolások beolvasása…'):
                files = client.list_snapshots()
                snapshots = [client.load(f['id'])[0] for f in files]
                events = [client.load_setting_event(f) for f in client.list_setting_events()]
                # A verified write can briefly lag in Drive's list index. Keep
                # its local copy until listed so a rerun cannot undo the choice.
                local = st.session_state.setdefault('confirmed_choice_events', {})
                listed = {e['id'] for e in events}
                for key in listed:
                    local.pop(key, None)
                if local:
                    from datetime import datetime
                    events = sorted(events + list(local.values()),
                                    key=lambda e: datetime.fromisoformat(e['created']).timestamp())
                restore_history(store, snapshots, events)
            st.session_state['auto_history_snapshots'] = snapshots
            st.session_state['auto_history_source'] = source_key
            st.session_state.pop('auto_history_result', None)
        st.session_state['history_ready'] = True
        snapshots = st.session_state['auto_history_snapshots']
        signature = (source_key, tuple(months), hashlib.sha256(settings_backup(store)).hexdigest())
        cached = st.session_state.get('auto_history_result')
        if cached is None or cached[0] != signature:
            with st.spinner('Korábbi hónapok összehasonlítása…'):
                results = check_history(raw, name, store, snapshots, months)
            signature = (source_key, tuple(months), hashlib.sha256(settings_backup(store)).hexdigest())
            st.session_state['auto_history_result'] = (signature, results)
        results = st.session_state['auto_history_result'][1]
        changed = [r for r in results if r.get('result', {}).get('changes')]
        unavailable = [r for r in results if r.get('missing') or r.get('error') or r.get('unknown_contacts')]
        if changed or unavailable:
            parts = []
            if changed:
                parts.append('Eltérés: ' + ', '.join(r['month'] for r in changed))
            if unavailable:
                parts.append('Nem teljesen ellenőrizhető: ' + ', '.join(r['month'] for r in unavailable))
            st.warning(' · '.join(parts))
        else:
            st.success('Az előző három hónap mentett kész állapotaihoz képest nincs eltérés.')
        with st.expander('Visszaellenőrzés részletei', expanded=bool(changed)):
            for item in results:
                month = item['month']
                if item.get('missing'):
                    st.write(month + ': nincs korábbi kész kimutatás.')
                elif item.get('error'):
                    st.write(month + ': nem ellenőrizhető. ' + item['error'])
                else:
                    result = item['result']
                    if item.get('unknown_contacts'):
                        st.write(month + ': nincs kész mentés ezekhez a fuvarszervezőkhöz: ' + ', '.join(item['unknown_contacts']))
                    st.write(month + f": {len(result['changes'])} érintett fuvar; jutalékeltérés: " +
                             f"{result['difference']:,}".replace(',', ' ') + ' Ft')
                    if result['changes']:
                        st.dataframe(money_display(pd.DataFrame(result['changes'])), hide_index=True)
                        st.dataframe(pd.DataFrame(result['details']), hide_index=True)
                        st.dataframe(money_display(pd.DataFrame(result['monthly'])), hide_index=True)
            st.caption('A hiányzó fuvar szűkebb exportot is jelenthet. Az eltérés nem automatikus ráfizetés. '
                       'Új kész állapot az Excel vagy PDF letöltésekor mentődik.')
    except (ArchiveError, ValueError) as exc:
        st.error('Az automatikus visszaellenőrzés nem sikerült: ' + str(exc))
