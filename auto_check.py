"""Read-only automatic history checks; checked archives remain explicit saves."""
import hashlib
import json
from datetime import timedelta

import pandas as pd

from core import today, valid_month, dec
from drive_archive import ArchiveError, build_snapshot, compare_snapshots
from web_session import settings_backup


def previous_months(day, count=3):
    cursor = day.replace(day=1)
    result = []
    for _ in range(count):
        cursor = (cursor - timedelta(days=1)).replace(day=1)
        result.append(cursor.strftime('%Y-%m'))
    return sorted(result)


def inherit_settings(store, snapshots):
    """Newest archived key wins; never replace choices made in this session."""
    merged = {table: {} for table in ('decisions', 'overrides', 'rates')}
    for snapshot in sorted(snapshots, key=lambda s: s['created'], reverse=True):
        for table in merged:
            for row in snapshot['settings'].get(table, []):
                merged[table].setdefault(row['month'] if table == 'rates' else row['freight'], row)
    count = 0
    with store.connect() as connection:
        for table, rows in merged.items():
            for row in rows.values():
                columns = list(row)
                sql = ('INSERT OR IGNORE INTO ' + table + ' (' + ','.join(columns) +
                       ') VALUES (' + ','.join('?' for _ in columns) + ')')
                # Validated archive settings only; fixed column names from the snapshot schema.
                allowed = {'decisions': {'freight','signature','month','note'},
                           'overrides': {'freight','month','note','created'},
                           'rates': {'month','payload'}}
                if (set(columns) != allowed[table] or
                    not all(isinstance(v, str) for v in row.values()) or
                    not valid_month(row['month'])):
                    raise ArchiveError('Hibás archivált beállítás.')
                if table == 'rates':
                    try:
                        if dec(json.loads(row['payload'])['rate']) <= 0:
                            raise ValueError()
                    except (ValueError, KeyError, TypeError):
                        raise ArchiveError('Hibás archivált árfolyam.')
                count += connection.execute(sql, tuple(row[c] for c in columns)).rowcount
        if count:
            store.bump(connection)
    return count


def check_history(raw, name, store, snapshots, months, rate_provider=None):
    results = []
    ordered = sorted(snapshots, key=lambda s: s['created'], reverse=True)
    for month in months:
        baseline = next((s for s in ordered if month in s['months']), None)
        if baseline is None:
            results.append({'month': month, 'missing': True})
            continue
        try:
            current = build_snapshot(raw, name, store, [month], rate_provider=rate_provider)
            result = compare_snapshots(baseline, current, [month])
            results.append({'month': month, 'baseline': baseline, 'result': result})
        except (ValueError, ArchiveError) as exc:
            results.append({'month': month, 'error': str(exc)})
    return results


def show_auto_check(st, store, raw, name):
    from archive_ui import archive_client
    from ui_features import money_display
    st.subheader('Előző 3 hónap automatikus visszaellenőrzése')
    months = previous_months(today())
    st.caption('Vizsgált naptári hónapok: ' + ', '.join(months) +
               '. Hónaponként a legutóbbi ellenőrzött mentést használjuk, minden fuvarszervezőre.')
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
                inherit_settings(store, snapshots)
            st.session_state['auto_history_snapshots'] = snapshots
            st.session_state['auto_history_source'] = source_key
            st.session_state.pop('auto_history_result', None)
        snapshots = st.session_state['auto_history_snapshots']
        signature = (source_key, tuple(months), hashlib.sha256(settings_backup(store)).hexdigest())
        cached = st.session_state.get('auto_history_result')
        if cached is None or cached[0] != signature:
            with st.spinner('Korábbi hónapok összehasonlítása…'):
                results = check_history(raw, name, store, snapshots, months)
            signature = (source_key, tuple(months), hashlib.sha256(settings_backup(store)).hexdigest())
            st.session_state['auto_history_result'] = (signature, results)
        for item in st.session_state['auto_history_result'][1]:
            month = item['month']
            if item.get('missing'):
                st.info(month + ': nincs ellenőrzött mentés; nem ellenőrizhető.')
            elif item.get('error'):
                st.warning(month + ': nem ellenőrizhető. ' + item['error'])
            else:
                result = item['result']
                if not result['changes']:
                    st.success(month + ': nincs eltérés.')
                else:
                    st.warning(month + f": {len(result['changes'])} érintett fuvar. Jutalékeltérés: " +
                               f"{result['difference']:,}".replace(',', ' ') + ' Ft')
                    with st.expander(month + ' – eltérések részletei'):
                        st.caption('Összehasonlítás alapja: ' + item['baseline']['created'][:19].replace('T', ' '))
                        st.dataframe(money_display(pd.DataFrame(result['changes'])), hide_index=True)
                        st.dataframe(pd.DataFrame(result['details']), hide_index=True)
                        st.dataframe(money_display(pd.DataFrame(result['monthly'])), hide_index=True)
                        st.caption('A hiányzó fuvar szűkebb exportot is jelenthet. A különbözet nem kifizetési javaslat.')
        st.caption('A korábban mentett hónapbesorolásokat automatikusan átvettük. '
                   'Új ellenőrzött alapot a Drive-archívum fülön menthetsz. '
                   'Új munkamenetben csak az archívumba elmentett döntések állíthatók vissza.')
    except (ArchiveError, ValueError) as exc:
        st.error('Az automatikus visszaellenőrzés nem sikerült: ' + str(exc))
