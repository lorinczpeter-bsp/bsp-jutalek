"""Streamlit archive workflow, kept separate from the normal monthly report."""
import copy
import hashlib
import json

import pandas as pd

from core import today
from drive_archive import ArchiveError, DriveArchive, build_snapshot, compare_snapshots
from exports import ft
from ui_features import money_display
from web_session import restore_settings, settings_backup


def archive_client(st):
    try:
        config = dict(st.secrets.get('drive_archive', {}))
        credentials = dict(st.secrets.get('gcp_service_account', {}))
    except Exception:
        config, credentials = {}, {}
    if not config.get('folder_id') or not credentials:
        return None
    return DriveArchive(config, credentials)


def show_archive(st, store, bundle, raw=None, source_name='fuvarlista.xlsx'):
    st.subheader('Drive-archívum és korábbi hónapok változásai')
    st.caption('Az archívum összehasonlítási alapot őriz. A mentés nem havi lezárás és nem jelent kifizetést.')
    try:
        client = archive_client(st)
    except ArchiveError as exc:
        st.error(str(exc)); return
    if client is None:
        st.info('A Drive-archívum még nincs összekapcsolva. A DRIVE_ARCHIVUM_BEALLITAS.md útmutató alapján '
                'állítsd be a céges megosztott meghajtó mappáját és a szolgáltatásfiókot a Streamlit Secretsben.')
        return
    if not st.checkbox('Drive-archívum megnyitása', key='archive_open'):
        return
    context = client.folder_id
    if st.session_state.get('archive_folder') != context:
        for key in ('archive_files', 'archive_loaded', 'archive_preview'):
            st.session_state.pop(key, None)
        st.session_state['archive_folder'] = context
    try:
        if st.button('Archívumlista frissítése', key='archive_refresh') or 'archive_files' not in st.session_state:
            client.check_folder()
            st.session_state['archive_files'] = client.list_snapshots()
            st.session_state.pop('archive_loaded', None)
        files = st.session_state['archive_files']
        baseline = None
        if files:
            names = {f['id']: f['name'] for f in files}
            selected = st.selectbox('Összehasonlítás alapja', list(names), format_func=names.get, key='archive_baseline')
            cache = st.session_state.get('archive_loaded')
            if not cache or cache[0] != selected:
                with st.spinner('Korábbi ellenőrzött állapot beolvasása…'):
                    saved, old_raw = client.load(selected)
                st.session_state['archive_loaded'] = (selected, saved, old_raw)
            _, baseline, old_raw = st.session_state['archive_loaded']
            st.caption('Mentve: ' + baseline['created'][:19].replace('T', ' ') +
                       ' | Ellenőrzött hónapok: ' + ', '.join(baseline['months']) +
                       ' | Forrás: ' + baseline['source_name'])
            if baseline.get('note'):
                st.write(baseline['note'])
            with st.expander('Korábbi forrás és beállítások elérése'):
                st.download_button('Archivált forrás-Excel letöltése', old_raw,
                                   file_name=baseline['source_name'], mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                   key='archive_source_download')
                st.download_button('Archivált beállítások letöltése (JSON)',
                                   json.dumps(baseline['settings'], ensure_ascii=False, indent=2).encode(),
                                   file_name='jutalek_archivalt_beallitasok.json', mime='application/json', key='archive_settings_download')
                st.caption('Új munkamenetben az átsorolások és pontos árfolyamok átvételéhez töltsd vissza '
                           'az archivált beállításokat. Kifizetési előzményt ez nem állít vissza.')
                accepted = st.checkbox('Átveszem az archivált dátumdöntéseket, átsorolásokat és árfolyamokat. '
                                       'Az azonos kulcsú helyi beállításokat felülírja.', key='archive_restore_confirm')
                if st.button('Archivált beállítások átvétele', disabled=not accepted, key='archive_restore'):
                    restore_settings(store, json.dumps(baseline['settings']).encode())
                    st.session_state['decision_feedback'] = 'Az archivált beállítások átvétele sikerült. Kifizetési jelölés nem változott.'
                    st.rerun()
        else:
            st.info('Még nincs mentett ellenőrzött állapot. Az első mentés hozza létre az összehasonlítás alapját. '
                    'Korábbi eredeti fuvarlistát is feltölthetsz a Kimutatás készítése fülön, és elmentheted első alapként.')
        if not bundle or raw is None:
            st.info('Az aktuális összehasonlításhoz először töltsd fel a teljes fuvarlistát a Kimutatás készítése fülön.')
            return
        _, _, records, errors = bundle
        if errors:
            st.error('A forrás hibás sorokat tartalmaz. Javítsd a forrást; addig ellenőrzött állapot nem menthető.')
            return
        available = sorted({m for r in records for m in r.months} |
                           {store.override(r) for r in records if store.override(r)} |
                           set(baseline['months'] if baseline else []))
        defaults = baseline['months'] if baseline else [m for m in available if m < today().strftime('%Y-%m')]
        scope_key = 'archive_months_' + (baseline['id'] if baseline else 'first')
        months = st.multiselect('Ellenőrzött hónapok', available, default=defaults, key=scope_key)
        st.caption('A mentés a kiválasztott hónapok minden fuvarszervezőjére vonatkozik, '
                   'nem csak a fő fülön kiválasztott személyre. A teljes eredeti Excel is megmarad.')
        if not months:
            st.info('Válassz legalább egy ellenőrzött hónapot.'); return
        note = st.text_input('Mentés megjegyzése (nem kötelező)', key='archive_note')
        # Calculation can populate the exact-rate cache, so retain the post-calculation key.
        def input_key():
            return hashlib.sha256(raw + settings_backup(store) +
                                  json.dumps([months, note], ensure_ascii=False).encode()).hexdigest()
        signature = input_key()
        cached = st.session_state.get('archive_preview')
        if not cached or cached[0] != signature:
            with st.spinner('Korábbi hónapok ellenőrzése és összehasonlítása…'):
                current = build_snapshot(raw, source_name, store, months, note)
            st.session_state['archive_preview'] = (input_key(), current)
        current = st.session_state['archive_preview'][1]
        if baseline:
            common = sorted(set(baseline['months']) & set(months))
            excluded = sorted(set(baseline['months']) - set(months))
            if excluded:
                st.warning('Most nem vizsgált korábbi hónapok: ' + ', '.join(excluded))
            extra = sorted(set(months) - set(baseline['months']))
            if extra:
                st.info('Ezekhez a hónapokhoz ebben az alapállapotban nincs ellenőrzött viszonyítás: ' + ', '.join(extra))
            if common:
                result = compare_snapshots(baseline, current, common)
                st.subheader('Korábbi hónapok változásai')
                a, b = st.columns(2)
                a.metric('Érintett fuvarok', len(result['changes']))
                b.metric('Számított jutalék eltérése', ft(result['difference']) + ' Ft')
                st.caption('Vizsgált hónapok: ' + ', '.join(common) + '. A teljes számított jutalékot hasonlítjuk, '
                           'kifizetési levonás nélkül. A különbözet nem automatikus kifizetési javaslat.')
                if result['changes']:
                    if any(r['Változás'] == 'Hiányzik az új forrásból' for r in result['changes']):
                        st.warning('Hiányzó fuvar is van. Ellenőrizd az export időszakát és szűrőit; ez nem bizonyít törlést.')
                    st.dataframe(money_display(pd.DataFrame(result['changes'])), hide_index=True)
                    with st.expander('Módosult mezők és pontos értékek'):
                        st.dataframe(pd.DataFrame(result['details']), hide_index=True)
                    st.subheader('Havi eltérés fuvarszervezőnként')
                    st.dataframe(money_display(pd.DataFrame(result['monthly'])), hide_index=True)
                else:
                    st.success('A vizsgált hónapokban nincs eltérés a mentett állapothoz képest.')
            else:
                st.warning('Nincs közös ellenőrzött hónap, ezért eltérésmentesség nem igazolható.')
        st.subheader('Ellenőrzött állapot mentése')
        complete = st.checkbox('Ellenőriztem: a kijelölt hónapok teljes fuvarlistája, minden érintett fuvarszervezővel '
                               'és a helyes hónapbesorolásokkal szerepel a forrásban. A változásokat áttekintettem.',
                               key='archive_complete_' + current['id'])
        st.caption('Külön, dátumozott állapotot mentünk. Korábbi mentést nem írunk felül. '
                   'Új feltöltés nem válik magától ellenőrzött összehasonlítási alappá.')
        if st.button('Ellenőrzött állapot mentése', disabled=not complete, type='primary', key='archive_save'):
            fresh_files = client.list_snapshots()
            if [f['id'] for f in fresh_files] != [f['id'] for f in files]:
                st.warning('Az archívumlista közben megváltozott. Frissítsd a listát, ellenőrizd az alapállapotot, majd ments újra.')
                return
            with st.spinner('Állapot mentése és Drive-visszaellenőrzése…'):
                saved = client.save(copy.deepcopy(current), raw)
            st.session_state['decision_feedback'] = 'Ellenőrzött állapot mentve és visszaellenőrizve: ' + saved['name']
            st.session_state.pop('archive_files', None)
            st.session_state.pop('archive_loaded', None)
            st.session_state.pop('archive_baseline', None)
            st.rerun()
    except (ArchiveError, ValueError) as exc:
        st.error(str(exc))
