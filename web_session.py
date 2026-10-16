"""Session-isolated state for Community Cloud and portable reassignment settings."""
import json,os,tempfile
from pathlib import Path
from core import Store,valid_month,dec

def session_store(state):
    explicit=os.environ.get('JUTALEK_DB')
    if explicit:return Store(explicit)
    if not state.get('session_db_path'):
        state['session_db_path']=str(Path(tempfile.mkdtemp(prefix='babolna_jutalek_'))/'session.sqlite3')
    return Store(state['session_db_path'])

def settings_backup(store):
    data={'format':'babolna-jutalek-settings-v1','decisions':[],'overrides':[],'rates':[]}
    with store.connect() as c:
        for table in ('decisions','overrides','rates'):
            c.row_factory=__import__('sqlite3').Row
            data[table]=[dict(r) for r in c.execute('SELECT * FROM '+table)]
    return json.dumps(data,ensure_ascii=False,indent=2).encode('utf-8')

def restore_settings(store,data):
    if len(data)>10_000_000:raise ValueError('A beállításmentés túl nagy.')
    try:payload=json.loads(data)
    except (ValueError,UnicodeDecodeError):raise ValueError('Nem olvasható JSON-mentés.')
    if not isinstance(payload,dict) or payload.get('format')!='babolna-jutalek-settings-v1':raise ValueError('Nem az alkalmazás beállításmentése.')
    schema={'decisions':['freight','signature','month','note'],'overrides':['freight','month','note','created'],'rates':['month','payload']}
    checked={}
    for table,cols in schema.items():
        rows=payload.get(table,[])
        if not isinstance(rows,list) or len(rows)>50000:raise ValueError('Hibás beállításlista.')
        checked[table]=[]
        for row in rows:
            if not isinstance(row,dict) or set(row)!=set(cols) or not all(isinstance(v,str) for v in row.values()):raise ValueError('Hibás mentett beállítás.')
            if not valid_month(row['month']):raise ValueError('Hibás mentett hónap.')
            if table=='rates':
                try:rate=json.loads(row['payload']);assert dec(rate['rate'])>0
                except Exception:raise ValueError('Hibás mentett árfolyam.')
            checked[table].append(tuple(row[col] for col in cols))
    with store.connect() as c:
        for table,cols in schema.items():
            sql='INSERT OR REPLACE INTO '+table+' ('+','.join(cols)+') VALUES('+','.join('?' for _ in cols)+')'
            c.executemany(sql,checked[table])
        store.bump(c)
    return sum(len(v) for v in checked.values())

def restore_panel(st,store):
    st.subheader('Mentett beállítások visszatöltése')
    st.caption('Az átsorolásokat, dátumdöntéseket és árfolyamokat tölti vissza. Nem tartalmaz fuvarlistát vagy kifizetési állapotot.')
    uploaded=st.file_uploader('Beállításmentés (JSON)',type=['json'],key='restore_settings')
    if uploaded and st.button('Beállítások visszatöltése',key='restore_settings_button'):
        try:count=restore_settings(store,uploaded.getvalue())
        except ValueError as e:st.error(str(e))
        else:
            st.session_state.decision_feedback=f'Beállítások visszatöltve: {count} bejegyzés.'
            st.rerun()
