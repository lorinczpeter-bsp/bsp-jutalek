"""Fuvarszám alapú átsorolás, kifizetési nyilvántartás és többdimenziós elemzés."""
import hashlib
from datetime import date
from decimal import Decimal
import pandas as pd
from core import origin,build_plan,dec,today
from exports import ft
from persistence import persist_choice
from runtime_cache import cached_analytics
from settlement import overrides,analytics,resolved_months

def money_display(frame):
    result=frame.copy()
    for col in result.columns:
        if col.endswith('(Ft)'):
            result[col]=result[col].map(lambda value:ft(value,2 if 'átlag' in col.lower() else 0)+' Ft')
    return result

def commission_chart_spec(kind):
    encoding={
        'x':{'field':'Hónap','type':'ordinal','sort':'ascending','axis':{'labelAngle':0}},
        'y':{'field':'Számított jutalék (Ft)','type':'quantitative','title':'Jutalék (Ft)',
             'axis':{'labelExpr':"replace(format(datum.value, ',.0f'), /,/g, ' ') + ' Ft'"}},
        'color':{'field':'Fuvarszervező','type':'nominal','scale':{'range':['#2A3756','#D5BA31','#547F9E','#A8B49C','#9E7483','#6C7F45']}},
        'tooltip':[{'field':'Hónap','type':'nominal'},{'field':'Fuvarszervező','type':'nominal'},{'field':'Jutalék','type':'nominal'}]}
    if kind=='bar':encoding['xOffset']={'field':'Fuvarszervező'}
    return {'mark':{'type':kind,'tooltip':True,**({'point':True} if kind=='line' else {'cornerRadiusTopLeft':3,'cornerRadiusTopRight':3})},
            'height':310,'encoding':encoding}

def available_months(records,store):
    months={m for r in records for m in r.months}|set(overrides(store).values())
    if not months:return []
    low=min(int(m[:4]) for m in months)-1;high=max(today().year+1,max(int(m[:4]) for m in months)+1)
    return [f'{y}-{m:02d}' for y in range(low,high+1) for m in range(1,13)]

def show_reassign(st,store,bundle):
    st.subheader('Fuvar átsorolása másik hónapra')
    st.caption('Add meg a fuvarszámot. Az átsorolás minden kimutatásban és az elemzésben érvényes; az eredeti kiegyenlítési dátum megmarad.')
    if not bundle:st.info('Előbb töltsd fel a teljes fuvarlistát a Kimutatás készítése fülön.');return
    source_hash,headers,records,errors=bundle
    key=st.text_input('Átsorolandó fuvarszám',key='reassign_freight').strip()
    if not key:return
    record=next((r for r in records if r.key==key),None)
    if not record:st.error('Ez a fuvarszám nincs a feltöltött fuvarlistában.');return
    st.write(f'{record.key} · {record.contact}')
    st.caption('Eredeti dátumok: '+', '.join(record.dates).replace('-','.'))
    current=origin(record,store)
    st.write('Jelenlegi elszámolási hónap: '+(current or 'Még nincs eldöntve'))
    months=available_months(records,store)
    with st.form('reassign_'+key):
        month=st.selectbox('Új elszámolási hónap',months,index=months.index(current) if current in months else months.index(record.months[-1]))
        note=st.text_input('Átsorolás megjegyzése (nem kötelező)')
        save=st.form_submit_button('Átsorolás mentése')
        if save:
            try:persist_choice(st,store,record,'override',month,note)
            except Exception as e:st.error('Sikertelen átsorolás: '+str(e))
            else:
                st.session_state.decision_feedback=f'Sikeres átsorolás: {key} → {month}. Az új hónap árfolyamával számolunk.'
                st.rerun()
    if key in overrides(store) and st.button('Kézi átsorolás visszavonása',key='clear_'+key):
        try:
            persist_choice(st,store,record,'clear_override')
        except Exception as exc:
            st.error('A visszavonás mentése sikertelen: '+str(exc));return
        st.session_state.decision_feedback=f'{key}: kézi átsorolás visszavonva; az eredeti dátum / korábbi dátumdöntés érvényes.'
        st.rerun()

def show_analytics(st,store,bundle):
    st.subheader('Jutalékelemzés')
    st.caption('A kiválasztott fuvarszervezők havi jutalékai, a kiválasztott hónapok összege és havi átlaga.')
    if not bundle:st.info('Előbb töltsd fel a teljes, több hónapot tartalmazó fuvarlistát.');return
    source_hash,headers,records,errors=bundle
    resolved=resolved_months(records,store)
    effective=sorted({m for m in resolved.values() if m})
    # A két szélső adathónap között az üres hónapok is választhatók, és beleszámítanak az átlag nevezőjébe.
    if effective:
        lo,hi=effective[0],effective[-1]
        months=[m for m in available_months(records,store) if lo<=m<=hi]
    else:months=[]
    contacts=sorted({r.contact for r in records})
    a,b=st.columns(2)
    selected_contacts=a.multiselect('Fuvarszervezők',contacts,default=contacts,key='analytics_contacts')
    closed=[m for m in months if m<today().strftime('%Y-%m')]
    selected_months=b.multiselect('Elszámolási hónapok',months,default=closed[-3:],key='analytics_months')
    if not selected_contacts or not selected_months:st.info('Válassz legalább egy fuvarszervezőt és egy hónapot.');return
    if errors:st.error(f'{len(errors)} hibás importált sor miatt az elemzés hiányos lehet.')
    with st.spinner('Elemzés és havi árfolyamok előkészítése...'):
        rows,issues=cached_analytics(st.session_state,store,source_hash,records,headers,selected_months,selected_contacts,analytics)
    for issue in issues:st.error(issue)
    unresolved=[r.key for r in records if r.contact in selected_contacts and resolved[r.key] is None and any(m in selected_months for m in r.months)]
    if unresolved:st.warning('Még nincs eldöntve a hónap, ezért nincs benne a számokban: '+', '.join(unresolved))
    frame=pd.DataFrame(rows)
    monthly=frame.groupby('Hónap')[['Számított jutalék (Ft)']].sum().sort_index()
    total=int(monthly['Számított jutalék (Ft)'].sum())
    a,b=st.columns(2)
    a.metric('Kiválasztott hónapok összege',ft(total)+' Ft')
    b.metric('Havi átlag',ft(Decimal(total)/len(set(selected_months)),2)+' Ft')
    st.caption(f'Az átlag nevezője {len(set(selected_months))} kiválasztott hónap, az üres hónapokkal együtt. Több fuvarszervező esetén a közös havi összeg átlaga.')
    st.subheader('Havi jutalék fuvarszervezőnként')
    chart_data=frame[['Hónap','Fuvarszervező','Számított jutalék (Ft)']].copy()
    chart_data['Jutalék']=chart_data['Számított jutalék (Ft)'].map(lambda x:ft(x)+' Ft')
    st.vega_lite_chart(chart_data,spec=commission_chart_spec('bar'),width='stretch')
    st.vega_lite_chart(chart_data,spec=commission_chart_spec('line'),width='stretch')
    st.subheader('Havi értékek')
    st.dataframe(money_display(monthly.reset_index()),hide_index=True)
    people=frame.groupby('Fuvarszervező')['Számított jutalék (Ft)'].sum().reset_index(name='Összesen (Ft)')
    people['Havi átlag (Ft)']=people['Összesen (Ft)']/len(set(selected_months))
    st.subheader('Összeg és havi átlag fuvarszervezőnként')
    st.dataframe(money_display(people),hide_index=True)
