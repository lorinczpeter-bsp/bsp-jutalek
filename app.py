"""Bábolna Sped - Alvállalkozói jutalékelszámolás."""
import base64, hashlib, hmac, io, os, sqlite3
from pathlib import Path
import pandas as pd
import streamlit as st
from core import (Store, load_workbook, build_plan, origin, export_columns, export_rows,
                  today, dec, valid_month, json_dump, CONTACT, MARGIN, DATE)
from exports import make_excel, make_pdf, ft, display_rows
from settlement import filename_stem,overrides
from ui_features import show_reassign,show_analytics
from web_session import session_store,settings_backup
from archive_ui import archive_client
from persistence import ExportJob, persist_choice
from auto_check import show_auto_check
from runtime_cache import parsed_source, monthly_plan
from settlement import resolved_months

ROOT=Path(__file__).parent
st.set_page_config(page_title="Jutalékelszámolás | Bábolna Sped",page_icon=str(ROOT/"assets"/"logo_ui.png"),layout="wide")

def auth():
    password=os.environ.get("JUTALEK_PASSWORD","")
    try:
        password=st.secrets.get("app_password",password)
    except (FileNotFoundError,st.errors.StreamlitSecretNotFoundError):
        pass
    if not password:
        try:
            drive_enabled=bool(st.secrets.get("drive_archive",{}).get("folder_id"))
        except (FileNotFoundError,st.errors.StreamlitSecretNotFoundError):
            drive_enabled=False
        if drive_enabled:
            st.error("A közös Drive-archívum használatához állíts be app_password belépési jelszót a Secretsben.")
            st.stop()
        return
    if not st.session_state.get("authenticated"):
        st.image(str(ROOT/"assets"/"logo_ui.png"),width=240)
        st.title("Jutalékelszámolás")
        entered=st.text_input("Belépési jelszó",type="password")
        if st.button("Belépés",type="primary"):
            if hmac.compare_digest(entered,password):
                st.session_state.authenticated=True
                st.rerun()
            else:st.error("Hibás jelszó.")
        st.stop()
auth()
def asset64(name):
    return base64.b64encode((ROOT/"assets"/name).read_bytes()).decode()
st.markdown("""
<style>
@font-face{font-family:Fira;src:url(data:font/ttf;base64,"""+asset64("FiraSans-Regular.ttf")+""")}
@font-face{font-family:Montserrat;src:url(data:font/ttf;base64,"""+asset64("Montserrat.ttf")+""")}
html,body,[class*="css"],.stApp{font-family:Fira,sans-serif;color:#2A3756}
.stApp{background:#F6F7F9}
h1,h2,h3{font-family:Montserrat,sans-serif!important;color:#2A3756}
[data-testid="stSidebar"]{background:#FFFFFF;border-right:1px solid #D1D5DB}
[data-testid="stSidebar"] [data-testid="stImage"]{background:white;padding:8px;border-radius:8px}
[data-testid="stSidebar"] p,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] label{color:#2A3756}
[data-testid="stMetric"]{background:white;padding:22px;border-radius:14px;border:1px solid #D1D5DB}
[data-testid="stMetricValue"]{font-size:30px;font-weight:700;letter-spacing:-0.5px}
[data-testid="stMetricLabel"]{font-size:13px;color:#657087}
[data-testid="stFileUploader"]{background:white;padding:18px;border:1px solid #D1D5DB;border-radius:14px}
button[data-baseweb="tab"]{border-radius:9px;padding:10px 18px}
button[data-baseweb="tab"][aria-selected="true"]{background:#F2E47D;color:#2A3756}
[data-testid="stAlert"]{border-radius:12px}
.report-cards{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px;margin:12px 0 18px}
.report-card{border:1px solid #D1D5DB;border-radius:14px;background:white;padding:24px}
.report-card.highlight{background:#FFFBEA;border-color:#DCC846}
.report-card p{margin:0 0 18px;font-size:14px;font-weight:600}
.report-card strong{display:block;font:700 30px Montserrat,sans-serif;letter-spacing:-0.5px}
.report-card small{display:block;margin-top:12px;color:#657087;font-size:12px}
.report-total{background:#2A3756;border-radius:14px;padding:24px;color:white;margin:4px 0 20px}
.report-total span{display:block;color:#F2E47D;font-size:11px;letter-spacing:1.5px;text-transform:uppercase}
.report-total strong{font:700 29px Montserrat,sans-serif}
.rate-strip{padding:16px 20px;border:1px solid #D1D5DB;background:white;border-radius:10px;margin:14px 0}
@media(max-width:700px){.report-cards{grid-template-columns:1fr}.report-card strong{font-size:24px}}
div.stButton>button[kind="primary"]{background:#F2E47D;color:#2A3756;border:0;font-weight:700}
div.stDownloadButton>button{border:1px solid #D1D5DB;border-radius:9px;background:#F2E47D;color:#2A3756;padding:12px 24px}
.block-container{padding-top:2rem;max-width:1700px}
.brandbar{background:#2A3756;color:white;padding:16px 24px;border-radius:12px;margin-bottom:22px}
.brandbar span{color:#F2E47D}
</style>""",unsafe_allow_html=True)

store=session_store(st.session_state)
with st.sidebar:
    st.image(str(ROOT/"assets"/"logo_ui.png"),width=230)
    st.markdown("## Jutalékelszámolás")
    st.caption("ALVÁLLALKOZÓI FUVAROK")
    st.markdown("**10% jutalék · HUF elszámolás**")
    st.caption("Árrés extra számlasorokkal • Teljesítés időpontja")
    st.divider()
    st.caption("A hónapdöntések automatikusan mentődnek. Excel vagy PDF letöltésekor a kész állapot a Drive-ra kerül.")
    st.caption("v2.4.1 · Streamlit web")

st.title("Alvállalkozói jutalékelszámolás")
st.caption("Válassz hónapot és kapcsolattartót, majd töltsd le az aktuális kimutatást. Nincs havi lezárás.")
feedback=st.session_state.pop("decision_feedback",None)
if feedback:st.success(feedback)
work_tab,reassign_tab,analytics_tab,help_tab=st.tabs(["Kimutatás készítése","Fuvar átsorolása","Jutalékelemzés","Használat és beállítások"])

def mark_download_started(job, target_count):
    st.session_state['export_polling'] = True
    st.session_state['polling_export_job'] = job
    st.session_state['polling_export_target'] = target_count


@st.fragment(run_every="2s" if st.session_state.get('export_polling') else None)
def download_status(job):
    active = st.session_state.get('polling_export_job', job)
    if (active.completed_count >= st.session_state.get('polling_export_target', active.completed_count+1)
            and st.session_state.get('export_polling')):
        st.session_state['export_polling'] = False
        st.session_state.pop('polling_export_job', None)
        st.rerun()
    job = active
    if job.error:
        st.error('A letöltéshez szükséges mentés nem sikerült: ' + job.error)
    elif job.status:
        if job.finished: st.success(job.status)
        else: st.info(job.status)


def downloads(plan,prefix):
    if not st.session_state.get('history_ready'):
        st.error('A mentett hónapdöntések beolvasása nem sikerült. Frissítsd az automatikus ellenőrzést a letöltés előtt.'); return
    if plan['blockers']:
        st.info('A letöltés a jelzett hibák rendezése után érhető el.'); return
    try:
        client = archive_client(st)
        if client is None:
            st.error('Letöltéshez állítsd be a Drive-mentést.'); return
    except ValueError as exc:
        st.error(str(exc)); return
    frozen = settings_backup(store)
    digest = hashlib.sha256(json_dump({k:v for k,v in plan.items() if k not in ('created','revision')}).encode() + frozen).hexdigest()
    key = prefix + '_' + today().isoformat() + '_' + digest
    jobs = st.session_state.setdefault('export_jobs', {})
    if key not in jobs:
        while len(jobs) >= 4:
            jobs.pop(next(iter(jobs)))
        jobs[key] = ExportJob(client, st.session_state.source_raw, st.session_state.source_name, frozen, plan)
    job = jobs[key]
    stem = filename_stem(plan)
    a,b = st.columns(2)
    a.download_button('Excel letöltése',lambda: job.download('xlsx'),stem+'.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key=key+'_x',on_click=mark_download_started,args=(job,job.completed_count+1))
    b.download_button('PDF letöltése',lambda: job.download('pdf'),stem+'.pdf',mime='application/pdf',key=key+'_p',on_click=mark_download_started,args=(job,job.completed_count+1))
    st.caption('Letöltéskor a kész Excel, PDF, forrás és hónapbesorolások automatikusan a Drive-ra mentődnek.')
    download_status(job)


def preview(plan):
    from html import escape
    from exports import centered_from
    cols=export_columns(plan);center_start=centered_from(plan)
    def cell(value,tag,index):
        align="center" if index>=center_start else "left"
        return f'<{tag} style="text-align:{align}">'+escape(str(value)).replace('\n','<br>')+f'</{tag}>'
    table='<div class="commission-preview"><table><thead><tr>'+''.join(cell(h,'th',i) for i,h in enumerate(cols))+'</tr></thead><tbody>'
    for row,line in zip(display_rows(plan),plan['lines']):
        row[-1]=ft(row[-1])+" Ft"
        table+='<tr'+(' class="negative"' if line['negative'] else '')+'>'+''.join(cell(v,'td',i) for i,v in enumerate(row))+'</tr>'
    table+='</tbody></table></div>'
    st.markdown('''<style>
    .commission-preview{overflow:auto;max-height:650px;border-radius:8px;background:white}
    .commission-preview table{width:100%;border-collapse:collapse;font-size:13px}
    .commission-preview th{position:sticky;top:0;background:#2A3756;color:white;padding:10px;z-index:1}
    .commission-preview td{padding:8px;vertical-align:top;border-bottom:1px solid #E9ECF1;min-width:80px}
    .commission-preview td:last-child{font-weight:bold;white-space:nowrap}
    .commission-preview tr:nth-child(even){background:#F2F4F7}
    .commission-preview tr.negative{color:#9C2222;background:#FFF0ED}
    </style>'''+table,unsafe_allow_html=True)
    if not plan['lines']:st.info("Nincs tétel a kiválasztott hónaphoz és kapcsolattartóhoz.")

with help_tab:
    st.subheader("Első használat")
    st.markdown("""
1. Töltsd fel a teljes, több hónapot és minden fuvarszervezőt tartalmazó fuvarlistát.
2. A rendszer visszatölti a mentett hónapdöntéseket, és ellenőrzi az előző három naptári hónapot.
3. Válassz hónapot és fuvarszervezőt. Rendezd a még eldöntetlen hónapbesorolásokat.
4. Töltsd le az Excelt vagy PDF-et. A letöltés előtt a kész állapot automatikusan a Drive-ra mentődik.

A hónapdöntés és átsorolás a rögzítésekor automatikusan mentődik, a visszavonással együtt.
A letöltés a kiválasztott fuvarszervező adott havi kész állapotát menti el összehasonlítási alapként.
A kész Excel és PDF, a teljes eredeti forrás és a beállítások együtt kerülnek megőrzésre.
Drive-hiba esetén a letöltés nem indul el mentés nélkül. Azonos változat Excel- és PDF-letöltése közös mentést használ.
Nincs havi lezárás, kézi állapotmentés vagy kifizetési nyilvántartás. Mindig a teljes számított jutalék látható.
A még le nem töltött adat nem válik új összehasonlítási alappá.

**Számítás:** a pozitív árrés 10%-a, tételenként egész forintra kerekítve. Negatív árrés: piros sor, 0 Ft jutalék.
Azonos pénznemnél az exportált extra árrés, vegyes HUF/EUR pénznemnél a forintosított díjak különbsége az alap.
Az elszámolási hónap pontos MNB havi átlagával számolunk; megjelenítéskor két tizedes látszik.
A folyamatban lévő hónap teljes havi árfolyama még nem ismert. Hiányos kimutatás nem tölthető le kész állapotként.
""")
    st.subheader("Árfolyam kézi rögzítése")
    st.caption("Csak automatikus lekérdezési hiba esetére. A forrást és az indoklást is rögzíteni kell; a mentett árfolyam nem írható felül.")
    with st.form("manual_rate"):
        rm=st.text_input("Árfolyam hónapja (ÉÉÉÉ-HH)",value=today().replace(day=1).isoformat()[:7])
        rv=st.text_input("Pontos MNB havi EUR/HUF átlag (teljes pontossággal)",value="")
        st.caption("A megadott pontos számot tároljuk és használjuk. Az árfolyam kijelzése mindenhol két tizedes.")
        reason=st.text_input("Forrás és indoklás")
        save=st.form_submit_button("Árfolyam rögzítése")
        if save:
            if not valid_month(rm) or rm>=today().strftime("%Y-%m") or not reason.strip() or not rv.strip():
                st.error("Lezárult hónap, pozitív árfolyam és forrás szükséges.")
            elif store.rate(rm):
                st.error("Ehhez a hónaphoz már van mentett árfolyam.")
            else:
                try:
                    store.save_rate(rm,{"rate":str(dec(rv)),"source":reason.strip(),"method":"Kézzel rögzített MNB havi átlag",
                                        "count":0,"daily":{},"retrieved":str(today())})
                except ValueError as e:st.error(str(e))
                else:st.success("Árfolyam elmentve.")

with work_tab:
    uploaded=st.file_uploader("Fuvarlista feltöltése",type=["xlsx"],
        help="A teljes eredeti exportot töltsd fel, a Fuvarlista munkalappal.")
    if uploaded:
        raw=uploaded.getvalue()
        st.session_state.source_raw=raw
        st.session_state.source_name=getattr(uploaded,"name","fuvarlista.xlsx")
        source_hash=hashlib.sha256(raw).hexdigest()
        try:
            source_hash,headers,records,errors=parsed_source(st.session_state,raw)
        except Exception as e:
            st.error(str(e));st.stop()
        st.caption(f"{len(records):,} feldolgozható tétel · {len(set(r.contact for r in records))} kapcsolattartó".replace(","," "))
        if not records:
            st.error("Nincs feldolgozható tétel a feltöltésben.")
            st.dataframe(pd.DataFrame(errors),hide_index=True)
            st.stop()
        st.session_state.source_bundle=(source_hash,headers,records,errors)
        show_auto_check(st,store,raw,st.session_state.source_name)
        resolved_month_map=resolved_months(records,store)
        all_months=sorted({m for r in records for m in r.months}|{m for m in resolved_month_map.values() if m}|set(overrides(store).values()))
        start=min(all_months)
        choices=all_months
        from datetime import timedelta
        previous=(today().replace(day=1)-timedelta(days=1)).strftime("%Y-%m")
        contacts=sorted(set(r.contact for r in records))
        a,b=st.columns(2)
        month=a.selectbox("Elszámolási hónap",choices,index=choices.index(previous) if previous in choices else len(choices)-1)
        contact=b.selectbox("Saját kapcsolattartó",contacts)
        try:
            with st.spinner("Elszámolás és MNB-árfolyamok ellenőrzése…"):
                plan=monthly_plan(st.session_state,store,records,headers,source_hash,month,contact,errors)
        except ValueError as e:
            st.info(str(e));st.stop()
        inspect_tab,list_tab=st.tabs(["Ellenőrzések","Teljes forráslista"])
        cross=[r for r in records if len(r.months)>1 and month in r.months and r.contact==contact]
        unresolved=[r for r in cross if resolved_month_map[r.key] is None]
        negatives=[r for r in plan["lines"] if r["negative"]]
        with inspect_tab:
            metrics=st.columns(3 if negatives else 2)
            a,c=metrics[0],metrics[-1]
            a.metric("Több hónapos, döntésre vár",len(unresolved))
            if negatives:metrics[1].metric("Negatív árrés",len(negatives))
            c.metric("Hibás importált sor",len(errors))
            if errors:st.error("Hibás importált sorok: a kimutatás hiányos lehet, javítsd a forrást.");st.dataframe(pd.DataFrame(errors),hide_index=True)
            if unresolved:
                st.warning("Válassz hónapot a még ellenőrizendő fuvarokhoz.")
                st.dataframe(pd.DataFrame([{"Fuvarszám":r.key,"Kapcsolattartó":r.contact,"Dátumok":", ".join(r.dates).replace("-","."),
                    "Árrés":float(r.margin),"Pénznem":r.currency} for r in unresolved]),hide_index=True)
                pending={r.key:r for r in unresolved}
                pending_key=hashlib.sha256(json_dump(list(pending)).encode()).hexdigest()[:12]
                chosen_key=st.selectbox("Ellenőrizendő fuvar",list(pending),format_func=lambda key:f"{key} | {pending[key].contact} | {', '.join(pending[key].months)}",key="decision_freight_"+pending_key)
                chosen=pending[chosen_key]
                with st.form("date_decision_"+chosen.signature):
                    assigned=st.selectbox("Hozzárendelendő hónap",chosen.months,index=chosen.months.index(month),key="assigned_"+chosen.signature)
                    note=st.text_input("Megjegyzés (nem kötelező)")
                    if st.form_submit_button("Hónap hozzárendelése"):
                        try:
                            persist_choice(st,store,chosen,'decision',assigned,note)
                        except (ValueError,sqlite3.Error,OSError) as e:
                            st.error("Sikertelen hónap-hozzárendelés: "+str(e))
                        else:
                            st.session_state.decision_feedback=f"Sikeres hozzárendelés: {chosen.key} → {assigned}. A fuvar kikerült az ellenőrizendők közül."
                            st.rerun()

            resolved=[r for r in cross if resolved_month_map[r.key] is not None]
            if resolved:
                with st.expander(f"Rendezett hónap-hozzárendelések ({len(resolved)})"):
                    st.dataframe(pd.DataFrame([{"Fuvarszám":r.key,"Elszámolási hónap":resolved_month_map[r.key],"Dátumok":", ".join(r.dates).replace("-",".")} for r in resolved]),hide_index=True)
            if negatives:
                with st.expander("Negatív árrések - jutalékuk 0 Ft"):
                    st.dataframe(pd.DataFrame([{"Fuvarszám":r["key"],"Dátumok":str(r["original"].get(DATE,"")),
                      "Árrés (Ft)":float(dec(r["margin_huf"])),"Jutalék (Ft)":0} for r in negatives]),hide_index=True)
        with list_tab:
            st.dataframe(pd.DataFrame([r.original for r in records]),hide_index=True)
        st.subheader(f"{month} · {contact}")
        from html import escape
        margin=sum(dec(x["margin_huf"]) for x in plan["lines"])
        st.markdown('<div class="report-cards">'
            '<div class="report-card"><p>Elszámolt fuvarok</p><strong>'+ft(len(plan["lines"]))+' db</strong><small>A kiválasztott hónap és kapcsolattartó tételei</small></div>'
            '<div class="report-card highlight"><p>Árrés extrákkal</p><strong>'+ft(margin)+' Ft</strong><small>Pozitív és negatív árrések összege; jutalék csak a pozitív tételekre.</small></div></div>'
            '<div class="report-total"><span>Jutalék összesen</span><strong>'+ft(plan["total"])+' Ft</strong></div>',unsafe_allow_html=True)
        for m,r in sorted(plan["rates"].items()):
            st.markdown('<div class="rate-strip"><b>MNB EUR/HUF</b> &nbsp; '+escape(m)+' havi átlag <b style="float:right">'+ft(r["rate"],2)+' Ft/EUR</b></div>',unsafe_allow_html=True)
        for msg in plan["blockers"]:st.error(msg)
        for msg in plan["warnings"]:st.warning(msg)
        preview(plan)
        downloads(plan,"draft")
    else:
        st.session_state.pop("source_bundle",None)
        st.session_state.pop("source_raw",None)
        st.session_state.pop("source_name",None)
        st.info("Töltsd fel a fuvarlistát az elszámolás elkészítéséhez.")

bundle=st.session_state.get("source_bundle")
with reassign_tab:show_reassign(st,store,bundle)
with analytics_tab:show_analytics(st,store,bundle)
