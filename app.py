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
st.set_page_config(page_title="Jutalékelszámolás | Bábolna Sped",page_icon=str(ROOT/"assets"/"logo_ui.png"),layout="wide",
                   initial_sidebar_state="collapsed")
NAVY,YELLOW,BG,BORDER,MUTED="#2A3756","#F2E47D","#F4F5F8","#E3E6EC","#5B6475"

def asset64(name):
    return base64.b64encode((ROOT/"assets"/name).read_bytes()).decode()

def brand():
    """A2 arculat: navy felső sáv logóval és morzsamenüvel, teljes szélességű tartalom."""
    st.markdown("<style>"
        "@font-face{font-family:Fira;src:url(data:font/ttf;base64,"+asset64("FiraSans-Regular.ttf")+")}"
        "@font-face{font-family:Fira;font-weight:700;src:url(data:font/ttf;base64,"+asset64("FiraSans-Bold.ttf")+")}"
        "@font-face{font-family:Montserrat;src:url(data:font/ttf;base64,"+asset64("Montserrat.ttf")+")}"
        f"""
html,body,[class*="css"],.stApp{{font-family:Fira,sans-serif;color:#1F2937}}
.stApp{{background:{BG}}}
header[data-testid="stHeader"]{{display:none}}
[data-testid="stSidebar"],[data-testid="collapsedControl"]{{display:none}}
.block-container{{padding-top:0;max-width:1360px}}
h1,h2,h3,h4{{font-family:Montserrat,sans-serif!important;color:{NAVY}}}
h2,h3{{font-size:22px!important}}
.topbar{{background:{NAVY};color:#fff;margin:0 calc(50% - 50vw) 24px;padding:0 calc(50vw - 50% + 1rem)}}
.topbar-in{{min-height:64px;display:flex;flex-wrap:wrap;align-items:center;gap:12px 20px}}
.topbar img{{height:40px;background:#fff;border-radius:8px;padding:2px}}
.crumb{{display:flex;align-items:center;gap:10px;font-size:14px}}
.crumb .apps{{background:rgba(255,255,255,.08);padding:9px 14px;border-radius:10px}}
.crumb .sep{{color:#7D88A3}} .crumb .cur{{color:{YELLOW};font-weight:600}}
.pagehead h1{{margin:0;font-size:28px;letter-spacing:-.01em}}
.pagehead p{{margin:4px 0 0;color:{MUTED};font-size:14px}}
.pill{{display:inline-block;font-size:12px;font-weight:600;color:{NAVY};background:{YELLOW};
       padding:3px 10px;border-radius:999px;margin-right:6px}}
.kpi,[data-testid="stMetric"]{{background:#fff;border:1px solid {BORDER};border-radius:14px;padding:18px 20px;height:100%;
      display:flex;flex-direction:column;gap:8px}}
.kpi.dark{{background:{NAVY};border-color:{NAVY};color:#fff}}
.kpi .l,[data-testid="stMetricLabel"] p{{font-size:13px;color:{MUTED}}} .kpi.dark .l{{color:#C9D0DF}}
.kpi .v,[data-testid="stMetricValue"]{{font-family:Montserrat,sans-serif;font-weight:700;font-size:30px;color:{NAVY};
         font-variant-numeric:tabular-nums;letter-spacing:-.02em}}
.kpi.dark .v{{color:#fff}}
.kpi .s{{font-size:12px;color:{MUTED}}} .kpi.dark .s{{color:#C9D0DF}}
.kpi.hl{{background:#FFFBEA;border-color:#DCC846}}
[data-testid="stTabs"] button[role="tab"]{{font-size:15px;height:44px}}
[data-testid="stTabs"] [aria-selected="true"]{{color:{NAVY};font-weight:600}}
[data-testid="stTabs"] [data-baseweb="tab-highlight"]{{background:{NAVY}}}
[data-testid="stAlert"]{{border-radius:12px}}
[data-testid="stExpander"],[data-testid="stFileUploader"]{{background:#fff;border-radius:12px}}
[data-testid="stExpander"] summary p{{font-family:Montserrat,sans-serif;font-weight:600;font-size:17px;color:{NAVY}}}
.rate-strip{{padding:16px 20px;border:1px solid {BORDER};background:white;border-radius:12px;margin:14px 0}}
div.stButton>button[kind="primary"],div.stDownloadButton>button{{background:{YELLOW};color:{NAVY};
    border:0;border-radius:10px;font-weight:600;min-height:44px}}
.foot{{color:{MUTED};font-size:12px;margin-top:28px;border-top:1px solid {BORDER};padding-top:10px}}
</style>""",unsafe_allow_html=True)
    st.markdown(f"""<div class="topbar"><div class="topbar-in">
      <img src="data:image/png;base64,{asset64('logo_ui.png')}" alt="Bábolna Sped">
      <div class="crumb"><span class="apps">Alkalmazások</span><span class="sep">/</span>
      <span class="cur">Jutalékelszámolás</span></div></div></div>""",unsafe_allow_html=True)

def kpi(col,label,value,sub="",style=""):
    col.markdown(f'<div class="kpi {style}"><span class="l">{label}</span><span class="v">{value}</span>'
                 +(f'<span class="s">{sub}</span>' if sub else "")+'</div>',unsafe_allow_html=True)

brand()

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
        st.markdown('<div class="pagehead"><h1>Jutalékelszámolás</h1><p>Belépés szükséges</p></div>',unsafe_allow_html=True)
        st.write("")
        entered=st.text_input("Belépési jelszó",type="password")
        if st.button("Belépés",type="primary"):
            if hmac.compare_digest(entered,password):
                st.session_state.authenticated=True
                st.rerun()
            else:st.error("Hibás jelszó.")
        st.stop()
auth()

store=session_store(st.session_state)
st.markdown('<div class="pagehead"><h1>Alvállalkozói jutalékelszámolás</h1>'
            '<p><span class="pill">10% jutalék · HUF elszámolás</span>'
            'Válassz hónapot és kapcsolattartót, majd töltsd le az aktuális kimutatást. Nincs havi lezárás.</p></div>',
            unsafe_allow_html=True)
st.write("")
feedback=st.session_state.pop("decision_feedback",None)
if feedback:st.success(feedback)
work_tab,reassign_tab,analytics_tab,help_tab=st.tabs(["Kimutatás készítése","Fuvar átsorolása","Jutalékelemzés","Használat és beállítások"])

@st.fragment(run_every="5s")
def download_status(job):
    # This fragment must never rerun the full app: doing so can replace the
    # deferred download registration before the browser has received its file.
    if job.error:
        st.error('A letöltéshez szükséges mentés nem sikerült: ' + job.error)
    elif job.status:
        if job.finished:
            st.success(job.status)
        else:
            st.info(job.status)
    if job.saved and job.reports:
        st.caption('Ha a letöltés nem indult el automatikusan, a kész fájlokat itt azonnal letöltheted:')
        stem = filename_stem(job.plan)
        a,b = st.columns(2)
        a.download_button('Mentett Excel letöltése',job.reports['report.xlsx'],stem+'.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',on_click='ignore',
            key='ready_x_'+job.snapshot['id'])
        b.download_button('Mentett PDF letöltése',job.reports['report.pdf'],stem+'.pdf',
            mime='application/pdf',on_click='ignore',key='ready_p_'+job.snapshot['id'])


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
    digest = hashlib.sha256(json_dump({k:v for k,v in plan.items() if k not in ('created','revision')}).encode() + frozen + CONTACT.encode()).hexdigest()
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
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key=key+'_x',on_click='ignore')
    b.download_button('PDF letöltése',lambda: job.download('pdf'),stem+'.pdf',mime='application/pdf',key=key+'_p',on_click='ignore')
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
    with st.expander("Forrásfájl",expanded=True):
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
        with st.expander("Elszámolási hónap és kapcsolattartó",expanded=True):
            a,b=st.columns(2)
            month=a.selectbox("Elszámolási hónap",choices,index=choices.index(previous) if previous in choices else len(choices)-1)
            contact=b.selectbox("Alvállalkozói saját kapcsolattartó",contacts)
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
            kpi(a,"Több hónapos, döntésre vár",len(unresolved),style="dark")
            if negatives:kpi(metrics[1],"Negatív árrés",len(negatives))
            kpi(c,"Hibás importált sor",len(errors))
            st.write("")
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
        rk=st.columns(3)
        kpi(rk[0],"Jutalék összesen",ft(plan["total"])+" Ft",style="dark")
        kpi(rk[1],"Elszámolt fuvarok",ft(len(plan["lines"]))+" db","A kiválasztott hónap és kapcsolattartó tételei")
        kpi(rk[2],"Árrés extrákkal",ft(margin)+" Ft","Pozitív és negatív árrések összege; jutalék csak a pozitív tételekre.",style="hl")
        st.write("")
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
st.markdown('<div class="foot">Árrés extra számlasorokkal • Teljesítés időpontja · A hónapdöntések automatikusan mentődnek. '
            'Excel vagy PDF letöltésekor a kész állapot a Drive-ra kerül. · v2.4.3 · Streamlit web</div>',unsafe_allow_html=True)
