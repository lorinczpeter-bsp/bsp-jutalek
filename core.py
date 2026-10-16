"""Pénzügyi számítás, MNB árfolyam és változtathatatlan elszámolási napló."""
from __future__ import annotations
import calendar, hashlib, io, json, os, re, sqlite3, uuid
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from zoneinfo import ZoneInfo
import xml.etree.ElementTree as ET
import openpyxl
import requests

CONTACT = "Saját kapcsolattartó"
MARGIN = "Árrés extra számlasorokkal (állapot szerint)"
CURRENCY = MARGIN + " pénzneme"
DATE = "Teljesítés időpontja"
KEY = "Fuvarszám"
EXCLUDED = {CONTACT, "Teljesítettség állapota"}
PERCENT = Decimal("0.10")
MNB_URL = "https://www.mnb.hu/arfolyamok.asmx"
NAVY, YELLOW, GRAY = "2A3756", "F2E47D", "D1D5DB"

def today():
    return datetime.now(ZoneInfo("Europe/Budapest")).date()

def dec(value):
    if value is None or isinstance(value, bool):
        raise ValueError("Hiányzó vagy hibás összeg")
    s = str(value).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        result = Decimal(s)
    except InvalidOperation:
        raise ValueError(f"Hibás összeg: {value}")
    if not result.is_finite():
        raise ValueError("Nem véges összeg")
    return result

def money(value):
    return int(dec(value).quantize(Decimal("1"), rounding=ROUND_HALF_UP))

def valid_month(month):
    try:
        return bool(re.fullmatch(r"\d{4}-\d{2}", month)) and date.fromisoformat(month+"-01").strftime("%Y-%m") == month
    except (ValueError, TypeError):
        return False

def json_dump(obj):
    return json.dumps(obj, ensure_ascii=False, sort_keys=True, default=str)

def fingerprint(original):
    # A döntés csak addig érvényes, amíg a dátumok és a kapcsolattartó nem változnak.
    return hashlib.sha256(json_dump([original.get(KEY), original.get(CONTACT), original.get(DATE)]).encode()).hexdigest()

def parse_dates(value):
    if isinstance(value, (datetime, date)):
        return [value.strftime("%Y-%m-%d")]
    text = str(value or "").strip()
    matches = re.findall(r"\d{4}-\d{2}-\d{2}", text)
    residue = re.sub(r"\d{4}-\d{2}-\d{2}", "", text)
    if not matches or re.sub(r"[\s,;]+", "", residue):
        raise ValueError(f"Nem értelmezhető dátum: {text}")
    return sorted(set(date.fromisoformat(x).isoformat() for x in matches))

@dataclass
class Record:
    key: str
    contact: str
    original: dict
    margin: Decimal
    currency: str
    dates: list[str]
    signature: str

    @property
    def months(self):
        return sorted(set(x[:7] for x in self.dates))

def load_workbook(data: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    required = {KEY, CONTACT, MARGIN, CURRENCY, DATE}
    matches = []
    for sheet in wb:
        headers = [str(x or "").strip() for x in next(sheet.iter_rows(values_only=True), [])]
        if required <= set(headers):
            matches.append((sheet, headers))
    if len(matches) != 1:
        raise ValueError("Pontosan egy, a szükséges oszlopokat tartalmazó munkalap szükséges.")
    sheet, headers = matches[0]
    if len(headers) != len(set(headers)) or "" in headers:
        raise ValueError("Üres vagy ismétlődő oszlopfejléc.")
    records, errors, seen = [], [], set()
    for n, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
        if all(v is None for v in values):
            continue
        original = {h: (v.isoformat() if isinstance(v, (date, datetime)) else v) for h, v in zip(headers, values)}
        key = str(original.get(KEY) or "").strip()
        try:
            if not key:
                raise ValueError("Hiányzó fuvarszám")
            if key in seen:
                raise ValueError("Ismétlődő fuvarszám; automatikus összesítés nem engedélyezett")
            seen.add(key)
            contact = str(original.get(CONTACT) or "").strip()
            if not contact:
                raise ValueError("Hiányzó kapcsolattartó")
            currency = str(original.get(CURRENCY) or "").upper().strip()
            if currency not in {"EUR", "HUF"}:
                raise ValueError("Az árrés pénzneme csak EUR vagy HUF lehet")
            dates = parse_dates(values[headers.index(DATE)])
            records.append(Record(key, contact, original, dec(original[MARGIN]), currency, dates, fingerprint(original)))
        except (ValueError, InvalidOperation) as exc:
            errors.append({"Excel-sor": n, "Fuvarszám": key, "Hiba": str(exc)})
    wb.close()
    if not records and not errors:
        raise ValueError("A munkalap nem tartalmaz tételeket.")
    return headers, records, errors

class Store:
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as c:
            c.executescript("""
            CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            INSERT OR IGNORE INTO meta VALUES ('revision','0');
            CREATE TABLE IF NOT EXISTS decisions (freight TEXT PRIMARY KEY, signature TEXT, month TEXT, note TEXT);
            CREATE TABLE IF NOT EXISTS rates (month TEXT PRIMARY KEY, payload TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS closings (
                id TEXT PRIMARY KEY, month TEXT NOT NULL, contact TEXT NOT NULL, created TEXT NOT NULL,
                source_hash TEXT NOT NULL, payload TEXT NOT NULL, UNIQUE(month,contact));
            CREATE TABLE IF NOT EXISTS ledger (
                closing_id TEXT, freight TEXT, contact TEXT, target TEXT NOT NULL,
                PRIMARY KEY(closing_id,freight,contact));
            """)
        from settlement import init_registry
        init_registry(self)

    @contextmanager
    def connect(self):
        c = sqlite3.connect(self.path, timeout=30)
        try:
            c.execute("PRAGMA foreign_keys=ON")
            with c:
                yield c
        finally:
            c.close()

    def override(self,record):
        with self.connect() as c:
            row=c.execute("SELECT month FROM overrides WHERE freight=?",(record.key,)).fetchone()
        return row[0] if row else None

    def revision(self):
        with self.connect() as c:
            return int(c.execute("SELECT value FROM meta WHERE key='revision'").fetchone()[0])

    @staticmethod
    def bump(c):
        c.execute("UPDATE meta SET value=CAST(value AS INTEGER)+1 WHERE key='revision'")

    def setting(self, key, default=None):
        with self.connect() as c:
            r = c.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        return r[0] if r else default

    def set_setting(self, key, value):
        with self.connect() as c:
            c.execute("INSERT OR REPLACE INTO meta VALUES (?,?)", (key,str(value)))
            self.bump(c)

    def decision(self, record):
        with self.connect() as c:
            r = c.execute("SELECT month FROM decisions WHERE freight=? AND signature=?", (record.key,record.signature)).fetchone()
        return r[0] if r else None

    def save_decision(self, record, month, note=""):
        if month not in record.months:
            raise ValueError("A felsorolt hónapok egyikét válaszd ki.")
        note=note.strip() or "Kézi hónap-hozzárendelés az ellenőrző felületen."
        with self.connect() as c:
            c.execute("INSERT OR REPLACE INTO decisions VALUES (?,?,?,?)",(record.key,record.signature,month,note))
            self.bump(c)
        if self.decision(record)!=month:
            raise ValueError("A hozzárendelést nem sikerült visszaolvasni. Próbáld újra a mentést.")

    def rate(self, month):
        with self.connect() as c:
            r = c.execute("SELECT payload FROM rates WHERE month=?", (month,)).fetchone()
        return json.loads(r[0]) if r else None

    def save_rate(self, month, payload):
        if not valid_month(month) or dec(payload["rate"]) <= 0:
            raise ValueError("Hibás árfolyam")
        # Lezárt számítás alapjául szolgáló árfolyamot nem írunk felül.
        with self.connect() as c:
            c.execute("INSERT OR IGNORE INTO rates VALUES (?,?)",(month,json_dump(payload)))

    def closings(self):
        with self.connect() as c:
            return [{"id":x[0],"month":x[1],"contact":x[2],"created":x[3]} for x in c.execute("SELECT id,month,contact,created FROM closings ORDER BY month DESC,created DESC")]

    def closing(self, close_id):
        with self.connect() as c:
            r = c.execute("SELECT payload FROM closings WHERE id=?", (close_id,)).fetchone()
        return json.loads(r[0]) if r else None

    def latest(self):
        with self.connect() as c:
            rows = c.execute("""SELECT l.freight,l.contact,l.target FROM ledger l
                JOIN closings c ON c.id=l.closing_id ORDER BY c.month,c.created""").fetchall()
        return {(r[0],r[1]):json.loads(r[2]) for r in rows}

    def close(self, plan):
        if plan["blockers"]:
            raise ValueError("Nyitott ellenőrzések mellett nem zárható az elszámolás.")
        if plan["month"] >= today().strftime("%Y-%m"):
            raise ValueError("Csak teljesen lezárult naptári hónap zárható.")
        with self.connect() as c:
            c.execute("BEGIN IMMEDIATE")
            actual = int(c.execute("SELECT value FROM meta WHERE key='revision'").fetchone()[0])
            if actual != plan["revision"]:
                raise ValueError("Az adatok időközben megváltoztak. Készíts új előnézetet.")
            last = c.execute("SELECT MAX(month) FROM closings WHERE contact=?", (plan["contact"],)).fetchone()[0]
            if last and plan["month"] <= last:
                raise ValueError("A hónap már lezárt, vagy későbbi hónap lezárása már megtörtént.")
            close_id = str(uuid.uuid4())
            final = dict(plan, id=close_id, status="LEZÁRT", created=datetime.now(ZoneInfo("Europe/Budapest")).isoformat(timespec="seconds"))
            c.execute("INSERT INTO closings VALUES (?,?,?,?,?,?)",(close_id,plan["month"],plan["contact"],final["created"],plan["source_hash"],json_dump(final)))
            for line in plan["lines"]:
                c.execute("INSERT INTO ledger VALUES (?,?,?,?)",(close_id,line["key"],plan["contact"],json_dump(line["target"])))
            self.bump(c)
        return close_id

    def backup(self):
        out = io.BytesIO()
        # serialize() SQLite konzisztens pillanatképet ad, megnyitott kapcsolat mellett is.
        with self.connect() as c:
            return c.serialize()

def fetch_mnb(month):
    if not valid_month(month):
        raise ValueError("Hibás hónap")
    if month >= today().strftime("%Y-%m"):
        raise ValueError("Teljes havi MNB-átlag még nem áll rendelkezésre.")
    y,m = map(int,month.split("-"))
    end = calendar.monthrange(y,m)[1]
    body = f"""<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body>
    <GetExchangeRates xmlns="http://www.mnb.hu/webservices/"><startDate>{month}-01</startDate>
    <endDate>{month}-{end:02d}</endDate><currencyNames>EUR</currencyNames>
    </GetExchangeRates></s:Body></s:Envelope>"""
    last_error = None
    used_url = MNB_URL
    # Az MNB WSDL-je HTTP végpontot is hirdet; kizárólag nyilvános árfolyamkérés.
    for endpoint in [MNB_URL, MNB_URL.replace("https://", "http://")]:
        try:
            response = requests.post(endpoint,data=body.encode(),headers={
                "Content-Type":"text/xml; charset=utf-8",
                "SOAPAction":'"http://www.mnb.hu/webservices/MNBArfolyamServiceSoap/GetExchangeRates"'},timeout=20)
            response.raise_for_status()
            used_url = endpoint
            break
        except requests.RequestException as exc:
            last_error = exc
    else:
        raise ValueError(f"MNB lekérdezési hiba: {last_error}")
    envelope = ET.fromstring(response.content)
    result = next((x for x in envelope.iter() if x.tag.split("}")[-1]=="GetExchangeRatesResult"),None)
    if result is None or not result.text:
        raise ValueError("Nem értelmezhető MNB-válasz")
    root = ET.fromstring(result.text)
    daily = {}
    for day in root.findall(".//Day"):
        d = date.fromisoformat(day.attrib["date"])
        if d.strftime("%Y-%m") != month:
            continue
        for rate in day.findall("Rate"):
            if rate.attrib.get("curr") == "EUR":
                daily[d.isoformat()] = str(dec(rate.text)/dec(rate.attrib.get("unit","1")))
    if len(daily)<15:
        raise ValueError("Hiányos MNB-adatsor; a havi átlag nem használható.")
    mean = sum((dec(x) for x in daily.values()),Decimal(0))/len(daily)
    return {"rate":str(mean),"source":used_url,"method":"A hónap közzétett napi EUR/HUF jegyzéseinek számtani átlaga",
            "count":len(daily),"daily":daily,"retrieved":datetime.now(ZoneInfo("Europe/Budapest")).isoformat()}

def origin(record, store):
    return store.override(record) or (record.months[0] if len(record.months)==1 else store.decision(record))

def build_plan(records, headers, store, month, contact, source_hash, rate_provider=None, errors=None, snapshot=False):
    initial_revision = store.revision()
    start = min([month]+[m for r in records for m in r.months]) if snapshot else store.setting("start_month")
    if not start:
        raise ValueError("Állítsd be az első kezelt elszámolási hónapot.")
    if month < start:
        raise ValueError("A kiválasztott hónap az induló hónap előtti.")
    closings = [] if snapshot else store.closings()
    last = max((c["month"] for c in closings if c["contact"]==contact),default="")
    if last and month <= last:
        raise ValueError("Ezt vagy egy későbbi hónapot már lezártad. Az Előzményekben letölthető.")
    latest = {} if snapshot else store.latest()
    if snapshot:records=[r for r in records if r.contact==contact]
    current = {r.key:r for r in records}
    from settlement import resolved_months
    origins=resolved_months(records,store)
    lines, blockers, warnings, rates = [], [], [], {}
    failed_rates = set()
    if errors:
        blockers.append(f"{len(errors)} hibás importált sor; javítsd a forrásfájlt.")
    for r in records:
        om = origins[r.key]
        affected = r.contact==contact or (r.key,contact) in latest
        if affected and om is None and ((month in r.months) if snapshot else (max(r.months)>=start and min(r.months)<=month)):
            blockers.append(f"{r.key}: több hónapot érintő kiegyenlítés; hónapválasztás szükséges.")
    for (key,owner),old in latest.items():
        if owner==contact and key not in current and old.get("commission",0)!=0:
            warnings.append(f"{key}: az előzőleg elszámolt tétel hiányzik az új fájlból; nincs automatikus visszavonás.")
    for r in records:
        old = latest.get((r.key,contact))
        om = origins[r.key]
        if om is None:
            continue
        if snapshot and om != month:
            continue
        if r.contact != contact:
            if old is None or old["commission"]==0:
                continue
            target = dict(old,commission=0,margin_huf="0",reason="Kapcsolattartó megváltozott")
            kind = "Kapcsolattartó-korrekció"
            delta = -old["commission"]
        else:
            if om < start or om > month:
                if old and old["commission"]:
                    blockers.append(f"{r.key}: a már elszámolt tétel dátuma a kezelt időszakon kívülre változott.")
                continue
            fee_currency=str(r.original.get("Alvállalkozói díj pénzneme",r.currency)).upper().strip()
            revenue_currency=str(r.original.get("Díj pénzneme",r.currency)).upper().strip()
            mixed=fee_currency!=revenue_currency
            if mixed and {fee_currency,revenue_currency}!={"EUR","HUF"}:
                blockers.append(f"{r.key}: a díjak pénzneme csak EUR vagy HUF lehet.")
                continue
            fx = Decimal(1)
            if r.currency=="EUR" or mixed:
                if om in failed_rates:
                    continue
                if om not in rates:
                    payload = store.rate(om)
                    if payload is None:
                        try:
                            payload = (rate_provider or fetch_mnb)(om)
                            store.save_rate(om,payload)
                            payload = store.rate(om)
                        except Exception as e:
                            failed_rates.add(om)
                            blockers.append(f"{om}: MNB-árfolyam nem elérhető ({e}).")
                            continue
                    rates[om] = payload
                fx = dec(rates[om]["rate"])
            mh = r.margin*fx
            calculated_original=dict(r.original)
            if mixed:
                try:
                    revenue=dec(r.original.get("Díj"))
                    cost=dec(r.original.get("Alvállalkozói díj"))
                except ValueError as e:
                    blockers.append(f"{r.key}: vegyes pénznemű díj nem számítható ({e}).")
                    continue
                mh=revenue*(fx if revenue_currency=="EUR" else 1)-cost*(fx if fee_currency=="EUR" else 1)
                calculated_original[MARGIN]=str(mh)
                calculated_original[CURRENCY]="HUF"
            total = money(mh*PERCENT) if mh > 0 else 0
            target = {"commission":total,"margin_huf":str(mh),"margin":str(r.margin),"currency":r.currency,
                      "rate":str(fx),"origin":om,"original":calculated_original,"signature":r.signature}
            delta = total-(old["commission"] if old else 0)
            if old:
                changed = any(target.get(k)!=old.get(k) for k in ["commission","margin","currency","origin"])
                if not changed:
                    continue
                kind = "Árrés-/dátumkorrekció"
            else:
                kind = "Tárgyhavi" if om==month else "Utólagos tétel"
        lines.append({"key":r.key,"kind":kind,"origin":target["origin"],"original":target["original"],
                      "rate":target["rate"],"margin_huf":target["margin_huf"],
                      "previous":old["commission"] if old else 0,"target_commission":target["commission"],
                      "commission":delta,"negative":dec(target["margin_huf"])<0,"target":target,"source_original":r.original})
    lines.sort(key=lambda x:(x["kind"]!="Tárgyhavi",x["origin"],x["key"]))
    current_sum = sum(x["commission"] for x in lines if x["kind"]=="Tárgyhavi")
    correction_sum = sum(x["commission"] for x in lines if x["kind"]!="Tárgyhavi")
    if store.revision() != initial_revision:
        blockers.append("Az adatok számítás közben megváltoztak. Frissítsd az előnézetet.")
    plan={"month":month,"contact":contact,"headers":[h for h in headers if h not in EXCLUDED],
            "source_hash":source_hash,"percent":"10%","status":"AKTUÁLIS ÁLLAPOT" if snapshot else "TERVEZET","lines":lines,
            "current":current_sum,"corrections":correction_sum,"total":current_sum+correction_sum,
            "blockers":list(dict.fromkeys(blockers)),"warnings":warnings,"rates":rates,
            "revision":initial_revision,"start_month":start,"snapshot":snapshot,
            "created":datetime.now(ZoneInfo("Europe/Budapest")).isoformat(timespec="seconds")}
    from settlement import payments
    registry=payments(store) if snapshot else {}
    paid_keys=[];actual_paid=Decimal(0);mismatches=[]
    for line in lines:
        line["earned_commission"]=line["commission"]
        line["paid"]=line["key"] in registry
        if line["paid"]:
            entry=registry[line["key"]]
            line["payment"]=entry
            actual_paid+=dec(entry["amount"])
            if abs(dec(entry["amount"])-line["earned_commission"])>1 or entry["contact"]!=contact:
                mismatches.append(line["key"])
            line["commission"]=0
            paid_keys.append(line["key"])
    plan["earned_total"]=sum(x["earned_commission"] for x in lines)
    plan["paid_total"]=sum(x["earned_commission"] for x in lines if x["paid"])
    plan["actual_paid_total"]=str(actual_paid)
    plan["paid_keys"]=paid_keys
    plan["current"]=sum(x["commission"] for x in lines if x["kind"]=="Tárgyhavi")
    plan["corrections"]=sum(x["commission"] for x in lines if x["kind"]!="Tárgyhavi")
    plan["total"]=plan["current"]+plan["corrections"]
    if paid_keys:plan["warnings"].append(f"{len(paid_keys)} fuvar kifizetettként nyilvántartva: újra fizetendő jutalékuk 0 Ft. Fuvarok: "+", ".join(paid_keys[:12])+(f" (+{len(paid_keys)-12} további fuvar)" if len(paid_keys)>12 else ""))
    if mismatches:plan["warnings"].append("A korábbi kifizetés összege vagy fuvarszervezője eltér a jelenlegi számítástól; nincs automatikus ráfizetés: "+", ".join(mismatches))
    if any(x["source_original"].get("Díj pénzneme")!=x["source_original"].get("Alvállalkozói díj pénzneme") for x in lines):
        plan["warnings"].append("Vegyes pénznemnél a díjakat az elszámolási hónap pontos árfolyamán forintosítjuk. Az Árrés extrákkal mező itt a számított árrés, HUF-ban; külön extra számlasorok nem adódnak hozzá.")
    return plan


def compact_fields(plan):
    """Megjelenítési mezők: az eredeti adatok pénznemmel együtt, adatvesztés nélkül."""
    pairs = {
        "Alvállalkozói díj": ("Alvállalkozói díj pénzneme", "Alváll.\ndíj"),
        "Díj": ("Díj pénzneme", "Fuvardíj"),
        MARGIN: (CURRENCY, "Árrés\nextrákkal"),
        "Árrés (állapot szerint)": ("Árrés (állapot szerint) pénzneme", "Árrés"),
    }
    invoice_fields = ["Számlázás", "Számla sorszáma"]
    labels = {KEY: "Fuvarszám", DATE: "Kiegyenlítés\ndátuma"}
    consumed, fields = {"Árrés (állapot szerint)","Árrés (állapot szerint) pénzneme"}, []
    for h in plan["headers"]:
        if h in consumed:
            continue
        if h in invoice_fields and all(k in plan["headers"] for k in invoice_fields):
            fields.append({"label": "Számla", "source": invoice_fields})
            consumed.update(invoice_fields)
        elif h in pairs and pairs[h][0] in plan["headers"]:
            currency,label = pairs[h]
            fields.append({"label":label,"source":h,"currency":currency})
            consumed.update([h,currency])
        else:
            fields.append({"label":labels.get(h,h),"source":h})
            consumed.add(h)
    return fields

def export_columns(plan):
    return [f["label"] for f in compact_fields(plan)] + ["Jutalék (Ft)"]

def export_rows(plan):
    rows=[]
    for line in plan["lines"]:
        row=[]
        original=line["original"]
        for field in compact_fields(plan):
            source=field["source"]
            if isinstance(source,list):
                a,b=[str(original.get(k) or "").strip() for k in source]
                value=a if a==b else "\n".join(label+": "+v for label,v in [("Számlázás",a),("Számlaszám",b)] if v)
            else:
                value=original.get(source,"")
                if field.get("currency") and value is not None and value!="":
                    try:
                        value=float(dec(value))
                    except ValueError:
                        pass
                if value is None:
                    value=""
                if source==DATE:
                    value=re.sub(r"(\d{4})-(\d{2})-(\d{2})",r"\1.\2.\3",str(value))
            row.append(value)
        rows.append(row+[line["commission"]])
    return rows
