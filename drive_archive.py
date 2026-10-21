"""Immutable Drive snapshots and freight-level comparisons. No payment inference."""
from __future__ import annotations

import hashlib
import io
import json
import re
import uuid
import zipfile
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from zoneinfo import ZoneInfo

from core import (CONTACT, CURRENCY, DATE, KEY, MARGIN, build_plan, dec,
                  load_workbook, origin, valid_month)
from web_session import settings_backup

APP_VERSION = '2.2.0'
CALC_VERSION = 'commission-v2-mixed-monthly-fx'
FORMAT = 'babolna-jutalek-archive-v1'
MAX_SOURCE = 40 * 1024 * 1024
MAX_ARCHIVE = 100 * 1024 * 1024
MAX_EXPANDED = 160 * 1024 * 1024
NUMERIC = {MARGIN, 'Díj', 'Alvállalkozói díj', 'Árrés (állapot szerint)'}


class ArchiveError(ValueError):
    pass


def number(value):
    """Canonical decimal text without float arithmetic or cosmetic differences."""
    d = dec(value)
    text = format(d, 'f')
    return '0' if not d else (text.rstrip('0').rstrip('.') if '.' in text else text)


def canonical(record):
    result = {}
    for key, value in record.original.items():
        if key == DATE:
            value = ', '.join(sorted(set(record.dates)))
        elif key in NUMERIC and value not in (None, ''):
            try:
                value = number(value)
            except ValueError:
                # Unused legacy/extra fields may contain descriptive text.
                # Calculation-required fields are validated by build_plan.
                value = str(value).strip()
        elif isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            value = number(value)
        elif isinstance(value, str):
            value = value.strip()
            if key.endswith('pénzneme'):
                value = value.upper()
        elif value is None:
            value = ''
        result[key] = value
    return result


def build_snapshot(raw, source_name, store, months, note='', rate_provider=None):
    """Capture the entire original source; verify calculations for explicit full months."""
    if not raw or len(raw) > MAX_SOURCE:
        raise ArchiveError('Üres vagy 40 MB-nál nagyobb forrásfájl.')
    headers, records, errors = load_workbook(raw)
    if errors:
        raise ArchiveError(f'{len(errors)} hibás forrássor miatt nem menthető ellenőrzött állapot.')
    if not records:
        raise ArchiveError('Nincs archiválható fuvar.')
    months = sorted(set(months))
    if not months or any(not valid_month(m) for m in months):
        raise ArchiveError('Válassz legalább egy érvényes ellenőrzött hónapot.')
    resolved = {r.key: origin(r, store) for r in records}
    pending = [r.key for r in records if resolved[r.key] is None and set(r.months) & set(months)]
    if pending:
        raise ArchiveError('Még nincs eldöntve az elszámolási hónap: ' + ', '.join(pending[:15]))
    calculated, rates = {}, {}
    source_hash = hashlib.sha256(raw).hexdigest()
    for month in months:
        contacts = sorted({r.contact for r in records if resolved[r.key] == month})
        for contact in contacts:
            plan = build_plan(records, headers, store, month, contact, source_hash,
                              snapshot=True, rate_provider=rate_provider)
            # MNB cache warming legitimately changes the local revision on the first pass.
            if plan['blockers'] == ['Az adatok számítás közben megváltoztak. Frissítsd az előnézetet.']:
                plan = build_plan(records, headers, store, month, contact, source_hash,
                                  snapshot=True, rate_provider=rate_provider)
            if plan['blockers']:
                raise ArchiveError(f'{month} / {contact}: ' + '; '.join(plan['blockers']))
            rates.update(plan['rates'])
            for line in plan['lines']:
                calculated[line['key']] = {
                    'commission': line['earned_commission'],
                    'margin_huf': number(line['margin_huf']), 'rate': number(line['rate'])}
    rows = {}
    for record in records:
        month = resolved[record.key]
        if month in months and record.key not in calculated:
            raise ArchiveError(f'{record.key}: hiányos számítás, nem menthető.')
        rows[record.key] = {'original': canonical(record), 'month': month,
                            'contact': record.contact, 'calculation': calculated.get(record.key)}
    return {
        'format': FORMAT, 'app_version': APP_VERSION, 'calculation_version': CALC_VERSION,
        'id': uuid.uuid4().hex,
        'created': datetime.now(ZoneInfo('Europe/Budapest')).isoformat(timespec='microseconds'),
        'source_name': str(source_name).replace('\\', '/').split('/')[-1][:200],
        'source_sha256': source_hash, 'months': months, 'note': note.strip()[:2000],
        'records': rows, 'rates': rates, 'headers': headers,
        'settings': json.loads(settings_backup(store)),
        # Intentionally no payment registry: an archive never establishes payment.
    }


def encode_snapshot(snapshot, raw):
    if snapshot['source_sha256'] != hashlib.sha256(raw).hexdigest():
        raise ArchiveError('A forrásfájl és az állapot nem egyezik.')
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('source.xlsx', raw)
        z.writestr('snapshot.json', json.dumps(snapshot, ensure_ascii=False, sort_keys=True))
    data = out.getvalue()
    if len(data) > MAX_ARCHIVE:
        raise ArchiveError('Az archívum túl nagy.')
    return data


def decode_snapshot(data):
    """Do not extract untrusted archive paths; validate all sizes and raw-source digest."""
    try:
        if len(data) > MAX_ARCHIVE:
            raise ValueError()
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            if sorted(z.namelist()) != ['snapshot.json', 'source.xlsx']:
                raise ValueError()
            if sum(i.file_size for i in z.infolist()) > MAX_EXPANDED:
                raise ValueError()
            raw = z.read('source.xlsx')
            if len(raw) > MAX_SOURCE:
                raise ValueError()
            snapshot = json.loads(z.read('snapshot.json'))
        if not isinstance(snapshot, dict) or snapshot.get('format') != FORMAT or snapshot.get('source_sha256') != hashlib.sha256(raw).hexdigest():
            raise ValueError()
        for field in ('created', 'source_name', 'app_version', 'calculation_version', 'note'):
            if not isinstance(snapshot[field], str):
                raise ValueError()
        datetime.fromisoformat(snapshot['created'])
        if not isinstance(snapshot['settings'], dict) or snapshot['settings'].get('format') != 'babolna-jutalek-settings-v1':
            raise ValueError()
        if not re.fullmatch('[a-f0-9]{32}', snapshot['id']):
            raise ValueError()
        if not snapshot['months'] or any(not valid_month(m) for m in snapshot['months']):
            raise ValueError()
        if not isinstance(snapshot['records'], dict) or not snapshot['records']:
            raise ValueError()
        for key, row in snapshot['records'].items():
            if not key or not isinstance(row['original'], dict) or not row['contact']:
                raise ValueError()
            if row['month'] is not None and not valid_month(row['month']):
                raise ValueError()
            calc = row['calculation']
            if row['month'] in snapshot['months'] and calc is None:
                raise ValueError()
            if calc is not None:
                if type(calc['commission']) is not int or calc['commission'] < 0:
                    raise ValueError()
                dec(calc['margin_huf']); dec(calc['rate'])
        _, source_records, source_errors = load_workbook(raw)
        if source_errors or {r.key for r in source_records} != set(snapshot['records']):
            raise ValueError()
        if any(canonical(r) != snapshot['records'][r.key]['original'] or
               r.contact != snapshot['records'][r.key]['contact'] for r in source_records):
            raise ValueError()
        return snapshot, raw
    except (KeyError, TypeError, ValueError, OSError, zipfile.BadZipFile, RuntimeError) as exc:
        raise ArchiveError('Sérült vagy nem támogatott jutalékarchívum. Nem használható összehasonlításhoz.') from exc


def compare_snapshots(previous, current, months=None):
    """Compare complete selected periods; missing rows are never called deletions."""
    periods = set(months or previous['months'])
    if not periods <= set(previous['months']) or not periods <= set(current['months']):
        raise ArchiveError('Csak mindkét állapotban ellenőrzött hónapok hasonlíthatók össze.')
    changes, detail, effects = [], [], defaultdict(lambda: [0, 0])
    old_rows, new_rows = previous['records'], current['records']
    def contribution(row):
        if row and row['month'] in periods:
            if row['calculation'] is None:
                raise ArchiveError('Hiányzó számított jutalék a vizsgált hónapban.')
            return row['calculation']['commission']
        return 0
    for key in sorted(set(old_rows) | set(new_rows)):
        old, new = old_rows.get(key), new_rows.get(key)
        if not any(r and r['month'] in periods for r in (old, new)):
            continue
        a, b = contribution(old), contribution(new)
        if old and old['month'] in periods and new and new['month'] is None:
            raise ArchiveError(f'{key}: a korábbi hónaphoz tartozó fuvar új hónapja nincs eldöntve; előbb rendezd a besorolást.')
        for index, row in enumerate((old, new)):
            if row and row['month'] in periods:
                effects[(row['month'], row['contact'])][index] += contribution(row)
        reasons, fields = [], []
        if old is None:
            status = 'Új a vizsgált forrásban'
            reasons.append('Új forrássor')
        elif new is None:
            status = 'Hiányzik az új forrásból'
            reasons.append('Lehet szűkebb export is; ellenőrzés szükséges')
        else:
            status = 'Módosult'
            for field in sorted(set(old['original']) | set(new['original'])):
                before, after = old['original'].get(field, ''), new['original'].get(field, '')
                if before != after:
                    fields.append(field)
                    detail.append({'Fuvarszám': key, 'Mező': field, 'Korábban': str(before), 'Most': str(after)})
            if fields:
                reasons.append('Forrásadat')
            if old['month'] != new['month']:
                reasons.append('Hónapbesorolás')
                detail.append({'Fuvarszám': key, 'Mező': 'Elszámolási hónap',
                               'Korábban': old['month'] or 'Nincs eldöntve', 'Most': new['month'] or 'Nincs eldöntve'})
            oc, nc = old.get('calculation'), new.get('calculation')
            if oc and nc:
                if oc['rate'] != nc['rate']:
                    reasons.append('Árfolyam')
                    detail.append({'Fuvarszám': key, 'Mező': 'Alkalmazott pontos árfolyam',
                                   'Korábban': oc['rate'], 'Most': nc['rate']})
                if oc['margin_huf'] != nc['margin_huf']:
                    detail.append({'Fuvarszám': key, 'Mező': 'Számított árrés (Ft)',
                                   'Korábban': oc['margin_huf'], 'Most': nc['margin_huf']})
            if previous['calculation_version'] != current['calculation_version']:
                reasons.append('Számítási szabály verziója')
            if not reasons and a != b:
                reasons.append('Számítási eltérés – egyeztetendő')
            if not reasons:
                continue
        changes.append({'Fuvarszám': key, 'Változás': status, 'Ok': ' · '.join(reasons),
                        'Módosult forrásmezők': ', '.join(fields),
                        'Korábbi hónap': old['month'] if old else '', 'Új hónap': new['month'] if new else '',
                        'Korábbi fuvarszervező': old['contact'] if old else '',
                        'Új fuvarszervező': new['contact'] if new else '',
                        'Korábbi jutalék a vizsgált hónapokra (Ft)': a,
                        'Új jutalék a vizsgált hónapokra (Ft)': b, 'Eltérés (Ft)': b - a})
    monthly = [{'Hónap': m, 'Fuvarszervező': c, 'Korábban (Ft)': v[0], 'Most (Ft)': v[1],
                'Eltérés (Ft)': v[1] - v[0]} for (m, c), v in sorted(effects.items())]
    return {'changes': changes, 'details': detail, 'monthly': monthly,
            'months': sorted(periods), 'difference': sum(r['Eltérés (Ft)'] for r in monthly)}


class DriveArchive:
    """Small authenticated REST client; credentials never enter archives or the repo."""
    def __init__(self, config, service_account_info):
        folder = str(config.get('folder_id', '')).strip()
        if not re.fullmatch(r'[A-Za-z0-9_-]+', folder):
            raise ArchiveError('Érvényes Drive folder_id szükséges.')
        try:
            from google.oauth2.service_account import Credentials
            from google.auth.transport.requests import AuthorizedSession
        except ImportError as exc:
            raise ArchiveError('Hiányzik a Drive-hitelesítés egyik programcsomagja. '
                               'A GitHub requirements.txt fájljában legyen google-auth>=2.40,<3 és requests>=2.32,<3, '
                               'majd indítsd újra az alkalmazást a Streamlit Manage app → Reboot lehetőségével.') from exc
        info = dict(service_account_info)
        missing = [field for field in ('client_email', 'token_uri', 'private_key')
                   if not isinstance(info.get(field), str) or not info[field].strip()]
        if missing:
            raise ArchiveError('A Secrets [gcp_service_account] részéből hiányzik vagy üres: ' + ', '.join(missing) + '.')
        if any('IDE_' in info[field] for field in ('client_email', 'token_uri', 'private_key')):
            raise ArchiveError('A Secrets [gcp_service_account] részében mintaérték maradt. '
                               'A client_email, token_uri és private_key mezőt a saját JSON-fájl alapján töltsd ki.')
        if not info['client_email'].strip().endswith('.iam.gserviceaccount.com'):
            raise ArchiveError('A gcp_service_account.client_email nem szolgáltatásfiók-cím. A JSON client_email értékét másold ide.')
        if info['token_uri'].strip() != 'https://oauth2.googleapis.com/token':
            raise ArchiveError('A gcp_service_account.token_uri értéke https://oauth2.googleapis.com/token legyen.')
        # JSON → TOML copying may retain one or two literal backslashes.
        # PEM base64 has no backslashes, so these replacements are unambiguous.
        key = info['private_key'].strip().replace('\\\\n', '\n').replace('\\n', '\n').replace('\r\n', '\n').strip()
        if not key.startswith('-----BEGIN PRIVATE KEY-----') or not key.endswith('-----END PRIVATE KEY-----'):
            raise ArchiveError('A gcp_service_account.private_key mezőből hiányzik a teljes BEGIN / END PRIVATE KEY keret. '
                               'A JSON teljes private_key értékét másold át, a keretsorokkal együtt.')
        if len(key.splitlines()) < 3:
            raise ArchiveError('A gcp_service_account.private_key sortörései hibásak. '
                               'A JSON-ban szereplő \\n jelek is tartozzanak a bemásolt értékhez.')
        info['private_key'] = key + '\n'
        info['client_email'] = info['client_email'].strip()
        info['token_uri'] = info['token_uri'].strip()
        try:
            credentials = Credentials.from_service_account_info(
                info, scopes=['https://www.googleapis.com/auth/drive'])
            self.http = AuthorizedSession(credentials)
        except ImportError as exc:
            raise ArchiveError('A Drive-kulcs feldolgozásához szükséges programcsomag hiányzik. '
                               'Frissítsd a requirements.txt függőségeit, majd indítsd újra az alkalmazást.') from exc
        except (ValueError, KeyError, TypeError) as exc:
            raise ArchiveError('A private_key tartalmát a Google hitelesítési könyvtára nem tudta beolvasni. '
                               'A kulcs hiányos vagy sérült lehet. A letöltött JSON teljes private_key értékét másold át újra. '
                               'A kulcsot ne küldd el üzenetben.') from exc
        self.folder_id = folder

    def _request(self, method, path, **kwargs):
        import requests
        from google.auth.exceptions import GoogleAuthError
        try:
            response = self.http.request(method, 'https://www.googleapis.com/' + path, timeout=90, **kwargs)
        except (requests.RequestException, GoogleAuthError) as exc:
            raise ArchiveError('A Drive nem érhető el. A mentés / beolvasás nem igazolható; frissítsd az archívumlistát.') from exc
        if not response.ok:
            raise ArchiveError(f'Drive-hiba ({response.status_code}). Ellenőrizd az API-t és a mappa jogosultságát. '
                               'Szolgáltatásfióknál céges megosztott meghajtó szükséges. Sikeres mentést nem igazoltunk.')
        return response

    def check_folder(self):
        data = self._request('GET', 'drive/v3/files/' + self.folder_id,
                             params={'fields': 'id,name,mimeType,driveId,capabilities(canAddChildren)',
                                     'supportsAllDrives': 'true'}).json()
        if data.get('mimeType') != 'application/vnd.google-apps.folder' or not data.get('driveId'):
            raise ArchiveError('A folder_id egy céges megosztott meghajtón belüli mappára mutasson, ne Saját meghajtó-mappára.')
        if not data.get('capabilities', {}).get('canAddChildren'):
            raise ArchiveError('A szolgáltatásfióknak nincs fájlhozzáadási joga a mappához.')
        return data

    def list_snapshots(self):
        files, token = [], None
        while True:
            params = {'q': f"'{self.folder_id}' in parents and trashed = false and "
                               "appProperties has { key='babolna_archive' and value='v1' }",
                      'fields': 'nextPageToken,files(id,name,createdTime,appProperties,size)',
                      'pageSize': 100, 'supportsAllDrives': 'true', 'includeItemsFromAllDrives': 'true',
                      'orderBy': 'createdTime desc'}
            if token:
                params['pageToken'] = token
            data = self._request('GET', 'drive/v3/files', params=params).json()
            files.extend(data.get('files', []))
            token = data.get('nextPageToken')
            if not token:
                return files

    def load(self, file_id):
        # Only allow files actually listed in the configured archive folder.
        meta = self._request('GET', 'drive/v3/files/' + file_id,
                             params={'fields': 'parents,size,appProperties', 'supportsAllDrives': 'true'}).json()
        if self.folder_id not in meta.get('parents', []) or meta.get('appProperties', {}).get('babolna_archive') != 'v1':
            raise ArchiveError('A fájl nem a beállított jutalékarchívum része.')
        if int(meta.get('size', MAX_ARCHIVE + 1)) > MAX_ARCHIVE:
            raise ArchiveError('Az archívum túl nagy.')
        response = self._request('GET', 'drive/v3/files/' + file_id,
                                 params={'alt': 'media', 'supportsAllDrives': 'true'})
        return decode_snapshot(response.content)

    def save(self, snapshot, raw):
        self.check_folder()
        for existing in self.list_snapshots():
            if existing.get('appProperties', {}).get('snapshot_id') == snapshot['id']:
                saved, saved_raw = self.load(existing['id'])
                if saved != snapshot or saved_raw != raw:
                    raise ArchiveError('Az állapotazonosító már eltérő tartalommal létezik. Frissítsd az archívumlistát.')
                return existing
        data = encode_snapshot(snapshot, raw)
        archive_id = snapshot['id']
        boundary = 'babolna_' + uuid.uuid4().hex
        timestamp = snapshot['created'][:19].replace(':', '').replace('-', '').replace('T', '_')
        name = f'Jutalek_archivum_{timestamp}_{archive_id[:8]}.zip'
        metadata = {'name': name, 'parents': [self.folder_id], 'mimeType': 'application/zip',
                    'appProperties': {'babolna_archive': 'v1', 'snapshot_id': archive_id,
                                      'created': snapshot['created']}}
        body = (f'--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'.encode() +
                json.dumps(metadata).encode() +
                f'\r\n--{boundary}\r\nContent-Type: application/zip\r\n\r\n'.encode() + data +
                f'\r\n--{boundary}--\r\n'.encode())
        result = self._request('POST', 'upload/drive/v3/files',
                               params={'uploadType': 'multipart', 'supportsAllDrives': 'true', 'fields': 'id,name,size'},
                               headers={'Content-Type': f'multipart/related; boundary={boundary}'}, data=body).json()
        # Read-back is deliberately required before a successful archive message.
        saved, saved_raw = self.load(result['id'])
        if saved != snapshot or saved_raw != raw:
            raise ArchiveError('A Drive-mentés visszaellenőrzése sikertelen. Ne ments újra automatikusan; frissítsd a listát.')
        return result
