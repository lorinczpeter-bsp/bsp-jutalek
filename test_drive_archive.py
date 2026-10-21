import copy
import io
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import Mock, patch

import openpyxl

from core import DATE, MARGIN, Store
from drive_archive import (ArchiveError, DriveRequestError, DriveArchive, build_snapshot, compare_snapshots,
                           decode_snapshot, encode_snapshot)
from settlement import import_payments, reassign
from test_core import HEADERS, fx, rec


def source(*records, headers=HEADERS):
    w = openpyxl.Workbook(); s = w.active; s.append(headers)
    for r in records:
        s.append([r.original.get(h) for h in headers])
    data = io.BytesIO(); w.save(data)
    return data.getvalue()


class ArchiveTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.a = Store(Path(self.tmp.name) / 'a.db')
        self.b = Store(Path(self.tmp.name) / 'b.db')

    def tearDown(self):
        self.tmp.cleanup()

    def snapshot(self, store, raw, months=('2026-07',), provider=fx):
        return build_snapshot(raw, 'fuvarlista.xlsx', store, months, rate_provider=provider)

    def test_roundtrip_preserves_source_exact_rate_and_settings_not_payment(self):
        raw = source(rec())
        reassign(self.a, rec(), '2026-07')
        import_payments(self.a, [{'freight': '1-1', 'contact': 'Peti', 'amount': '4000'}],
                        'paid.xlsx', 'hash', '2026-08-01')
        rate = lambda m: dict(fx(m), rate='365.0531818181818181818')
        old = self.snapshot(self.a, raw, provider=rate)
        result, restored = decode_snapshot(encode_snapshot(old, raw))
        self.assertEqual(result, old); self.assertEqual(restored, raw)
        self.assertEqual(result['rates']['2026-07']['rate'], '365.0531818181818181818')
        self.assertEqual(result['records']['1-1']['calculation']['commission'], 3651)
        self.assertNotIn('payments', result['settings'])
        self.assertNotIn('private_key', json.dumps(result))

    def test_changed_amount_rate_and_positive_delta(self):
        old = self.snapshot(self.a, source(rec()))
        new = self.snapshot(self.b, source(rec(margin='150')), provider=lambda m: dict(fx(m), rate='450'))
        result = compare_snapshots(old, new)
        self.assertEqual(result['difference'], 2750)
        self.assertIn('Forrásadat', result['changes'][0]['Ok'])
        self.assertIn('Árfolyam', result['changes'][0]['Ok'])
        self.assertEqual(result['monthly'][0]['Most (Ft)'], 6750)

    def test_mixed_currency_uses_exact_same_business_calculation(self):
        r = rec(margin='9999', currency='EUR')
        r.original['Díj'] = '250'; r.original['Díj pénzneme'] = 'EUR'
        r.original['Alvállalkozói díj'] = '80000'; r.original['Alvállalkozói díj pénzneme'] = 'HUF'
        old = self.snapshot(self.a, source(r))
        self.assertEqual(old['records']['1-1']['calculation']['margin_huf'], '20000')
        self.assertEqual(old['records']['1-1']['calculation']['commission'], 2000)

    def test_negative_does_not_reduce_totals_and_payment_has_no_effect(self):
        raw = source(rec(), rec('2-1', margin='-100'))
        old = self.snapshot(self.a, raw)
        import_payments(self.b, [{'freight': '1-1', 'contact': 'Peti', 'amount': '4000'}],
                        'paid.xlsx', 'x', '2026-08-01')
        new = self.snapshot(self.b, raw)
        result = compare_snapshots(old, new)
        self.assertEqual(result['changes'], [])
        self.assertEqual(result['monthly'][0]['Most (Ft)'], 4000)

    def test_added_missing_and_unchanged_are_not_conflated(self):
        old = self.snapshot(self.a, source(rec(), rec('2-1', margin='200')))
        new = self.snapshot(self.b, source(rec(), rec('3-1', margin='300')))
        result = compare_snapshots(old, new)
        self.assertEqual({r['Fuvarszám'] for r in result['changes']}, {'2-1', '3-1'})
        self.assertEqual(result['changes'][0]['Változás'], 'Hiányzik az új forrásból')
        self.assertEqual(result['difference'], 4000)

    def test_reassignment_effect_balances_between_months(self):
        raw = source(rec())
        old = self.snapshot(self.a, raw, ('2026-07', '2026-08'))
        reassign(self.b, rec(), '2026-08')
        new = self.snapshot(self.b, raw, ('2026-07', '2026-08'))
        result = compare_snapshots(old, new)
        self.assertEqual(result['difference'], 0)
        self.assertEqual([r['Eltérés (Ft)'] for r in result['monthly']], [-4000, 4000])
        self.assertIn('Hónapbesorolás', result['changes'][0]['Ok'])

    def test_organizer_move_is_visible_even_when_grand_total_unchanged(self):
        old = self.snapshot(self.a, source(rec()))
        new = self.snapshot(self.b, source(rec(owner='Anna')))
        result = compare_snapshots(old, new)
        self.assertEqual(result['difference'], 0)
        self.assertEqual(result['changes'][0]['Új fuvarszervező'], 'Anna')
        self.assertEqual({r['Fuvarszervező']: r['Eltérés (Ft)'] for r in result['monthly']},
                         {'Anna': 4000, 'Peti': -4000})

    def test_cross_month_invalid_import_and_missing_fx_block_checked_save(self):
        for raw in (source(rec(), rec()), source(rec(dates='2026-07-01, 2026-08-01'))):
            with self.assertRaises(ArchiveError):
                self.snapshot(self.a, raw)
        with self.assertRaises(ArchiveError):
            self.snapshot(self.a, source(rec()), provider=Mock(side_effect=ValueError('offline')))

    def test_normalized_numeric_format_date_order_and_reordering_do_not_make_changes(self):
        r = rec(margin='100.00', dates='2026-07-03, 2026-07-01')
        old = self.snapshot(self.a, source(r))
        r = rec(margin='100', dates='2026-07-01, 2026-07-03')
        new = self.snapshot(self.b, source(r, headers=list(reversed(HEADERS))))
        self.assertEqual(compare_snapshots(old, new)['changes'], [])

    def test_out_of_scope_months_are_not_compared_as_zero(self):
        old = self.snapshot(self.a, source(rec()), ('2026-07',))
        new = self.snapshot(self.b, source(rec()), ('2026-08',))
        with self.assertRaises(ArchiveError):
            compare_snapshots(old, new)

    def test_changed_source_date_outside_scope_with_unresolved_month_is_not_fake_removal(self):
        old = self.snapshot(self.a, source(rec()))
        new = self.snapshot(self.b, source(rec(dates='2026-08-01, 2026-09-01')))
        with self.assertRaises(ArchiveError):
            compare_snapshots(old, new)

    def test_changed_calculation_version_is_explicit_even_if_no_money_change(self):
        old = self.snapshot(self.a, source(rec()))
        new = copy.deepcopy(old); new['calculation_version'] = 'future'
        self.assertIn('Számítási szabály verziója', compare_snapshots(old, new)['changes'][0]['Ok'])

    def test_corruption_wrong_source_duplicate_entries_and_tampered_rows_are_rejected(self):
        raw = source(rec()); s = self.snapshot(self.a, raw)
        with self.assertRaises(ArchiveError):
            encode_snapshot(s, b'wrong')
        for changed in ('digest', 'records', 'missing'):
            broken = copy.deepcopy(s)
            if changed == 'digest': broken['source_sha256'] = 'invalid'
            if changed == 'records': broken['records']['1-1']['original'][MARGIN] = '666'
            if changed == 'missing': broken['records'] = {}
            out = io.BytesIO()
            with zipfile.ZipFile(out, 'w') as z:
                z.writestr('source.xlsx', raw); z.writestr('snapshot.json', json.dumps(broken))
            with self.assertRaises(ArchiveError):decode_snapshot(out.getvalue())
        with self.assertRaises(ArchiveError): decode_snapshot(b'not a ZIP')


class DriveTransportTests(unittest.TestCase):
    @patch('drive_archive.time.sleep')
    def test_complete_save_survives_500_on_metadata_and_binary_readback(self, sleep):
        with tempfile.TemporaryDirectory() as d:
            raw = source(rec(currency='HUF'))
            s = build_snapshot(raw, 'source.xlsx', Store(Path(d) / 'db'), ['2026-07'])
            client = self.client(); client.check_folder = Mock(); client.list_snapshots = Mock(return_value=[])
            created = Mock(ok=True)
            created.json.return_value = {'id': 'saved', 'name': 'archive.zip'}
            metadata = Mock(ok=True)
            metadata.json.return_value = {'parents': ['folder'], 'size': 100,
                                          'appProperties': {'babolna_archive': 'v1'}}
            downloaded = Mock(ok=True, content=encode_snapshot(s, raw))
            client.http.request.side_effect = [created, Mock(ok=False, status_code=500), metadata,
                                                Mock(ok=False, status_code=500), Mock(ok=False, status_code=503), downloaded]
            self.assertEqual(client.save(s, raw), {'id': 'saved', 'name': 'archive.zip'})
            self.assertEqual([call.args[0] for call in client.http.request.call_args_list],
                             ['POST', 'GET', 'GET', 'GET', 'GET', 'GET'])
            self.assertEqual([call.args[0] for call in sleep.call_args_list], [1, 1, 2])

    @patch('drive_archive.time.sleep')
    def test_transient_reads_retry_and_recover_without_false_permission_message(self, sleep):
        import requests
        for failures in ([Mock(ok=False, status_code=500), Mock(ok=False, status_code=503)],
                         [requests.Timeout('private details'), requests.ConnectionError('private details')]):
            client = self.client(); success = Mock(ok=True)
            client.http.request.side_effect = list(failures) + [success]
            self.assertIs(client._request('GET', 'drive/v3/files'), success)
            self.assertEqual(client.http.request.call_count, 3)
        self.assertEqual([c.args[0] for c in sleep.call_args_list], [1, 2, 1, 2])

    @patch('drive_archive.time.sleep')
    def test_persistent_server_error_is_distinct_from_permission_and_is_bounded(self, sleep):
        client = self.client()
        client.http.request.return_value = Mock(ok=False, status_code=500)
        with self.assertRaises(DriveRequestError) as caught: client._request('GET', 'drive/v3/files')
        self.assertEqual(client.http.request.call_count, 3)
        self.assertTrue(caught.exception.transient)
        self.assertIn('Átmeneti Drive-hiba', str(caught.exception))
        self.assertIn('archívumlista beolvasása', str(caught.exception))
        self.assertNotIn('Ellenőrizd az API-t', str(caught.exception))

    @patch('drive_archive.time.sleep')
    def test_rate_limit_can_retry_but_permission_denial_does_not(self, sleep):
        for status, reason, expected in [(429, '', 3), (403, 'rateLimitExceeded', 3),
                                         (403, 'userRateLimitExceeded', 3), (403, 'insufficientFilePermissions', 1),
                                         (404, '', 1)]:
            client = self.client()
            response = Mock(ok=False, status_code=status)
            response.json.return_value = {'error': {'errors': [{'reason': reason}]}}
            client.http.request.return_value = response
            with self.assertRaises(DriveRequestError):client._request('GET', 'drive/v3/files')
            self.assertEqual(client.http.request.call_count, expected)

    @patch('drive_archive.time.sleep')
    def test_upload_request_itself_is_never_blindly_retried(self, sleep):
        client = self.client()
        client.http.request.return_value = Mock(ok=False, status_code=500)
        with self.assertRaises(DriveRequestError):client._request('POST', 'upload/drive/v3/files')
        self.assertEqual(client.http.request.call_count, 1)
        sleep.assert_not_called()

    @patch('drive_archive.time.sleep')
    def test_uncertain_upload_finds_saved_snapshot_and_verifies_without_second_post(self, sleep):
        with tempfile.TemporaryDirectory() as d:
            raw = source(rec(currency='HUF'))
            s = build_snapshot(raw, 'source.xlsx', Store(Path(d) / 'db'), ['2026-07'])
            client = self.client(); client.check_folder = Mock()
            saved = {'id': 'saved', 'name': 'archive.zip', 'appProperties': {'snapshot_id': s['id']}}
            client.list_snapshots = Mock(side_effect=[[], [], [saved]])
            client._request = Mock(side_effect=DriveRequestError('temporary', status=500, transient=True))
            client.load = Mock(return_value=(s, raw))
            self.assertEqual(client.save(s, raw), saved)
            self.assertEqual(client._request.call_count, 1)
            client.load.assert_called_once_with('saved')
            # Recovery never treats a mismatching payload as a successful save.
            client.list_snapshots = Mock(side_effect=[[], [saved]])
            client.load.return_value = (dict(s, note='different'), raw)
            with self.assertRaisesRegex(ArchiveError, 'visszaellenőrzése'):client.save(s, raw)

    @patch('drive_archive.time.sleep')
    def test_uncertain_upload_without_match_stays_unconfirmed_and_never_uploads_twice(self, sleep):
        with tempfile.TemporaryDirectory() as d:
            raw = source(rec(currency='HUF'))
            s = build_snapshot(raw, 'source.xlsx', Store(Path(d) / 'db'), ['2026-07'])
            client = self.client(); client.check_folder = Mock(); client.list_snapshots = Mock(return_value=[])
            client._request = Mock(side_effect=DriveRequestError('temporary', status=500, transient=True))
            client.load = Mock()
            with self.assertRaisesRegex(DriveRequestError, 'Nem indítottunk ismételt feltöltést'):client.save(s, raw)
            self.assertEqual(client._request.call_count, 1)
            self.assertEqual(client.list_snapshots.call_count, 4)
            client.load.assert_not_called()

    def test_confirmed_upload_with_unavailable_readback_is_not_called_failed_upload(self):
        with tempfile.TemporaryDirectory() as d:
            raw = source(rec(currency='HUF'))
            s = build_snapshot(raw, 'source.xlsx', Store(Path(d) / 'db'), ['2026-07'])
            client = self.client(); client.check_folder = Mock(); client.list_snapshots = Mock(return_value=[])
            client._request = Mock(return_value=Mock(**{'json.return_value': {'id': 'saved', 'name': 'archive.zip'}}))
            client.load = Mock(side_effect=DriveRequestError('Átmeneti Drive-hiba (500)', status=500, transient=True))
            with self.assertRaisesRegex(DriveRequestError, 'fájlja létrejött') as caught:client.save(s, raw)
            self.assertTrue(caught.exception.transient)
            self.assertEqual(client._request.call_count, 1)

    def test_credential_diagnostics_and_newline_normalization_without_secret_leak(self):
        from google.oauth2.service_account import Credentials
        from google.auth.transport.requests import AuthorizedSession
        info = {'client_email': 'example@project.iam.gserviceaccount.com',
                'token_uri': 'https://oauth2.googleapis.com/token',
                'private_key': '-----BEGIN PRIVATE KEY-----\nTEST_SECRET\n-----END PRIVATE KEY-----\n'}
        for key in (info['private_key'], info['private_key'].replace('\n', '\\n'),
                    info['private_key'].replace('\n', '\\\\n')):
            with patch.object(Credentials, 'from_service_account_info', return_value=Mock()) as factory:
                DriveArchive({'folder_id': 'folder'}, dict(info, private_key=key))
                self.assertEqual(factory.call_args.args[0]['private_key'], info['private_key'])
        with self.assertRaisesRegex(ArchiveError, 'hiányzik vagy üres'):
            DriveArchive({'folder_id': 'folder'}, dict(info, token_uri=''))
        with self.assertRaisesRegex(ArchiveError, 'mintaérték'):
            DriveArchive({'folder_id': 'folder'}, dict(info, private_key='IDE_A_JSON_PRIVATE_KEY_ERTEKE'))
        with self.assertRaisesRegex(ArchiveError, 'BEGIN / END'):
            DriveArchive({'folder_id': 'folder'}, dict(info, private_key='secret'))
        with patch.object(Credentials, 'from_service_account_info', side_effect=ValueError('TEST_SECRET')):
            with self.assertRaises(ArchiveError) as caught:
                DriveArchive({'folder_id': 'folder'}, info)
            self.assertNotIn('TEST_SECRET', str(caught.exception))

    def client(self):
        client = object.__new__(DriveArchive)
        client.folder_id = 'folder'; client.http = Mock()
        return client

    def test_service_account_cannot_write_shared_my_drive_folder(self):
        client = self.client()
        client._request = Mock(return_value=Mock(**{'json.return_value': {
            'mimeType': 'application/vnd.google-apps.folder', 'capabilities': {'canAddChildren': True}}}))
        with self.assertRaisesRegex(ArchiveError, 'Saját meghajtó'):client.check_folder()

    def test_pagination_and_shared_drive_flags(self):
        client = self.client()
        client._request = Mock(side_effect=[Mock(**{'json.return_value': {'files': [{'id': 'a'}], 'nextPageToken': 'n'}}),
                                           Mock(**{'json.return_value': {'files': [{'id': 'b'}]}})])
        self.assertEqual([f['id'] for f in client.list_snapshots()], ['a', 'b'])
        self.assertEqual(client._request.call_args.kwargs['params']['pageToken'], 'n')
        self.assertEqual(client._request.call_args.kwargs['params']['supportsAllDrives'], 'true')

    def test_save_readback_and_idempotent_retry_no_updates_or_delete(self):
        with tempfile.TemporaryDirectory() as d:
            raw = source(rec(currency='HUF'))
            s = build_snapshot(raw, 'source.xlsx', Store(Path(d) / 'db'), ['2026-07'])
            client = self.client(); client.check_folder = Mock(); client.list_snapshots = Mock(return_value=[])
            client._request = Mock(return_value=Mock(**{'json.return_value': {'id': 'saved', 'name': 'arch.zip'}}))
            client.load = Mock(return_value=(s, raw))
            self.assertEqual(client.save(s, raw)['id'], 'saved')
            self.assertEqual(client._request.call_args.args[0], 'POST')
            client._request.reset_mock()
            client.list_snapshots.return_value = [{'id': 'saved', 'name': 'arch.zip', 'appProperties': {'snapshot_id': s['id']}}]
            self.assertEqual(client.save(s, raw)['id'], 'saved')
            client._request.assert_not_called()
            client.list_snapshots.return_value = []
            client.load.return_value = (dict(s, note='tampered'), raw)
            with self.assertRaisesRegex(ArchiveError, 'visszaellenőrzése'):client.save(s, raw)

    def test_foreign_folder_file_cannot_be_downloaded(self):
        client = self.client()
        client._request = Mock(return_value=Mock(**{'json.return_value': {'parents': ['other']}}))
        with self.assertRaises(ArchiveError):client.load('foreign')
        self.assertEqual(client._request.call_count, 1)

    def test_api_failure_is_not_success_and_does_not_expose_error_body(self):
        client = self.client()
        client.http.request.return_value = Mock(ok=False, status_code=403, text='sensitive error details')
        with self.assertRaises(ArchiveError) as caught:client._request('GET', 'drive/v3/files')
        self.assertNotIn('sensitive', str(caught.exception))

    def test_auth_failure_is_a_controlled_archive_error(self):
        try:
            from google.auth.exceptions import RefreshError
        except ImportError:
            self.skipTest('google-auth is required for Drive transport tests')
        client = self.client()
        client.http.request.side_effect = RefreshError('internal credential details')
        with self.assertRaises(ArchiveError) as caught: client._request('GET', 'drive/v3/files')
        self.assertNotIn('internal credential', str(caught.exception))


if __name__ == '__main__':unittest.main()
