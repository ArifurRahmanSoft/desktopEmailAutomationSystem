import hashlib
import sys
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from tracking_sync_service import TRACKING_COLUMNS, TrackingSynchronizationService, synchronization_is_due


class TrackingSynchronizationTests(unittest.TestCase):
    def workbook(self, folder):
        path = Path(folder) / "mail_list.xlsx"
        wb = Workbook(); ws = wb.active
        ws.append(["Email", "Status", "TrackingId", "Unrelated"])
        ws.append(["one@example.com", "Sent", "track-1", "keep-one"])
        ws.append(["two@example.com", "Sent", "track-2", "keep-two"])
        wb.save(path)
        return path

    def test_first_sync_downloads_all_and_updates_only_tracking_columns(self):
        with tempfile.TemporaryDirectory() as temp:
            path = self.workbook(temp); urls = []
            payload = [
                {"trackingId": "track-1", "openCount": 3, "clickCount": 2, "firstOpen": "2026-01-01", "lastOpen": "2026-01-02", "firstClick": "2026-01-03", "lastClick": "2026-01-04"},
                {"tracking_id": "missing-row", "open_count": 99},
            ]
            service = TrackingSynchronizationService("https://server.test", path, Path(temp) / "logs", lambda url: urls.append(url) or payload)
            result = service.sync("")
            self.assertEqual(urls, ["https://server.test/api/tracking/sync"])
            self.assertEqual(result["records_downloaded"], 2)
            self.assertEqual(result["rows_updated"], 1)
            self.assertEqual(result["last_sync_time"], "")
            wb = load_workbook(path, data_only=True); ws = wb.active
            headers = {str(c.value): i for i, c in enumerate(ws[1], 1)}
            self.assertEqual(ws.cell(2, headers["OpenCount"]).value, 3)
            self.assertEqual(ws.cell(2, headers["ClickCount"]).value, 2)
            self.assertEqual(ws.cell(2, headers["Unrelated"]).value, "keep-one")
            self.assertEqual(ws.cell(2, headers["Status"]).value, "Sent")
            self.assertTrue(all(name in headers for name in TRACKING_COLUMNS))
            wb.close()

    def test_incremental_sync_uses_updated_after(self):
        with tempfile.TemporaryDirectory() as temp:
            path = self.workbook(temp); urls = []
            service = TrackingSynchronizationService("https://server.test", path, Path(temp) / "logs", lambda url: urls.append(url) or [])
            service.sync("2026-06-30T10:20:30+00:00")
            self.assertIn("updated_after=2026-06-30T10%3A20%3A30%2B00%3A00", urls[0])

    def test_download_fields_are_appended_updated_and_logged(self):
        with tempfile.TemporaryDirectory() as temp:
            path = self.workbook(temp)
            wb = load_workbook(path)
            ws = wb.active
            for name in ("OpenCount", "ClickCount", "FirstOpen", "LastOpen", "FirstClick", "LastClick"):
                ws.cell(1, ws.max_column + 1, name)
            ws.cell(2, 4, "keep-one")
            wb.save(path)
            wb.close()
            payload = [
                {
                    "tracking_id": "track-1",
                    "download_count": 4,
                    "first_download": "2026-07-04T01:02:03Z",
                    "last_download": "2026-07-04T04:05:06Z",
                },
                {"trackingId": "missing-row", "downloadCount": 2},
                {"trackingId": "track-2", "openCount": 8},
            ]
            logs = Path(temp) / "logs"
            service = TrackingSynchronizationService("https://server.test", path, logs, lambda _: payload)
            result = service.sync("")
            wb = load_workbook(path, data_only=True)
            ws = wb.active
            headers = {str(cell.value): index for index, cell in enumerate(ws[1], 1)}
            self.assertEqual(headers["DownloadCount"], headers["LastClick"] + 1)
            self.assertEqual(headers["FirstDownload"], headers["LastClick"] + 2)
            self.assertEqual(headers["LastDownload"], headers["LastClick"] + 3)
            self.assertEqual(ws.cell(2, headers["DownloadCount"]).value, 4)
            self.assertEqual(ws.cell(2, headers["FirstDownload"]).value, "2026-07-04T01:02:03Z")
            self.assertEqual(ws.cell(2, headers["LastDownload"]).value, "2026-07-04T04:05:06Z")
            self.assertEqual(ws.cell(2, headers["Unrelated"]).value, "keep-one")
            wb.close()
            self.assertEqual(result["download_records_received"], 2)
            self.assertEqual(result["download_rows_updated"], 1)
            content = next(logs.glob("tracking-sync-*.log")).read_text(encoding="utf-8")
            self.assertIn('"download_records_received": 2', content)
            self.assertIn('"download_rows_updated": 1', content)
            self.assertIn('"execution_time_seconds":', content)

    def test_cursor_uses_maximum_server_updated_at_and_logs_debug_values(self):
        with tempfile.TemporaryDirectory() as temp:
            path = self.workbook(temp)
            payload = [
                {"trackingId": "track-1", "openCount": 1, "updated_at": "2026-06-30T09:00:00Z"},
                {"trackingId": "track-2", "clickCount": 2, "updatedAt": "2026-06-30T12:30:00+00:00"},
                {"trackingId": "track-2", "openCount": 3},
            ]
            logs = Path(temp) / "logs"
            service = TrackingSynchronizationService("https://server.test", path, logs, lambda _: payload)
            result = service.sync("2026-06-30T08:00:00Z")
            self.assertEqual(result["last_sync_time"], "2026-06-30T12:30:00+00:00")
            content = next(logs.glob("tracking-sync-*.log")).read_text(encoding="utf-8")
            self.assertIn('"previous_last_sync": "2026-06-30T08:00:00Z"', content)
            self.assertIn('"maximum_updated_at_received": "2026-06-30T12:30:00+00:00"', content)
            self.assertIn('"new_last_sync": "2026-06-30T12:30:00+00:00"', content)
            self.assertIn('"records_downloaded": 3', content)
            self.assertIn('"rows_updated": 2', content)

    def test_zero_records_keeps_previous_cursor(self):
        with tempfile.TemporaryDirectory() as temp:
            path = self.workbook(temp)
            previous = "2026-06-30T08:00:00Z"
            service = TrackingSynchronizationService("https://server.test", path, Path(temp) / "logs", lambda _: [])
            result = service.sync(previous)
            self.assertEqual(result["last_sync_time"], previous)

    def test_full_debug_sync_ignores_cursor_and_logs_reconciliation_counts(self):
        with tempfile.TemporaryDirectory() as temp:
            path = self.workbook(temp); urls = []
            payload = [
                {"trackingId": "track-1", "openCount": 4, "updated_at": "2026-06-30T12:00:00Z"},
                {"trackingId": "not-in-excel", "clickCount": 1, "updated_at": "2026-06-30T13:00:00Z"},
            ]
            previous = "2026-06-29T08:00:00Z"; logs = Path(temp) / "logs"
            service = TrackingSynchronizationService("https://server.test", path, logs, lambda url: urls.append(url) or payload)
            result = service.sync(previous, full_synchronization_debug=True)
            self.assertEqual(urls, ["https://server.test/api/tracking/sync"])
            self.assertEqual(result["last_sync_time"], previous)
            self.assertEqual(result["tracking_ids_matched"], 1)
            self.assertEqual(result["rows_updated"], 1)
            self.assertEqual(result["tracking_ids_not_found"], 1)
            content = next(logs.glob("tracking-sync-*.log")).read_text(encoding="utf-8")
            self.assertIn('"total_records_downloaded": 2', content)
            self.assertIn('"total_tracking_ids_matched": 1', content)
            self.assertIn('"total_excel_rows_updated": 1', content)
            self.assertIn('"total_tracking_ids_not_found": 1', content)
            self.assertIn('"full_synchronization_debug": true', content)

    def test_api_failure_does_not_modify_excel(self):
        with tempfile.TemporaryDirectory() as temp:
            path = self.workbook(temp)
            before = hashlib.sha256(path.read_bytes()).hexdigest()
            def failed(_): raise RuntimeError("API unavailable")
            service = TrackingSynchronizationService("https://server.test", path, Path(temp) / "logs", failed)
            with self.assertRaisesRegex(RuntimeError, "API unavailable"): service.sync("")
            self.assertEqual(before, hashlib.sha256(path.read_bytes()).hexdigest())

    def test_schedule_due_supports_first_run_interval_and_restart_catchup(self):
        now = datetime(2026, 6, 30, 12, 0, tzinfo=timezone.utc)
        self.assertTrue(synchronization_is_due(True, 5, "", now))
        self.assertFalse(synchronization_is_due(False, 5, "", now))
        self.assertFalse(synchronization_is_due(True, 5, (now - timedelta(hours=4)).isoformat(), now))
        self.assertTrue(synchronization_is_due(True, 5, (now - timedelta(hours=6)).isoformat(), now))


if __name__ == "__main__": unittest.main()
