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
