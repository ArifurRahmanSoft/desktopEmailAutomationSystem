import json
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from openpyxl import load_workbook


TRACKING_COLUMNS = ("OpenCount", "ClickCount", "FirstOpen", "LastOpen", "FirstClick", "LastClick")


class TrackingSynchronizationService:
    def __init__(self, base_url, excel_path, log_folder, http_get=None):
        self.base_url = base_url.rstrip("/")
        self.excel_path = Path(excel_path)
        self.log_folder = Path(log_folder)
        self.http_get = http_get or self._http_get

    @staticmethod
    def _http_get(url):
        request = Request(url, headers={"Accept": "application/json", "User-Agent": "EmailAutomation/2"})
        with urlopen(request, timeout=60) as response:
            return json.loads(response.read().decode("utf-8"))

    @staticmethod
    def _records(payload):
        if isinstance(payload, list):
            return payload
        if isinstance(payload, dict):
            for key in ("records", "data", "results", "items"):
                if isinstance(payload.get(key), list):
                    return payload[key]
        raise ValueError("Tracking API returned an unsupported JSON format")

    @staticmethod
    def _normalized(record):
        return {str(key).replace("_", "").casefold(): value for key, value in record.items()}

    def _log(self, event, **values):
        self.log_folder.mkdir(parents=True, exist_ok=True)
        path = self.log_folder / f"tracking-sync-{datetime.now():%Y-%m-%d}.log"
        entry = {"timestamp": datetime.now(timezone.utc).isoformat(), "event": event, **values}
        with path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(entry, ensure_ascii=False, default=str) + "\n")

    def sync(self, last_sync_time=""):
        started = time.monotonic()
        synchronization_time = datetime.now(timezone.utc).isoformat()
        query = urlencode({"updated_after": last_sync_time}) if last_sync_time else ""
        api_url = f"{self.base_url}/api/tracking/sync" + (f"?{query}" if query else "")
        self._log("Synchronization Started", api_url=api_url)
        try:
            payload = self.http_get(api_url)
            records = self._records(payload)
            if not self.excel_path.exists():
                raise FileNotFoundError(f"Excel file not found: {self.excel_path}")
            workbook = load_workbook(self.excel_path)
            worksheet = workbook.active
            headers = {str(cell.value).strip().casefold(): index for index, cell in enumerate(worksheet[1], 1) if cell.value is not None}
            tracking_column = headers.get("trackingid")
            if not tracking_column:
                workbook.close()
                raise ValueError("TrackingId column not found in mail_list.xlsx")
            column_indexes = {}
            columns_added = False
            for name in TRACKING_COLUMNS:
                key = name.casefold()
                if key not in headers:
                    index = worksheet.max_column + 1
                    worksheet.cell(1, index, name)
                    headers[key] = index
                    columns_added = True
                column_indexes[name] = headers[key]
            row_by_tracking_id = {}
            for row in range(2, worksheet.max_row + 1):
                value = worksheet.cell(row, tracking_column).value
                if value not in (None, ""):
                    row_by_tracking_id[str(value).strip().casefold()] = row
            updated_rows = set()
            for record in records:
                if not isinstance(record, dict):
                    continue
                values = self._normalized(record)
                tracking_id = values.get("trackingid")
                row = row_by_tracking_id.get(str(tracking_id or "").strip().casefold())
                if not row:
                    continue
                for name in TRACKING_COLUMNS:
                    normalized_name = name.casefold()
                    if normalized_name in values:
                        worksheet.cell(row, column_indexes[name], values[normalized_name])
                updated_rows.add(row)
            if updated_rows or columns_added:
                workbook.save(self.excel_path)
            workbook.close()
            elapsed = time.monotonic() - started
            result = {
                "records_downloaded": len(records),
                "rows_updated": len(updated_rows),
                "execution_time": elapsed,
                "last_sync_time": synchronization_time,
                "api_url": api_url,
            }
            self._log("Synchronization Finished", api_url=api_url, returned_record_count=len(records), updated_excel_row_count=len(updated_rows), execution_time_seconds=round(elapsed, 3))
            return result
        except Exception as exc:
            elapsed = time.monotonic() - started
            self._log("Synchronization Error", api_url=api_url, error=str(exc), execution_time_seconds=round(elapsed, 3))
            raise


def synchronization_is_due(enabled, interval_hours, last_sync_time, now=None):
    if not enabled:
        return False
    if not last_sync_time:
        return True
    now = now or datetime.now(timezone.utc)
    try:
        previous = datetime.fromisoformat(last_sync_time.replace("Z", "+00:00"))
        if previous.tzinfo is None:
            previous = previous.replace(tzinfo=timezone.utc)
        return (now - previous).total_seconds() >= int(interval_hours) * 3600
    except (ValueError, TypeError):
        return True
