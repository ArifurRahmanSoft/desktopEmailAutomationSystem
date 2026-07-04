import http.client
import json
import mimetypes
import secrets
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urlsplit
from urllib.request import Request, urlopen


class AttachmentLibraryService:
    def __init__(self, base_url, log_folder, connection_factory=None, request_opener=None):
        self.base_url = base_url.rstrip("/")
        self.log_folder = Path(log_folder)
        self.connection_factory = connection_factory
        self.request_opener = request_opener or urlopen

    def _log(self, event, **values):
        self.log_folder.mkdir(parents=True, exist_ok=True)
        path = self.log_folder / f"attachment-{datetime.now():%Y-%m-%d}.log"
        with path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps({"timestamp": datetime.now().isoformat(), "event": event, **values}, ensure_ascii=False, default=str) + "\n")

    def _request_json(self, method, path, payload=None):
        url = self.base_url + path
        data = json.dumps(payload).encode("utf-8") if payload is not None else None
        headers = {"Accept": "application/json", "User-Agent": "EmailAutomation/2"}
        if data is not None:
            headers["Content-Type"] = "application/json"
        request = Request(url, data=data, method=method, headers=headers)
        try:
            with self.request_opener(request, timeout=60) as response:
                raw = response.read().decode("utf-8")
                return json.loads(raw) if raw else {}
        except Exception as exc:
            self._log("Server Error", operation=method, url=url, error=str(exc))
            raise RuntimeError(f"Attachment server is unavailable: {exc}") from exc

    @staticmethod
    def normalize_list(payload):
        records = payload
        if isinstance(payload, dict):
            records = next((payload[key] for key in ("attachments", "records", "data", "items") if isinstance(payload.get(key), list)), None)
        if not isinstance(records, list):
            raise ValueError("Attachment server returned an unsupported response.")
        result = []
        for item in records:
            if not isinstance(item, dict): continue
            attachment_id = item.get("attachment_id") or item.get("id") or item.get("_id")
            name = item.get("original_file_name") or item.get("file_name") or item.get("filename") or item.get("original_name") or item.get("name") or ""
            size = item.get("file_size") if item.get("file_size") is not None else item.get("size", 0)
            uploaded = item.get("upload_date") or item.get("uploaded_at") or item.get("created_at") or ""
            if attachment_id is not None: result.append({"id": str(attachment_id), "file_name": str(name), "file_size": int(size or 0), "upload_date": str(uploaded)})
        return result

    def list_attachments(self):
        try:
            result = self.normalize_list(self._request_json("GET", "/api/attachments/list"))
            self._log("Refresh", attachment_count=len(result), status="Success")
            return result
        except Exception as exc:
            if not isinstance(exc, RuntimeError): self._log("Server Error", operation="Refresh", error=str(exc))
            raise

    def upload(self, file_path, progress=None):
        file_path = Path(file_path)
        boundary = "----EmailAutomation" + secrets.token_hex(16)
        content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        preamble = (f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"{file_path.name.replace(chr(34), '')}\"\r\nContent-Type: {content_type}\r\n\r\n").encode("utf-8")
        ending = f"\r\n--{boundary}--\r\n".encode("ascii")
        file_size = file_path.stat().st_size
        parsed = urlsplit(self.base_url)
        factory = self.connection_factory or (http.client.HTTPSConnection if parsed.scheme == "https" else http.client.HTTPConnection)
        connection = factory(parsed.hostname, parsed.port, timeout=120)
        path = (parsed.path.rstrip("/") if parsed.path else "") + "/api/attachments/upload"
        try:
            connection.putrequest("POST", path)
            connection.putheader("Content-Type", f"multipart/form-data; boundary={boundary}")
            connection.putheader("Content-Length", str(len(preamble) + file_size + len(ending)))
            connection.putheader("Accept", "application/json")
            connection.endheaders()
            connection.send(preamble)
            transferred = 0
            with file_path.open("rb") as handle:
                while True:
                    chunk = handle.read(1024 * 256)
                    if not chunk: break
                    connection.send(chunk); transferred += len(chunk)
                    if progress: progress(transferred, file_size)
            connection.send(ending)
            response = connection.getresponse(); raw = response.read().decode("utf-8", errors="replace")
            if not 200 <= response.status < 300: raise RuntimeError(f"Upload failed ({response.status}): {raw}")
            result = json.loads(raw) if raw else {}
            self._log("Upload", file_name=file_path.name, file_size=file_size, status="Success")
            return result
        except Exception as exc:
            self._log("Upload", file_name=file_path.name, file_size=file_size, status="Failed", error=str(exc))
            raise RuntimeError(f"Upload failed for {file_path.name}: {exc}") from exc
        finally:
            connection.close()

    def delete(self, attachment_id):
        try:
            result = self._request_json("DELETE", f"/api/attachments/{quote(str(attachment_id), safe='')}")
            self._log("Delete", attachment_id=str(attachment_id), status="Success")
            return result
        except Exception as exc:
            if not isinstance(exc, RuntimeError): self._log("Delete", attachment_id=str(attachment_id), status="Failed", error=str(exc))
            raise

    def register_tracking_attachments(self, tracking_id, attachment_ids):
        payload = {"tracking_id": str(tracking_id), "attachment_ids": [str(value) for value in attachment_ids]}
        try:
            result = self._request_json("POST", "/api/tracking/attachments", payload)
            self._log("Register Attachment Mapping", tracking_id=str(tracking_id), attachment_ids=payload["attachment_ids"], status="Success")
            return result
        except Exception as exc:
            self._log("Register Attachment Mapping", tracking_id=str(tracking_id), attachment_ids=payload["attachment_ids"], status="Failed", error=str(exc))
            raise RuntimeError(f"Attachment mapping registration failed: {exc}") from exc
