import email
import imaplib
import json
import re
from datetime import datetime, timezone
from email import policy
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from urllib.request import Request, urlopen


REGISTER_REPLY_ENDPOINT = "/api/tracking/register-reply"
MESSAGE_ID_PATTERN = re.compile(r"<[^<>]+>")


def extract_original_message_id(message):
    in_reply_to = str(message.get("In-Reply-To") or "").strip()
    if in_reply_to:
        matches = MESSAGE_ID_PATTERN.findall(in_reply_to)
        return matches[0] if matches else in_reply_to
    references = str(message.get("References") or "").strip()
    if references:
        matches = MESSAGE_ID_PATTERN.findall(references)
        return matches[-1] if matches else references
    return ""


def reply_time_from_message(message):
    value = message.get("Date")
    if not value:
        return datetime.now(timezone.utc).isoformat()
    try:
        parsed = parsedate_to_datetime(str(value))
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc).isoformat()
    except Exception:
        return datetime.now(timezone.utc).isoformat()


class ReplyTrackingService:
    def __init__(self, base_url, log_folder=None, imap_ssl_factory=None, imap_factory=None, http_post=None):
        self.base_url = str(base_url or "").rstrip("/")
        self.log_folder = Path(log_folder) if log_folder else None
        self.imap_ssl_factory = imap_ssl_factory or imaplib.IMAP4_SSL
        self.imap_factory = imap_factory or imaplib.IMAP4
        self.http_post = http_post or self._http_post

    @staticmethod
    def _http_post(url, payload):
        data = json.dumps(payload).encode("utf-8")
        request = Request(
            url,
            data=data,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json",
                "User-Agent": "EmailAutomation/2",
            },
            method="POST",
        )
        with urlopen(request, timeout=30) as response:
            body = response.read().decode("utf-8", errors="replace")
            return {
                "status_code": getattr(response, "status", None) or response.getcode(),
                "body": body,
                "json": json.loads(body) if body else {},
            }

    def _log(self, event, **values):
        if not self.log_folder:
            return
        try:
            self.log_folder.mkdir(parents=True, exist_ok=True)
            entry = {"timestamp": datetime.now(timezone.utc).isoformat(), "event": event, **values}
            with (self.log_folder / "reply-tracking.log").open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(entry, ensure_ascii=False, default=str) + "\n")
        except Exception:
            pass

    def _connect(self, configuration):
        encryption = configuration.get("imap_encryption", "SSL/TLS")
        factory = self.imap_ssl_factory if encryption == "SSL/TLS" else self.imap_factory
        client = factory(configuration["imap_host"], int(configuration["imap_port"]))
        if encryption == "STARTTLS":
            client.starttls()
        client.login(configuration["email"], configuration["password"])
        self._log(
            "IMAP Login Success",
            email=configuration.get("email"),
            imap_host=configuration.get("imap_host"),
            imap_port=configuration.get("imap_port"),
            imap_encryption=encryption,
            success=True,
        )
        return client

    def register_reply(self, message_id, from_email, reply_time):
        url = f"{self.base_url}{REGISTER_REPLY_ENDPOINT}"
        payload = {
            "message_id": message_id,
            "from_email": from_email,
            "reply_time": reply_time,
        }
        self._log("Register Reply API Call", called=True, method="POST", url=url, payload=payload)
        self._log("API request", method="POST", url=url, payload=payload)
        response = self.http_post(url, payload)
        self._log(
            "API response",
            method="POST",
            url=url,
            http_status=response.get("status_code") if isinstance(response, dict) else None,
            response_body=response.get("body") if isinstance(response, dict) else "",
            response=response,
        )
        return response

    def check_account(self, configuration):
        if not self.base_url:
            raise ValueError("Tracking base URL is not configured.")
        self._log("Synchronization Started", flow="reply_detection", account=configuration.get("email"))
        client = self._connect(configuration)
        detected = registered = skipped = 0
        processed_reply_message_ids = set()
        try:
            status, _ = client.select("INBOX")
            if status != "OK":
                raise RuntimeError("Unable to select IMAP inbox")
            self._log("Mailbox Selected", mailbox="INBOX", status=status)
            self._log("IMAP Search Criteria", criteria=["UNSEEN"], mode="SEARCH")
            status, data = client.search(None, "UNSEEN")
            self._log("Raw IMAP SEARCH Response", mode="SEARCH", status=status, raw_response=[item.decode(errors="replace") if isinstance(item, bytes) else str(item) for item in (data or [])])
            if status != "OK":
                raise RuntimeError("Unable to search unread IMAP messages")
            message_numbers = data[0].split() if data and data[0] else []
            self._log("Unread Email Count", count=len(message_numbers))
            for message_number in message_numbers:
                try:
                    status, payload = client.fetch(message_number, "(RFC822)")
                    if status != "OK":
                        skipped += 1
                        continue
                    raw = next((item[1] for item in payload if isinstance(item, tuple) and len(item) > 1), None)
                    if not raw:
                        skipped += 1
                        continue
                    message = email.message_from_bytes(raw, policy=policy.default)
                    message_subject = str(message.get("Subject") or "")
                    message_sender = str(message.get("From") or "")
                    message_id = str(message.get("Message-ID") or "")
                    in_reply_to = str(message.get("In-Reply-To") or "")
                    references = str(message.get("References") or "")
                    original_message_id = extract_original_message_id(message)
                    self._log(
                        "Unread Email Inspected",
                        uid=message_number.decode(errors="replace") if isinstance(message_number, bytes) else str(message_number),
                        subject=message_subject,
                        from_email=message_sender,
                        content_type=message.get_content_type(),
                        message_id=message_id,
                        in_reply_to=in_reply_to,
                        references=references,
                        classification="Reply" if original_message_id else "Neither",
                        extracted_original_message_id=original_message_id,
                        extraction_source="In-Reply-To" if in_reply_to and original_message_id else ("References" if references and original_message_id else ""),
                    )
                    if not original_message_id:
                        skipped += 1
                        self._log(
                            "Warning",
                            warning="Original Message-ID not found",
                            reason="No Message-ID found in In-Reply-To or References headers.",
                            uid=message_number.decode(errors="replace") if isinstance(message_number, bytes) else str(message_number),
                            message_id=message_id,
                            in_reply_to=in_reply_to,
                            references=references,
                            api_called=False,
                        )
                        continue
                    reply_message_id = str(message.get("Message-ID") or message_number)
                    if reply_message_id in processed_reply_message_ids:
                        client.store(message_number, "+FLAGS", "\\Seen")
                        continue
                    processed_reply_message_ids.add(reply_message_id)
                    from_email = parseaddr(str(message.get("From") or ""))[1]
                    reply_time = reply_time_from_message(message)
                    detected += 1
                    self._log(
                        "Reply detected",
                        account=configuration.get("email"),
                        original_message_id=original_message_id,
                        reply_message_id=reply_message_id,
                        from_email=from_email,
                        reply_time=reply_time,
                    )
                    self.register_reply(original_message_id, from_email, reply_time)
                    registered += 1
                    client.store(message_number, "+FLAGS", "\\Seen")
                except Exception as exc:
                    skipped += 1
                    self._log("Errors", account=configuration.get("email"), message_number=str(message_number), error=str(exc))
            return {"detected": detected, "registered": registered, "skipped": skipped}
        finally:
            try:
                client.logout()
            except Exception:
                pass
