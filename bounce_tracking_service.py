import email
import imaplib
import json
import re
from datetime import datetime, timedelta, timezone
from email import policy
from email.header import decode_header, make_header
from pathlib import Path
from urllib.request import Request, urlopen


REGISTER_BOUNCE_ENDPOINT = "/api/tracking/register-bounce"
BOUNCE_SENDERS = ("mail delivery subsystem", "mailer-daemon", "mailer daemon", "postmaster")
BOUNCE_SUBJECT_WORDS = ("undeliver", "delivery status notification", "delivery failure", "mail delivery failed", "returned mail")
MESSAGE_ID_PATTERN = re.compile(r"<[^<>]+>")
ORIGINAL_MESSAGE_ID_HEADERS = ("In-Reply-To", "References", "Original-Message-ID", "X-Original-Message-ID")


def decode_text(value):
    try:
        return str(make_header(decode_header(value or "")))
    except Exception:
        return str(value or "")


def message_text(message):
    values = []
    for part in message.walk():
        if part.get_content_type() == "text/plain":
            try:
                values.append(part.get_content())
            except Exception:
                payload = part.get_payload(decode=True) or b""
                values.append(payload.decode(part.get_content_charset() or "utf-8", errors="replace"))
    return "\n".join(values)


def is_bounce_message(message):
    sender = decode_text(message.get("From")).casefold()
    subject = decode_text(message.get("Subject")).casefold()
    has_delivery_status = any(part.get_content_type() == "message/delivery-status" for part in message.walk())
    return has_delivery_status or any(value in sender for value in BOUNCE_SENDERS) or any(value in subject for value in BOUNCE_SUBJECT_WORDS)


def first_message_id(value, prefer_last=False):
    matches = MESSAGE_ID_PATTERN.findall(str(value or ""))
    if not matches:
        return ""
    return matches[-1 if prefer_last else 0].strip()


def message_id_from_embedded_message(part):
    payload = part.get_payload()
    nested_messages = payload if isinstance(payload, list) else []
    if not nested_messages:
        try:
            raw_payload = part.get_payload(decode=True)
        except Exception:
            raw_payload = None
        if raw_payload:
            nested_messages = [email.message_from_bytes(raw_payload, policy=policy.default)]
    for nested in nested_messages:
        value = nested.get("Message-ID")
        message_id = first_message_id(value)
        if message_id:
            return message_id
    return ""


def extract_original_message_id_details(message):
    for part in message.walk():
        if part.get_content_type() != "message/rfc822":
            continue
        message_id = message_id_from_embedded_message(part)
        if message_id:
            return message_id, "message/rfc822"
    for header in ORIGINAL_MESSAGE_ID_HEADERS:
        value = message.get(header)
        message_id = first_message_id(value, prefer_last=(header == "References"))
        if message_id:
            return message_id, header
    return "", ""


def extract_original_message_id(message):
    return extract_original_message_id_details(message)[0]


def extract_bounce_reason(message):
    for part in message.walk():
        if part.get_content_type() != "message/delivery-status":
            continue
        payload = part.get_payload()
        blocks = payload if isinstance(payload, list) else [part]
        for block in blocks:
            diagnostic = block.get("Diagnostic-Code")
            if diagnostic:
                return str(diagnostic).split(";", 1)[-1].strip()[:1000]
            status = block.get("Status")
            if status:
                return f"Status {status}"[:1000]
    text = message_text(message)
    match = re.search(r"Diagnostic-Code:\s*(?:smtp;)?\s*(.+)", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()[:1000]
    return (decode_text(message.get("Subject")) or "Delivery failed")[:1000]


class BounceTrackingService:
    def __init__(self, base_url, state_path, log_folder=None, imap_ssl_factory=None, imap_factory=None, http_post=None, now=None):
        self.base_url = str(base_url or "").rstrip("/")
        self.state_path = Path(state_path)
        self.log_folder = Path(log_folder) if log_folder else None
        self.imap_ssl_factory = imap_ssl_factory or imaplib.IMAP4_SSL
        self.imap_factory = imap_factory or imaplib.IMAP4
        self.http_post = http_post or self._http_post
        self.now = now or (lambda: datetime.now(timezone.utc))

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
            entry = {"timestamp": self.now().isoformat(), "event": event, **values}
            with (self.log_folder / "bounce-sync.log").open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(entry, ensure_ascii=False, default=str) + "\n")
        except Exception:
            pass

    def _load_processed(self):
        try:
            if self.state_path.exists():
                data = json.loads(self.state_path.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    return set(str(value) for value in data)
        except Exception:
            pass
        return set()

    def _save_processed(self, processed):
        self.state_path.parent.mkdir(parents=True, exist_ok=True)
        self.state_path.write_text(json.dumps(sorted(processed), indent=2), encoding="utf-8")

    def _connect(self, configuration):
        encryption = configuration.get("imap_encryption", "SSL/TLS")
        factory = self.imap_ssl_factory if encryption == "SSL/TLS" else self.imap_factory
        client = factory(configuration["imap_host"], int(configuration["imap_port"]))
        if encryption == "STARTTLS":
            client.starttls()
        client.login(configuration["email"], configuration["password"])
        self._log("IMAP Connected", email=configuration.get("email"), imap_host=configuration.get("imap_host"), imap_port=configuration.get("imap_port"), imap_encryption=encryption, success=True)
        return client

    def _recent_search_date(self):
        return (self.now() - timedelta(days=14)).strftime("%d-%b-%Y")

    def _search_recent_unread(self, client):
        since = self._recent_search_date()
        criteria = ["UNSEEN", "SINCE", since]
        self._log("IMAP Search Criteria", criteria=criteria, mode="UID SEARCH")
        try:
            status, data = client.uid("SEARCH", None, "UNSEEN", "SINCE", since)
            self._log("Raw IMAP SEARCH Response", mode="UID SEARCH", status=status, raw_response=[item.decode(errors="replace") if isinstance(item, bytes) else str(item) for item in (data or [])])
            if status == "OK":
                identifiers = data[0].split() if data and data[0] else []
                self._log("Unread Email Count", count=len(identifiers))
                return True, identifiers
        except Exception:
            self._log("Errors", error="UID SEARCH failed; falling back to SEARCH", exception=True)
        self._log("IMAP Search Criteria", criteria=criteria, mode="SEARCH")
        status, data = client.search(None, "UNSEEN", "SINCE", since)
        self._log("Raw IMAP SEARCH Response", mode="SEARCH", status=status, raw_response=[item.decode(errors="replace") if isinstance(item, bytes) else str(item) for item in (data or [])])
        if status != "OK":
            raise RuntimeError("Unable to search recent unread IMAP messages")
        identifiers = data[0].split() if data and data[0] else []
        self._log("Unread Email Count", count=len(identifiers))
        return False, identifiers

    def _fetch_message(self, client, use_uid, identifier):
        if use_uid:
            status, payload = client.uid("FETCH", identifier, "(RFC822)")
        else:
            status, payload = client.fetch(identifier, "(RFC822)")
        if status != "OK":
            return None
        return next((item[1] for item in payload if isinstance(item, tuple) and len(item) > 1), None)

    def _mark_seen(self, client, use_uid, identifier):
        if use_uid:
            return client.uid("STORE", identifier, "+FLAGS", "\\Seen")
        return client.store(identifier, "+FLAGS", "\\Seen")

    def register_bounce(self, message_id, bounce_reason):
        url = f"{self.base_url}{REGISTER_BOUNCE_ENDPOINT}"
        payload = {"message_id": message_id, "bounce_reason": bounce_reason}
        self._log("Register Bounce API Call", called=True, method="POST", url=url, payload=payload)
        self._log("API Request", method="POST", url=url, payload=payload)
        response = self.http_post(url, payload)
        self._log(
            "API Response",
            method="POST",
            url=url,
            http_status=response.get("status_code") if isinstance(response, dict) else None,
            response_body=response.get("body") if isinstance(response, dict) else "",
            response=response,
        )
        status_code = int(response.get("status_code") or 0) if isinstance(response, dict) else 0
        if status_code and status_code >= 400:
            raise RuntimeError(f"register-bounce API failed with HTTP {status_code}: {response.get('body', '')}")
        return response

    def check_account(self, configuration):
        if not self.base_url:
            raise ValueError("Tracking base URL is not configured.")
        self._log("Synchronization Started", flow="bounce_detection", account=configuration.get("email"))
        client = self._connect(configuration)
        processed = self._load_processed()
        detected = registered = duplicates = skipped = errors = 0
        try:
            status, _ = client.select("INBOX")
            if status != "OK":
                raise RuntimeError("Unable to select IMAP inbox")
            self._log("Mailbox Selected", mailbox="INBOX", status=status)
            use_uid, identifiers = self._search_recent_unread(client)
            for identifier in identifiers:
                identifier_text = identifier.decode() if isinstance(identifier, bytes) else str(identifier)
                duplicate_key = f"uid:{identifier_text}" if use_uid else ""
                try:
                    if duplicate_key and duplicate_key in processed:
                        duplicates += 1
                        self._log("Duplicate Bounce Skipped", identifier=identifier_text)
                        self._mark_seen(client, use_uid, identifier)
                        continue
                    raw = self._fetch_message(client, use_uid, identifier)
                    if not raw:
                        skipped += 1
                        self._log("Unread Email Fetch Empty", identifier=identifier_text)
                        continue
                    message = email.message_from_bytes(raw, policy=policy.default)
                    message_header_id = str(message.get("Message-ID") or "")
                    message_subject = decode_text(message.get("Subject"))
                    message_sender = decode_text(message.get("From"))
                    message_in_reply_to = str(message.get("In-Reply-To") or "")
                    message_references = str(message.get("References") or "")
                    message_content_type = message.get_content_type()
                    classified_as_bounce = is_bounce_message(message)
                    self._log(
                        "Unread Email Inspected",
                        uid=identifier_text,
                        identifier=identifier_text,
                        subject=message_subject,
                        from_email=message_sender,
                        sender=message_sender,
                        content_type=message_content_type,
                        message_id=message_header_id,
                        in_reply_to=message_in_reply_to,
                        references=message_references,
                        classification="Bounce" if classified_as_bounce else "Neither",
                        classified_as_bounce=classified_as_bounce,
                        bounce_email_message_id=message_header_id,
                    )
                    message_key = f"message-id:{message_header_id}" if message_header_id else ""
                    if message_key and message_key in processed:
                        duplicates += 1
                        self._log("Duplicate Bounce Skipped", identifier=identifier_text, message_id=message_header_id)
                        if duplicate_key:
                            processed.add(duplicate_key)
                            self._save_processed(processed)
                        self._mark_seen(client, use_uid, identifier)
                        continue
                    if not duplicate_key:
                        duplicate_key = message_key or f"message-id:{identifier_text}"
                    if duplicate_key in processed:
                        duplicates += 1
                        self._log("Duplicate Bounce Skipped", identifier=identifier_text, message_id=message_header_id)
                        self._mark_seen(client, use_uid, identifier)
                        continue
                    if not classified_as_bounce:
                        skipped += 1
                        continue
                    original_message_id, extraction_source = extract_original_message_id_details(message)
                    bounce_reason = extract_bounce_reason(message)
                    self._log(
                        "Bounce Message-ID Extraction",
                        identifier=identifier_text,
                        bounce_email_message_id=message_header_id,
                        original_message_id=original_message_id,
                        extraction_source=extraction_source,
                        bounce_reason=bounce_reason,
                    )
                    if not original_message_id:
                        skipped += 1
                        self._log(
                            "Warning",
                            warning="Original Message-ID not found",
                            reason="No Message-ID found in embedded message/rfc822 or allowed original-message headers.",
                            uid=identifier_text,
                            identifier=identifier_text,
                            bounce_email_message_id=message_header_id,
                            extracted_original_message_id=original_message_id,
                            extraction_source=extraction_source,
                            api_called=False,
                        )
                        continue
                    detected += 1
                    self._log(
                        "Bounce Detected",
                        uid=identifier_text,
                        identifier=identifier_text,
                        bounce_email_message_id=message_header_id,
                        extracted_original_message_id=original_message_id,
                        original_message_id=original_message_id,
                        extraction_source=extraction_source,
                        bounce_reason=bounce_reason,
                        sender=message_sender,
                        api_called=True,
                    )
                    self.register_bounce(original_message_id, bounce_reason)
                    registered += 1
                    processed.add(duplicate_key)
                    if message_key:
                        processed.add(message_key)
                    self._save_processed(processed)
                    self._mark_seen(client, use_uid, identifier)
                except Exception as exc:
                    errors += 1
                    self._log("Errors", identifier=identifier_text, error=str(exc))
            return {"detected": detected, "registered": registered, "duplicates": duplicates, "skipped": skipped, "errors": errors}
        finally:
            try:
                client.logout()
            except Exception:
                pass
