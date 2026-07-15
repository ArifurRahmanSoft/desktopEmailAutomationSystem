import sys
import tempfile
import unittest
import email
from email import policy
from email.message import EmailMessage
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from bounce_tracking_service import BounceTrackingService, extract_bounce_reason, extract_original_message_id, extract_original_message_id_details, is_bounce_message


def bounce_message(message_id="<bounce-1@example.com>", original_message_id="<original-send@example.com>"):
    raw = (
        "From: Mail Delivery Subsystem <mailer-daemon@googlemail.com>\r\n"
        "Subject: Delivery Status Notification (Failure)\r\n"
        f"Message-ID: {message_id}\r\n"
        "MIME-Version: 1.0\r\n"
        'Content-Type: multipart/report; report-type=delivery-status; boundary="dsn-boundary"\r\n'
        "\r\n"
        "--dsn-boundary\r\n"
        "Content-Type: text/plain; charset=utf-8\r\n"
        "\r\n"
        "Delivery failed.\r\n"
        "--dsn-boundary\r\n"
        "Content-Type: message/delivery-status\r\n"
        "\r\n"
        "Final-Recipient: rfc822; bad@example.com\r\n"
        "Action: failed\r\n"
        "Status: 5.1.1\r\n"
        "Diagnostic-Code: smtp; 550 5.1.1 User unknown\r\n"
        "\r\n"
        "--dsn-boundary\r\n"
        "Content-Type: message/rfc822\r\n"
        "\r\n"
        "From: Sender <sender@example.com>\r\n"
        "To: bad@example.com\r\n"
        "Subject: Original tracked email\r\n"
        f"Message-ID: {original_message_id}\r\n"
        "\r\n"
        "Original body.\r\n"
        "--dsn-boundary--\r\n"
    )
    return email.message_from_bytes(raw.encode("utf-8"), policy=policy.default)


class FakeIMAP:
    def __init__(self, messages):
        self.messages = messages
        self.seen = []
        self.logged_out = False
        self.uid_supported = True
    def login(self, email, password):
        self.login_values = (email, password)
    def select(self, folder):
        return "OK", [b""]
    def uid(self, command, *args):
        command = command.upper()
        if command == "SEARCH":
            values = b" ".join(str(index + 101).encode() for index in range(len(self.messages)))
            return "OK", [values]
        if command == "FETCH":
            uid = int(args[0]) - 101
            return "OK", [(b"RFC822", self.messages[uid])]
        if command == "STORE":
            self.seen.append(args[0])
            return "OK", []
        return "BAD", []
    def search(self, *_):
        values = b" ".join(str(index + 1).encode() for index in range(len(self.messages)))
        return "OK", [values]
    def fetch(self, number, _):
        return "OK", [(b"RFC822", self.messages[int(number)-1])]
    def store(self, number, *_):
        self.seen.append(number)
        return "OK", []
    def logout(self):
        self.logged_out = True


class BounceTrackingTests(unittest.TestCase):
    def config(self):
        return {"email":"sender@example.com","password":"secret","imap_host":"imap.example.com","imap_port":993,"imap_encryption":"SSL/TLS"}

    def test_detects_bounce_and_extracts_original_message_id_and_reason(self):
        message = bounce_message("<6a54996b.b50783f9.2d7ac2.d86f.GMR@mx.google.com>", "<f47c782b27eb40c9a17d8d063bdac0a9.20260713075748302571@emailautomation-v2.local>")
        self.assertTrue(is_bounce_message(message))
        self.assertEqual(extract_original_message_id(message), "<f47c782b27eb40c9a17d8d063bdac0a9.20260713075748302571@emailautomation-v2.local>")
        self.assertNotEqual(extract_original_message_id(message), "<6a54996b.b50783f9.2d7ac2.d86f.GMR@mx.google.com>")
        self.assertEqual(extract_original_message_id_details(message), ("<f47c782b27eb40c9a17d8d063bdac0a9.20260713075748302571@emailautomation-v2.local>", "message/rfc822"))
        self.assertIn("550 5.1.1 User unknown", extract_bounce_reason(message))

    def test_extracts_original_message_id_from_allowed_headers_only(self):
        message = EmailMessage()
        message["From"] = "Mail Delivery Subsystem <mailer-daemon@example.com>"
        message["Subject"] = "Delivery Status Notification (Failure)"
        message["Message-ID"] = "<bounce-message@example.com>"
        message["References"] = "<older@example.com> <original-from-references@example.com>"
        message.set_content("Delivery failed.")

        self.assertEqual(extract_original_message_id_details(message), ("<original-from-references@example.com>", "References"))

    def test_does_not_use_bounce_email_message_id_when_original_is_missing(self):
        message = EmailMessage()
        message["From"] = "Mail Delivery Subsystem <mailer-daemon@example.com>"
        message["Subject"] = "Delivery Status Notification (Failure)"
        message["Message-ID"] = "<bounce-message@example.com>"
        message.set_content("Delivery failed.\nMessage-ID: <bounce-message@example.com>")

        self.assertEqual(extract_original_message_id_details(message), ("", ""))

    def test_registers_bounce_and_marks_seen_after_success(self):
        with tempfile.TemporaryDirectory() as temp:
            posts = []
            client = FakeIMAP([bounce_message().as_bytes()])
            service = BounceTrackingService(
                "https://server.test",
                Path(temp) / "processed.json",
                Path(temp) / "logs",
                imap_ssl_factory=lambda *_: client,
                http_post=lambda url, body: posts.append((url, body)) or {"status_code": 200, "body": "{}"},
            )

            result = service.check_account(self.config())

            self.assertEqual(result["detected"], 1)
            self.assertEqual(result["registered"], 1)
            self.assertEqual(posts, [("https://server.test/api/tracking/register-bounce", {"message_id": "<original-send@example.com>", "bounce_reason": "550 5.1.1 User unknown"})])
            self.assertEqual(client.seen, [b"101"])
            self.assertTrue((Path(temp) / "processed.json").exists())
            self.assertTrue((Path(temp) / "logs" / "bounce-sync.log").exists())

    def test_duplicate_bounce_is_not_registered_twice(self):
        with tempfile.TemporaryDirectory() as temp:
            state = Path(temp) / "processed.json"
            posts = []
            raw = bounce_message().as_bytes()
            client_one = FakeIMAP([raw])
            service_one = BounceTrackingService("https://server.test", state, None, imap_ssl_factory=lambda *_: client_one, http_post=lambda url, body: posts.append((url, body)) or {})
            self.assertEqual(service_one.check_account(self.config())["registered"], 1)

            client_two = FakeIMAP([raw])
            service_two = BounceTrackingService("https://server.test", state, None, imap_ssl_factory=lambda *_: client_two, http_post=lambda url, body: posts.append((url, body)) or {})
            result = service_two.check_account(self.config())

            self.assertEqual(result["registered"], 0)
            self.assertEqual(result["duplicates"], 1)
            self.assertEqual(len(posts), 1)
            self.assertEqual(client_two.seen, [b"101"])

    def test_api_failure_does_not_mark_seen_or_store_duplicate_key(self):
        with tempfile.TemporaryDirectory() as temp:
            client = FakeIMAP([bounce_message().as_bytes()])
            service = BounceTrackingService("https://server.test", Path(temp) / "processed.json", None, imap_ssl_factory=lambda *_: client, http_post=lambda *_: (_ for _ in ()).throw(RuntimeError("server down")))

            result = service.check_account(self.config())

            self.assertEqual(result["detected"], 1)
            self.assertEqual(result["registered"], 0)
            self.assertEqual(result["errors"], 1)
            self.assertEqual(client.seen, [])
            self.assertFalse((Path(temp) / "processed.json").exists())


if __name__ == "__main__":
    unittest.main()
