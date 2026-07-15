import sys
import tempfile
import unittest
from email.message import EmailMessage
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from reply_tracking_service import ReplyTrackingService, extract_original_message_id


def reply_message(message_id="<reply-1@example.com>", in_reply_to="<original@example.com>", references=""):
    message = EmailMessage()
    message["From"] = "Recipient <recipient@example.com>"
    message["Date"] = "Sat, 11 Jul 2026 10:15:00 +0600"
    message["Message-ID"] = message_id
    if in_reply_to:
        message["In-Reply-To"] = in_reply_to
    if references:
        message["References"] = references
    message.set_content("Thanks for your email.")
    return message


class FakeIMAP:
    def __init__(self, messages):
        self.messages = messages
        self.seen = []
        self.logged_out = False
    def login(self, email, password):
        self.login_values = (email, password)
    def select(self, folder):
        return "OK", [b""]
    def search(self, *_):
        numbers = b" ".join(str(index + 1).encode() for index in range(len(self.messages)))
        return "OK", [numbers]
    def fetch(self, number, _):
        return "OK", [(b"RFC822", self.messages[int(number) - 1])]
    def store(self, number, *_):
        self.seen.append(number)
        return "OK", []
    def logout(self):
        self.logged_out = True


class ReplyTrackingTests(unittest.TestCase):
    def config(self):
        return {
            "email": "sender@example.com",
            "password": "secret",
            "imap_host": "imap.example.com",
            "imap_port": 993,
            "imap_encryption": "SSL/TLS",
        }

    def test_extracts_in_reply_to_first(self):
        message = reply_message(in_reply_to="<direct@example.com>", references="<older@example.com> <latest@example.com>")
        self.assertEqual(extract_original_message_id(message), "<direct@example.com>")

    def test_extracts_references_when_in_reply_to_missing(self):
        message = reply_message(in_reply_to="", references="<older@example.com> <latest@example.com>")
        self.assertEqual(extract_original_message_id(message), "<latest@example.com>")

    def test_registers_unread_reply_and_marks_seen_after_success(self):
        with tempfile.TemporaryDirectory() as temp:
            posts = []
            client = FakeIMAP([reply_message().as_bytes()])
            service = ReplyTrackingService(
                "https://server.test",
                Path(temp),
                imap_ssl_factory=lambda *_: client,
                http_post=lambda url, body: posts.append((url, body)) or {"status_code": 200, "body": "{}"},
            )
            result = service.check_account(self.config())

            self.assertEqual(result["detected"], 1)
            self.assertEqual(result["registered"], 1)
            self.assertEqual(posts[0][0], "https://server.test/api/tracking/register-reply")
            self.assertEqual(posts[0][1]["message_id"], "<original@example.com>")
            self.assertEqual(posts[0][1]["from_email"], "recipient@example.com")
            self.assertTrue(posts[0][1]["reply_time"].startswith("2026-07-11T04:15:00"))
            self.assertEqual(client.seen, [b"1"])
            self.assertTrue(client.logged_out)
            self.assertTrue((Path(temp) / "reply-tracking.log").exists())

    def test_skips_non_reply_without_marking_seen(self):
        message = reply_message(in_reply_to="")
        del message["Message-ID"]
        message["Message-ID"] = "<not-a-reply@example.com>"
        client = FakeIMAP([message.as_bytes()])
        posts = []
        service = ReplyTrackingService("https://server.test", None, imap_ssl_factory=lambda *_: client, http_post=lambda url, body: posts.append((url, body)))

        result = service.check_account(self.config())

        self.assertEqual(result["detected"], 0)
        self.assertEqual(result["registered"], 0)
        self.assertEqual(posts, [])
        self.assertEqual(client.seen, [])

    def test_duplicate_reply_message_is_counted_once(self):
        raw = reply_message().as_bytes()
        client = FakeIMAP([raw, raw])
        posts = []
        service = ReplyTrackingService("https://server.test", None, imap_ssl_factory=lambda *_: client, http_post=lambda url, body: posts.append((url, body)) or {})

        result = service.check_account(self.config())

        self.assertEqual(result["detected"], 1)
        self.assertEqual(result["registered"], 1)
        self.assertEqual(len(posts), 1)
        self.assertEqual(client.seen, [b"1", b"2"])

    def test_api_failure_does_not_mark_seen(self):
        client = FakeIMAP([reply_message().as_bytes()])
        service = ReplyTrackingService("https://server.test", None, imap_ssl_factory=lambda *_: client, http_post=lambda *_: (_ for _ in ()).throw(RuntimeError("server down")))

        result = service.check_account(self.config())

        self.assertEqual(result["detected"], 1)
        self.assertEqual(result["registered"], 0)
        self.assertEqual(client.seen, [])


if __name__ == "__main__":
    unittest.main()
