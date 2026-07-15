import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from account_service import AccountService


class FakeCredentials:
    def __init__(self): self.values = {}
    def save_password(self, email, password): self.values[email] = password
    def get_password(self, email): return self.values.get(email)
    def delete_password(self, email): self.values.pop(email, None)


class AccountServiceTests(unittest.TestCase):
    def test_account_lifecycle_and_password_retention(self):
        with tempfile.TemporaryDirectory() as temp:
            credentials = FakeCredentials()
            service = AccountService(Path(temp) / "accounts.json", credentials)
            service.save_account("Sales", "Sales@Gmail.com", "secret", True)
            self.assertEqual(service.smtp_credentials("sales@gmail.com"), ("sales@gmail.com", "secret"))
            service.save_account("Sales Team", "sales@gmail.com", "", True, "sales@gmail.com")
            self.assertEqual(credentials.get_password("sales@gmail.com"), "secret")
            service.set_enabled("sales@gmail.com", False)
            self.assertIsNone(service.smtp_credentials("sales@gmail.com"))
            service.delete_account("sales@gmail.com")
            self.assertEqual(service.list_accounts(), [])
            self.assertIsNone(credentials.get_password("sales@gmail.com"))

    def test_legacy_account_receives_gmail_defaults_without_migration(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "accounts.json"
            path.write_text('[{"name":"Legacy","email":"legacy@gmail.com","enabled":true}]', encoding="utf-8")
            account = AccountService(path, FakeCredentials()).list_accounts()[0]
            self.assertEqual(account["provider"], "Gmail")
            self.assertEqual(account["smtp_host"], "smtp.gmail.com")
            self.assertEqual(account["smtp_port"], 587)
            self.assertEqual(account["encryption"], "STARTTLS")
            self.assertEqual((account["imap_host"], account["imap_port"], account["imap_encryption"]), ("imap.gmail.com", 993, "SSL/TLS"))
            self.assertEqual(account["display_name"], "Legacy")
            self.assertEqual(account["sender_alias"], "")

    def test_gmail_alias_is_stored_and_used_only_as_visible_sender(self):
        with tempfile.TemporaryDirectory() as temp:
            credentials = FakeCredentials(); service = AccountService(Path(temp) / "accounts.json", credentials)
            service.save_account("Sales", "login@gmail.com", "app password", sender_alias="Alias@Example.com")
            configuration = service.smtp_configuration("login@gmail.com")
            self.assertEqual(configuration["email"], "login@gmail.com")
            self.assertEqual(configuration["sender_alias"], "alias@example.com")
            self.assertEqual(AccountService.sender_address(configuration), "alias@example.com")
            self.assertEqual(service.smtp_credentials("login@gmail.com"), ("login@gmail.com", "apppassword"))

    def test_blank_or_non_gmail_alias_keeps_configured_email(self):
        gmail = {"provider":"Gmail","smtp_host":"smtp.gmail.com","email":"login@gmail.com","sender_alias":""}
        custom = {"provider":"Custom SMTP","smtp_host":"mail.example.com","email":"login@example.com","sender_alias":"alias@example.com"}
        self.assertEqual(AccountService.sender_address(gmail), "login@gmail.com")
        self.assertEqual(AccountService.sender_address(custom), "login@example.com")

    def test_invalid_gmail_alias_is_rejected(self):
        with tempfile.TemporaryDirectory() as temp:
            service = AccountService(Path(temp) / "accounts.json", FakeCredentials())
            with self.assertRaisesRegex(ValueError, "Sender Alias"):
                service.save_account("X", "login@gmail.com", "password", sender_alias="not-an-email")

    def test_custom_smtp_fields_are_stored_and_password_spaces_preserved(self):
        with tempfile.TemporaryDirectory() as temp:
            credentials = FakeCredentials(); service = AccountService(Path(temp) / "accounts.json", credentials)
            service.save_account("Business", "mail@business.com", "pass word", True, provider="Custom SMTP", smtp_host="mail.business.com", smtp_port="465", encryption="SSL/TLS", imap_host="imap.business.com", imap_port="993", imap_encryption="SSL/TLS")
            account = service.list_accounts()[0]
            self.assertEqual((account["provider"], account["smtp_host"], account["smtp_port"], account["encryption"]), ("Custom SMTP", "mail.business.com", 465, "SSL/TLS"))
            self.assertEqual((account["imap_host"], account["imap_port"]), ("imap.business.com", 993))
            self.assertEqual(credentials.get_password("mail@business.com"), "pass word")

    def test_validation_rejects_invalid_custom_settings(self):
        with tempfile.TemporaryDirectory() as temp:
            service = AccountService(Path(temp) / "accounts.json", FakeCredentials())
            with self.assertRaisesRegex(ValueError, "valid email"): service.save_account("X", "invalid", "p")
            with self.assertRaisesRegex(ValueError, "Host"): service.save_account("X", "x@example.com", "p", provider="Custom SMTP", smtp_host="", smtp_port=587, encryption="STARTTLS")
            with self.assertRaisesRegex(ValueError, "numeric"): service.save_account("X", "x@example.com", "p", provider="Custom SMTP", smtp_host="smtp.example.com", smtp_port="abc", encryption="STARTTLS")

    def test_connection_modes(self):
        class FakeSMTP:
            def __init__(self, host, port, timeout): self.calls=[("init",host,port,timeout)]
            def ehlo(self): self.calls.append(("ehlo",))
            def starttls(self): self.calls.append(("starttls",))
            def login(self, email, password): self.calls.append(("login",email,password))
        base = {"email":"user@example.com","password":"secret","smtp_host":"smtp.example.com","smtp_port":587}
        with patch("account_service.smtplib.SMTP", FakeSMTP), patch("account_service.smtplib.SMTP_SSL", FakeSMTP):
            starttls = AccountService.connect_smtp({**base,"encryption":"STARTTLS"})
            self.assertIn(("starttls",), starttls.calls)
            plain = AccountService.connect_smtp({**base,"encryption":"None"})
            self.assertNotIn(("starttls",), plain.calls)
            ssl = AccountService.connect_smtp({**base,"encryption":"SSL/TLS"})
            self.assertNotIn(("starttls",), ssl.calls)
            gmail_alias = AccountService.connect_smtp({**base,"provider":"Gmail","smtp_host":"smtp.gmail.com","sender_alias":"alias@example.com","encryption":"STARTTLS"})
            self.assertIn(("login","user@example.com","secret"), gmail_alias.calls)


if __name__ == "__main__": unittest.main()
