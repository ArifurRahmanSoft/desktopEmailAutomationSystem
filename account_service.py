import json
import re
import smtplib
from pathlib import Path

from credential_service import CredentialService


class AccountService:
    GMAIL_DEFAULTS = {"provider": "Gmail", "smtp_host": "smtp.gmail.com", "smtp_port": 587, "encryption": "STARTTLS", "imap_host": "imap.gmail.com", "imap_port": 993, "imap_encryption": "SSL/TLS"}
    PROVIDERS = ("Gmail", "Custom SMTP")
    ENCRYPTION_TYPES = ("STARTTLS", "SSL/TLS", "None")

    def __init__(self, storage_path, credentials=None):
        self.storage_path = Path(storage_path)
        self.credentials = credentials or CredentialService()

    def list_accounts(self):
        if not self.storage_path.exists():
            return []
        try:
            return [self.normalize_account(account) for account in json.loads(self.storage_path.read_text(encoding="utf-8"))]
        except Exception:
            return []

    @classmethod
    def normalize_account(cls, account):
        normalized = dict(account)
        provider = normalized.get("provider") or "Gmail"
        normalized["provider"] = provider
        normalized["display_name"] = normalized.get("display_name") or normalized.get("name") or normalized.get("email", "")
        normalized["name"] = normalized["display_name"]
        if provider == "Gmail":
            normalized.update(cls.GMAIL_DEFAULTS)
        else:
            normalized["smtp_host"] = str(normalized.get("smtp_host") or "").strip()
            normalized["smtp_port"] = int(normalized.get("smtp_port") or 0)
            normalized["encryption"] = normalized.get("encryption") or "STARTTLS"
            normalized["imap_host"] = str(normalized.get("imap_host") or normalized["smtp_host"]).strip()
            normalized["imap_port"] = int(normalized.get("imap_port") or 993)
            normalized["imap_encryption"] = normalized.get("imap_encryption") or "SSL/TLS"
        normalized["enabled"] = bool(normalized.get("enabled", True))
        return normalized

    def _save(self, accounts):
        self.storage_path.parent.mkdir(parents=True, exist_ok=True)
        self.storage_path.write_text(json.dumps(accounts, indent=2), encoding="utf-8")

    def find(self, email, enabled_only=False):
        key = str(email or "").strip().lower()
        for account in self.list_accounts():
            if account["email"].lower() == key and (not enabled_only or account.get("enabled", True)):
                return account
        return None

    def save_account(self, name, email, password="", enabled=True, original_email=None, provider="Gmail", smtp_host=None, smtp_port=None, encryption=None, imap_host=None, imap_port=None, imap_encryption=None):
        email = email.strip().lower()
        if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email):
            raise ValueError("A valid email address is required.")
        if provider not in self.PROVIDERS:
            raise ValueError("Invalid mail provider.")
        if provider == "Gmail":
            smtp_host, smtp_port, encryption = "smtp.gmail.com", 587, "STARTTLS"
            imap_host, imap_port, imap_encryption = "imap.gmail.com", 993, "SSL/TLS"
        else:
            smtp_host = str(smtp_host or "").strip()
            if not smtp_host:
                raise ValueError("SMTP Host cannot be empty.")
            try:
                smtp_port = int(smtp_port)
            except (TypeError, ValueError):
                raise ValueError("SMTP Port must be numeric.")
            if not 1 <= smtp_port <= 65535:
                raise ValueError("SMTP Port must be between 1 and 65535.")
            if encryption not in self.ENCRYPTION_TYPES:
                raise ValueError("Invalid encryption type.")
            imap_host = str(imap_host or smtp_host).strip()
            try: imap_port = int(imap_port or 993)
            except (TypeError, ValueError): raise ValueError("IMAP Port must be numeric.")
            if not 1 <= imap_port <= 65535: raise ValueError("IMAP Port must be between 1 and 65535.")
            if imap_encryption not in self.ENCRYPTION_TYPES: raise ValueError("Invalid IMAP encryption type.")
        account_values = {"name": name.strip() or email, "display_name": name.strip() or email, "email": email, "provider": provider, "smtp_host": smtp_host, "smtp_port": smtp_port, "encryption": encryption, "imap_host": imap_host, "imap_port": imap_port, "imap_encryption": imap_encryption, "enabled": bool(enabled)}
        accounts = self.list_accounts()
        old_key = (original_email or email).strip().lower()
        existing = next((a for a in accounts if a["email"].lower() == old_key), None)
        if existing:
            existing.update(account_values)
        else:
            accounts.append(account_values)
        if original_email and old_key != email:
            old_password = self.credentials.get_password(old_key)
            self.credentials.delete_password(old_key)
            if not password and old_password:
                password = old_password
        if password:
            stored_password = password.replace(" ", "") if provider == "Gmail" else password
            self.credentials.save_password(email, stored_password)
        elif not self.credentials.get_password(email):
            raise ValueError("Password / App Password is required for a new account.")
        self._save(accounts)

    def delete_account(self, email):
        key = email.strip().lower()
        self._save([a for a in self.list_accounts() if a["email"].lower() != key])
        self.credentials.delete_password(key)

    def set_enabled(self, email, enabled):
        accounts = self.list_accounts()
        for account in accounts:
            if account["email"].lower() == email.strip().lower():
                account["enabled"] = bool(enabled)
        self._save(accounts)

    def smtp_credentials(self, email):
        account = self.find(email, enabled_only=True)
        password = self.credentials.get_password(email) if account else None
        if not account or not password:
            return None
        return account["email"], password

    def smtp_configuration(self, email):
        account = self.find(email, enabled_only=True)
        password = self.credentials.get_password(email) if account else None
        if not account or not password:
            return None
        return {**account, "password": password}

    def imap_configuration(self, email):
        account = self.find(email, enabled_only=True)
        password = self.credentials.get_password(email) if account else None
        if not account or not password:
            return None
        return {**account, "password": password}

    @staticmethod
    def connect_smtp(configuration, timeout=30):
        host = configuration["smtp_host"]
        port = int(configuration["smtp_port"])
        encryption = configuration["encryption"]
        if encryption == "SSL/TLS":
            smtp = smtplib.SMTP_SSL(host, port, timeout=timeout)
            smtp.ehlo()
        else:
            smtp = smtplib.SMTP(host, port, timeout=timeout)
            smtp.ehlo()
            if encryption == "STARTTLS":
                smtp.starttls()
                smtp.ehlo()
        smtp.login(configuration["email"], configuration["password"])
        return smtp

    def test_smtp(self, email):
        configuration = self.smtp_configuration(email)
        if not configuration:
            raise ValueError("Sender Account Not Configured")
        smtp = self.connect_smtp(configuration, timeout=30)
        try:
            smtp.noop()
        finally:
            try: smtp.quit()
            except Exception: smtp.close()
        return True
