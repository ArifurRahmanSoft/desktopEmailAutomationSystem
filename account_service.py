import json
import smtplib
from pathlib import Path

from credential_service import CredentialService


class AccountService:
    def __init__(self, storage_path, credentials=None):
        self.storage_path = Path(storage_path)
        self.credentials = credentials or CredentialService()

    def list_accounts(self):
        if not self.storage_path.exists():
            return []
        try:
            return json.loads(self.storage_path.read_text(encoding="utf-8"))
        except Exception:
            return []

    def _save(self, accounts):
        self.storage_path.parent.mkdir(parents=True, exist_ok=True)
        self.storage_path.write_text(json.dumps(accounts, indent=2), encoding="utf-8")

    def find(self, email, enabled_only=False):
        key = str(email or "").strip().lower()
        for account in self.list_accounts():
            if account["email"].lower() == key and (not enabled_only or account.get("enabled", True)):
                return account
        return None

    def save_account(self, name, email, password="", enabled=True, original_email=None):
        email = email.strip().lower()
        if not email or "@" not in email:
            raise ValueError("A valid Gmail address is required.")
        accounts = self.list_accounts()
        old_key = (original_email or email).strip().lower()
        existing = next((a for a in accounts if a["email"].lower() == old_key), None)
        if existing:
            existing.update({"name": name.strip() or email, "email": email, "enabled": bool(enabled)})
        else:
            accounts.append({"name": name.strip() or email, "email": email, "enabled": bool(enabled)})
        if original_email and old_key != email:
            old_password = self.credentials.get_password(old_key)
            self.credentials.delete_password(old_key)
            if not password and old_password:
                password = old_password
        if password:
            self.credentials.save_password(email, password.replace(" ", ""))
        elif not self.credentials.get_password(email):
            raise ValueError("Google App Password is required for a new account.")
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

    def test_smtp(self, email):
        values = self.smtp_credentials(email)
        if not values:
            raise ValueError("Sender Account Not Configured")
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as smtp:
            smtp.ehlo(); smtp.starttls(); smtp.ehlo(); smtp.login(*values)
        return True
