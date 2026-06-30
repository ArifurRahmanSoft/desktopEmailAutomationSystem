import tempfile
import unittest
from pathlib import Path

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


if __name__ == "__main__": unittest.main()
