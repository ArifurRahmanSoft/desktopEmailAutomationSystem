import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import email_automation as app


class EmailValidationTests(unittest.TestCase):
    def test_valid_syntax_domain_and_mx(self):
        with patch.object(app, "domain_has_mx", return_value=True) as mx:
            self.assertTrue(app.validate_recipient_email("person@example.com"))
            mx.assert_called_once_with("example.com")

    def test_missing_mx_is_invalid(self):
        with patch.object(app, "domain_has_mx", return_value=False):
            self.assertFalse(app.validate_recipient_email("person@example.com"))

    def test_invalid_syntax_does_not_query_dns(self):
        invalid = ("plainaddress", "a@@example.com", ".a@example.com", "a..b@example.com", "a@-example.com", "a@example")
        with patch.object(app, "domain_has_mx") as mx:
            for value in invalid:
                self.assertFalse(app.validate_recipient_email(value), value)
            mx.assert_not_called()

    def test_idn_domain_is_normalized_before_mx_lookup(self):
        with patch.object(app, "domain_has_mx", return_value=True) as mx:
            self.assertTrue(app.validate_recipient_email("person@bücher.de"))
            mx.assert_called_once_with("xn--bcher-kva.de")


if __name__ == "__main__": unittest.main()
