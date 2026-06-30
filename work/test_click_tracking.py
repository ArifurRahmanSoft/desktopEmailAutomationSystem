import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from email_automation import build_click_tracked_html


class ClickTrackingTests(unittest.TestCase):
    def setUp(self):
        self.tracking_id = "7d19af31-2d65-49db-b52e-2c92b5d39b61"

    def test_rewrites_url_and_preserves_visible_text(self):
        result = build_click_tracked_html("Visit https://powersoft.com today", self.tracking_id)
        expected_href = f"https://emailtrackingserver.onrender.com/email/click/{self.tracking_id}?url=https%3A%2F%2Fpowersoft.com"
        self.assertIn(f'href="{expected_href}"', result)
        self.assertIn(">https://powersoft.com</a>", result)
        self.assertTrue(result.startswith("Visit "))
        self.assertTrue(result.endswith(" today"))

    def test_supports_unlimited_http_and_https_urls(self):
        result = build_click_tracked_html("http://example.com and https://github.com/user/repo", self.tracking_id)
        self.assertEqual(result.count("/email/click/"), 2)
        self.assertIn("url=http%3A%2F%2Fexample.com", result)
        self.assertIn("url=https%3A%2F%2Fgithub.com%2Fuser%2Frepo", result)

    def test_trailing_sentence_punctuation_is_not_part_of_link(self):
        result = build_click_tracked_html("Open https://powersoft.com.", self.tracking_id)
        self.assertIn(">https://powersoft.com</a>.", result)
        self.assertNotIn("powersoft.com.%", result)

    def test_no_url_preserves_existing_html_conversion(self):
        self.assertEqual(build_click_tracked_html("Hello & welcome\nNext line", self.tracking_id), "Hello &amp; welcome<br>\nNext line")

    def test_original_html_is_escaped(self):
        result = build_click_tracked_html("<b>https://example.com</b>", self.tracking_id)
        self.assertTrue(result.startswith("&lt;b&gt;"))
        self.assertTrue(result.endswith("&lt;/b&gt;"))


if __name__ == "__main__": unittest.main()
