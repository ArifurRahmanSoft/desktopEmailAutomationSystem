import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from placeholder_service import PlaceholderService


class PlaceholderServiceTests(unittest.TestCase):
    def setUp(self):
        self.headers = ["First_Name", "Last_Name", "Company", "Designation", "Phone"]
        self.values = ["John", "Doe", None, "Senior Developer", "+1 555"]
        self.context = PlaceholderService.create_context(self.headers, self.values)

    def test_case_insensitive_and_whitespace_tolerant(self):
        value = PlaceholderService.render("Hi {{ first_name }} {{LAST_NAME}}", self.context)
        self.assertEqual(value, "Hi John Doe")

    def test_empty_value_becomes_empty_string(self):
        value = PlaceholderService.render("{{Designation}} of {{Company}}", self.context)
        self.assertEqual(value, "Senior Developer of ")

    def test_unknown_placeholder_becomes_empty_string(self):
        self.assertEqual(PlaceholderService.render("X{{RandomField}}Y", self.context), "XY")

    def test_future_column_is_automatically_available(self):
        self.assertEqual(PlaceholderService.render("Call {{Phone}}", self.context), "Call +1 555")

    def test_render_row_builds_context_dynamically(self):
        self.assertEqual(PlaceholderService.render_row("{{First_Name}}", self.headers, self.values), "John")


if __name__ == "__main__":
    unittest.main()
