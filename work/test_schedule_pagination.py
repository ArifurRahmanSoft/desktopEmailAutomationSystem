import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import email_automation as app


def schedules(count):
    return [
        {
            "id": f"schedule-{index}",
            "name": f"Schedule {index:02d}",
            "excel_file": f"C:/Temp/mail_{index:02d}.xlsx",
            "days": ["Monday", "Tuesday"],
            "time": "09:00",
            "max_emails": index,
            "enabled": index % 2 == 0,
        }
        for index in range(1, count + 1)
    ]


class SchedulePaginationTests(unittest.TestCase):
    def test_daily_scheduling_pagination_controls(self):
        with patch.object(app, "load_schedules", return_value=schedules(45)):
            window = app.MultiScheduleWindow()
            try:
                window.root.update_idletasks()
                self.assertEqual(window.page, 1)
                self.assertEqual(window.total_pages, 3)
                self.assertEqual(window.page_label.get(), "Page 1 of 3")
                self.assertEqual(len(window.tree.get_children()), 20)
                self.assertEqual(str(window.first_button.cget("state")), "disabled")
                self.assertEqual(str(window.previous_button.cget("state")), "disabled")
                self.assertEqual(str(window.next_button.cget("state")), "normal")
                self.assertEqual(str(window.last_button.cget("state")), "normal")

                window.goto_page(2)
                self.assertEqual(window.page_label.get(), "Page 2 of 3")
                self.assertEqual(str(window.first_button.cget("state")), "normal")
                self.assertEqual(str(window.previous_button.cget("state")), "normal")
                self.assertEqual(str(window.next_button.cget("state")), "normal")
                self.assertEqual(str(window.last_button.cget("state")), "normal")

                window.goto_page(window.total_pages)
                self.assertEqual(window.page_label.get(), "Page 3 of 3")
                self.assertEqual(len(window.tree.get_children()), 5)
                self.assertEqual(str(window.next_button.cget("state")), "disabled")
                self.assertEqual(str(window.last_button.cget("state")), "disabled")

                window.page_size.set(10)
                window.change_page_size()
                self.assertEqual(window.page_label.get(), "Page 1 of 5")
                self.assertEqual(len(window.tree.get_children()), 10)
            finally:
                window.root.destroy()


if __name__ == "__main__":
    unittest.main()
