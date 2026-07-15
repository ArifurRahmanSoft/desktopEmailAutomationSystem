import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import email_automation as app


def make_workbook(path, count):
    wb = Workbook()
    ws = wb.active
    ws.append(["First_Name", "Last_Name", "Email", "Subject", "Body", "Sender_Email", "Status", "Result", "SentDate"])
    for index in range(1, count + 1):
        ws.append([f"First{index:02d}", f"Last{index:02d}", f"user{index:02d}@example.com", "Subject", "Body", "sender@example.com", "Pending", "", ""])
    wb.save(path)


class SendMailPaginationTests(unittest.TestCase):
    def test_send_mail_pagination_controls_and_selection_persistence(self):
        with tempfile.TemporaryDirectory() as temp:
            installed = Path(temp) / "runtime"
            data = Path(temp) / "data"
            data.mkdir()
            (installed / "config").mkdir(parents=True)
            workbook = data / "mail_list.xlsx"
            make_workbook(workbook, 45)
            with patch.object(app, "install_dir", return_value=installed):
                window = app.SenderWindow()
                try:
                    window.excel_file.set(str(workbook))
                    window.refresh(reset_page=True)
                    window.root.update_idletasks()

                    self.assertEqual(window.page_label.get(), "Page 1 of 3")
                    self.assertEqual(str(window.first_button.cget("state")), "disabled")
                    self.assertEqual(str(window.previous_button.cget("state")), "disabled")
                    self.assertEqual(str(window.next_button.cget("state")), "normal")
                    self.assertEqual(str(window.last_button.cget("state")), "normal")
                    self.assertEqual(len(window.grid.get_children()), 20)

                    first_item = window.grid.get_children()[0]
                    first_values = tuple(str(value) for value in window.grid.item(first_item, "values"))
                    window.grid.selection_set(first_item)
                    window.remember_selection()

                    window.goto_page(2)
                    self.assertEqual(window.page_label.get(), "Page 2 of 3")
                    self.assertEqual(str(window.first_button.cget("state")), "normal")
                    self.assertEqual(str(window.previous_button.cget("state")), "normal")
                    self.assertEqual(str(window.next_button.cget("state")), "normal")
                    self.assertEqual(str(window.last_button.cget("state")), "normal")

                    window.goto_page(window.total_pages)
                    self.assertEqual(window.page_label.get(), "Page 3 of 3")
                    self.assertEqual(len(window.grid.get_children()), 5)
                    self.assertEqual(str(window.next_button.cget("state")), "disabled")
                    self.assertEqual(str(window.last_button.cget("state")), "disabled")

                    window.goto_page(1)
                    selected_values = [tuple(str(value) for value in window.grid.item(item, "values")) for item in window.grid.selection()]
                    self.assertIn(first_values, selected_values)
                finally:
                    window.root.destroy()


if __name__ == "__main__":
    unittest.main()
