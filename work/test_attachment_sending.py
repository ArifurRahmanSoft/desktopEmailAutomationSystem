import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import email_automation as app


class FakeLibrary:
    def __init__(self, attachments): self.attachments=attachments; self.calls=0; self.registrations=[]
    def list_attachments(self): self.calls+=1; return self.attachments
    def register_tracking_attachments(self, tracking_id, attachment_ids): self.registrations.append((tracking_id,attachment_ids))


class AttachmentSendingTests(unittest.TestCase):
    def test_no_selection_skips_attachment_api(self):
        with patch.object(app, "attachment_library_service") as service:
            self.assertEqual(app.resolve_selected_attachments([]), [])
            service.assert_not_called()

    def test_selected_attachments_are_revalidated_and_order_preserved(self):
        library=FakeLibrary([{"id":"a1","file_name":"Proposal.pdf"},{"id":"a2","file_name":"Profile.pdf"}])
        with patch.object(app,"attachment_library_service",return_value=library):
            result=app.resolve_selected_attachments(["a2","a1"])
        self.assertEqual([x["id"] for x in result],["a2","a1"]);self.assertEqual(library.calls,1)

    def test_deleted_attachment_blocks_send_with_friendly_error(self):
        library=FakeLibrary([{"id":"a1","file_name":"Proposal.pdf"}])
        with patch.object(app,"attachment_library_service",return_value=library):
            with self.assertRaisesRegex(ValueError,"deleted from the server"):app.resolve_selected_attachments(["missing"])

    def test_download_links_use_same_tracking_id_and_original_names(self):
        attachments=[{"id":"a/1","file_name":"Proposal & Terms.pdf"},{"id":"a2","file_name":"CompanyProfile.pdf"}]
        html,urls=app.build_attachment_links_html("tracking-one",attachments)
        self.assertEqual(len(urls),2);self.assertTrue(all("/download/tracking-one/" in url for url in urls));self.assertIn("a%2F1",urls[0]);self.assertIn(f'<a href="{urls[0]}">Proposal &amp; Terms.pdf</a>',html);self.assertIn(f'<a href="{urls[1]}">CompanyProfile.pdf</a>',html);self.assertNotIn(f'>{urls[0]}</a>',html)
        _html2,urls2=app.build_attachment_links_html("tracking-two",attachments);self.assertNotEqual(urls,urls2)

    def test_no_attachments_skips_mapping_api(self):
        with patch.object(app,"attachment_library_service") as service:
            app.register_attachment_mapping("tracking-1",[])
            service.assert_not_called()

    def test_selected_attachments_register_one_mapping_with_all_ids(self):
        library=FakeLibrary([])
        attachments=[{"id":"a1","file_name":"Proposal.pdf"},{"id":"a2","file_name":"Profile.pdf"}]
        with patch.object(app,"attachment_library_service",return_value=library):
            app.register_attachment_mapping("tracking-1",attachments)
        self.assertEqual(library.registrations,[("tracking-1",["a1","a2"])])

    def test_mapping_api_failure_stops_attachment_flow(self):
        library=FakeLibrary([])
        library.register_tracking_attachments=lambda *_: (_ for _ in ()).throw(RuntimeError("mapping rejected"))
        with patch.object(app,"attachment_library_service",return_value=library):
            with self.assertRaisesRegex(RuntimeError,"mapping rejected"):
                app.register_attachment_mapping("tracking-1",[{"id":"a1","file_name":"Proposal.pdf"}])

if __name__=="__main__":unittest.main()
