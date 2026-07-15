import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from attachment_library_service import AttachmentLibraryService


class FakeResponse:
    def __init__(self, payload, status=200): self.payload=payload; self.status=status
    def read(self): return self.payload
    def __enter__(self): return self
    def __exit__(self, *_): pass


class FakeConnection:
    instances = []
    def __init__(self, host, port, timeout): self.host=host; self.port=port; self.timeout=timeout; self.sent=[]; self.headers={}; self.__class__.instances.append(self)
    def putrequest(self, method, path): self.method=method; self.path=path
    def putheader(self, key, value): self.headers[key]=value
    def endheaders(self): pass
    def send(self, data): self.sent.append(data)
    def getresponse(self): return FakeResponse(b'{"id":"new-1"}', 201)
    def close(self): self.closed=True


class AttachmentLibraryTests(unittest.TestCase):
    def test_list_normalizes_direct_and_wrapped_responses(self):
        direct = [{"id":1,"file_name":"a.pdf","file_size":1024,"upload_date":"2026-07-01"}]
        self.assertEqual(AttachmentLibraryService.normalize_list(direct)[0]["file_name"], "a.pdf")
        wrapped = {"attachments":[{"attachment_id":"x","original_file_name":"b.txt","stored_file_name":"generated-123.bin","size":5,"created_at":"today"}]}
        self.assertEqual(AttachmentLibraryService.normalize_list(wrapped)[0], {"id":"x","file_name":"b.txt","file_size":5,"upload_date":"today"})

    def test_refresh_and_delete_use_expected_endpoints(self):
        calls=[]
        def opener(request, timeout):
            calls.append((request.get_method(), request.full_url, timeout))
            payload = b'[]' if request.get_method()=="GET" else b'{"deleted":true}'
            return FakeResponse(payload)
        with tempfile.TemporaryDirectory() as temp:
            service=AttachmentLibraryService("https://server.test",Path(temp)/"logs",request_opener=opener)
            self.assertEqual(service.list_attachments(),[]);service.delete("id/with space")
            self.assertEqual(calls[0][0:2],("GET","https://server.test/api/attachments/list"))
            self.assertEqual(calls[1][0:2],("DELETE","https://server.test/api/attachments/id%2Fwith%20space"))

    def test_upload_streams_multipart_and_reports_progress(self):
        FakeConnection.instances.clear()
        with tempfile.TemporaryDirectory() as temp:
            file_path=Path(temp)/"report.txt";file_path.write_bytes(b"abcdefghij");progress=[]
            service=AttachmentLibraryService("https://server.test",Path(temp)/"logs",connection_factory=FakeConnection)
            result=service.upload(file_path,lambda sent,total:progress.append((sent,total)))
            connection=FakeConnection.instances[0]
            self.assertEqual(result["id"],"new-1");self.assertEqual((connection.method,connection.path),("POST","/api/attachments/upload"))
            body=b"".join(connection.sent);self.assertIn(b'name="file"; filename="report.txt"',body);self.assertIn(b"abcdefghij",body);self.assertEqual(progress[-1],(10,10));self.assertTrue(connection.closed)

    def test_tracking_attachment_registration_posts_expected_payload(self):
        requests=[]
        def opener(request, timeout):
            requests.append(request)
            return FakeResponse(b'{"success":true}', 201)
        with tempfile.TemporaryDirectory() as temp:
            service=AttachmentLibraryService("https://server.test",Path(temp)/"logs",request_opener=opener)
            result=service.register_tracking_attachments("tracking-1",["attachment-1","attachment-2"])
            request=requests[0]
            self.assertTrue(result["success"])
            self.assertEqual(request.get_method(),"POST")
            self.assertEqual(request.full_url,"https://server.test/api/tracking/attachments")
            self.assertEqual(json.loads(request.data),{"tracking_id":"tracking-1","attachment_ids":["attachment-1","attachment-2"]})
            self.assertEqual(request.headers["Content-type"],"application/json")

    def test_tracking_attachment_registration_failure_is_reported(self):
        def failed(*_, **__): raise OSError("mapping API offline")
        with tempfile.TemporaryDirectory() as temp:
            logs=Path(temp)/"logs"
            service=AttachmentLibraryService("https://server.test",logs,request_opener=failed)
            with self.assertRaisesRegex(RuntimeError,"Attachment mapping registration failed"):
                service.register_tracking_attachments("tracking-1",["attachment-1"])
            self.assertFalse(list(logs.glob("*.log")))

    def test_server_failure_is_friendly_without_local_logs(self):
        def failed(*_, **__): raise OSError("offline")
        with tempfile.TemporaryDirectory() as temp:
            logs=Path(temp)/"logs";service=AttachmentLibraryService("https://server.test",logs,request_opener=failed)
            with self.assertRaisesRegex(RuntimeError,"server is unavailable"):service.list_attachments()
            self.assertFalse(list(logs.glob("*.log")))


if __name__=="__main__":unittest.main()
