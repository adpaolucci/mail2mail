#!/usr/bin/env python3
"""
tests/test_jmap.py — Unit & Integration Test Suite for JMAP RFC 8620 & RFC 8621

Verifies:
  1. RFC 8620 JMAP Session Resource discovery & primary account resolution.
  2. RFC 8621 Mailbox/get, Email/query, and Email/get batch invocations.
  3. System keywords ($seen, $flagged, $answered, $draft) to MailRecord flag mapping.
  4. Static JMAP JSON file sniffing and ingestion via mail_formats.iter_mail_records.
  5. End-to-end sync to MBOX and EML formats with incremental delta checkpointing.
  6. mail2jmap CLI subcommand execution.
"""

from __future__ import annotations

import email.message
import http.server
import json
import os
import shutil
import tempfile
import threading
import unittest
from typing import Any, Dict, List, Tuple

import mail2jmap
import mail_formats
from mail2jmap import JmapClient, iter_jmap_records, jmap_item_to_mail_record


class MockJmapHandler(http.server.BaseHTTPRequestHandler):
    """Mock RFC 8620 / RFC 8621 HTTP Server."""

    def log_message(self, format: str, *args: Any) -> None:
        pass  # Suppress console logging during test runs

    def do_GET(self) -> None:
        if self.path == "/jmap/session":
            session_data = {
                "capabilities": {
                    "urn:ietf:params:jmap:core": {
                        "maxSizeUpload": 50000000,
                        "maxObjectsInGet": 500,
                    },
                    "urn:ietf:params:jmap:mail": {
                        "maxMailboxesPerEmail": 50,
                    },
                },
                "accounts": {
                    "acc_fastmail_test": {
                        "name": "tester@fastmail.com",
                        "isPersonal": True,
                    }
                },
                "primaryAccounts": {
                    "urn:ietf:params:jmap:mail": "acc_fastmail_test"
                },
                "apiUrl": f"http://127.0.0.1:{self.server.server_port}/jmap/api",
                "downloadUrl": f"http://127.0.0.1:{self.server.server_port}/jmap/download/{{accountId}}/{{blobId}}/{{name}}",
                "uploadUrl": f"http://127.0.0.1:{self.server.server_port}/jmap/upload/{{accountId}}/",
            }
            body = json.dumps(session_data).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self) -> None:
        if self.path == "/jmap/api":
            content_length = int(self.headers.get("Content-Length", 0))
            raw_body = self.rfile.read(content_length).decode("utf-8")
            req = json.loads(raw_body)
            method_calls = req.get("methodCalls", [])
            responses = []

            for method_name, args, tag in method_calls:
                if method_name == "Mailbox/get":
                    responses.append([
                        "Mailbox/get",
                        {
                            "accountId": "acc_fastmail_test",
                            "state": "state_1",
                            "list": [
                                {
                                    "id": "mb_inbox",
                                    "name": "Inbox",
                                    "totalEmails": 1,
                                    "unreadEmails": 0,
                                }
                            ],
                        },
                        tag,
                    ])
                elif method_name == "Email/query":
                    responses.append([
                        "Email/query",
                        {
                            "accountId": "acc_fastmail_test",
                            "ids": ["jmap_test_001"],
                            "total": 1,
                        },
                        tag,
                    ])
                elif method_name == "Email/get":
                    responses.append([
                        "Email/get",
                        {
                            "accountId": "acc_fastmail_test",
                            "list": [
                                {
                                    "id": "jmap_test_001",
                                    "blobId": "blob_001",
                                    "threadId": "thread_fastmail_01",
                                    "mailboxIds": {"mb_inbox": True},
                                    "keywords": {
                                        "$seen": True,
                                        "$flagged": True,
                                    },
                                    "size": 512,
                                    "receivedAt": "2026-11-07T14:30:00Z",
                                    "sentAt": "2026-11-07T14:29:00Z",
                                    "messageId": ["<mock-jmap-01@fastmail.com>"],
                                    "subject": "Mock JMAP Fastmail Test",
                                    "from": [
                                        {"name": "Rob Mueller", "email": "rob@fastmail.com"}
                                    ],
                                    "to": [
                                        {"name": "Anthony Paolucci", "email": "adpaolucci@fastmail.com"}
                                    ],
                                    "textBody": [
                                        {"partId": "t0", "size": 30, "type": "text/plain"}
                                    ],
                                    "bodyValues": {
                                        "t0": {
                                            "value": "RFC 8620/8621 JMAP Sync Passed!",
                                            "isTruncated": False,
                                        }
                                    },
                                    "attachments": [],
                                }
                            ],
                        },
                        tag,
                    ])

            resp_payload = {
                "using": ["urn:ietf:params:jmap:core", "urn:ietf:params:jmap:mail"],
                "methodResponses": responses,
            }
            body = json.dumps(resp_payload).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_response(404)
            self.end_headers()


class TestJmapRFCCompliance(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.server = http.server.HTTPServer(("127.0.0.1", 0), MockJmapHandler)
        cls.port = cls.server.server_port
        cls.server_thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.server_thread.start()

    @classmethod
    def tearDownClass(cls) -> None:
        cls.server.shutdown()
        cls.server.server_close()

    def setUp(self) -> None:
        self.temp_dir = tempfile.mkdtemp(prefix="test_jmap_")

    def tearDown(self) -> None:
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_jmap_session_discovery_and_account_resolution(self) -> None:
        """Verify RFC 8620 session discovery against mock server."""
        client = JmapClient(f"http://127.0.0.1:{self.port}/jmap/session", token="mock-bearer-token")
        client.connect()

        self.assertEqual(client.account_id, "acc_fastmail_test")
        self.assertTrue(client.session.api_url.endswith("/jmap/api"))
        self.assertIn("urn:ietf:params:jmap:mail", client.session.capabilities)

    def test_jmap_mailbox_and_email_query(self) -> None:
        """Verify RFC 8621 Mailbox/get, Email/query, and Email/get."""
        client = JmapClient(f"http://127.0.0.1:{self.port}/jmap/session", token="mock-bearer-token")
        client.connect()

        mailboxes = client.list_mailboxes()
        self.assertEqual(len(mailboxes), 1)
        self.assertEqual(mailboxes[0]["name"], "Inbox")

        email_ids = client.query_emails(mailboxes[0]["id"])
        self.assertEqual(email_ids, ["jmap_test_001"])

        emails = client.get_emails(email_ids)
        self.assertEqual(len(emails), 1)
        self.assertEqual(emails[0]["subject"], "Mock JMAP Fastmail Test")
        self.assertEqual(emails[0]["from"][0]["email"], "rob@fastmail.com")

    def test_jmap_keyword_to_status_flags_mapping(self) -> None:
        """Verify that RFC 8621 system keywords accurately map to MailRecord boolean flags."""
        sample_item = {
            "id": "msg_flags_test",
            "subject": "Flag Mapping Test",
            "from": [{"name": "Sender", "email": "sender@test.local"}],
            "mailboxIds": {"mb_inbox": True},
            "keywords": {
                "$seen": True,
                "$flagged": True,
                "$answered": True,
                "$draft": True,
            },
            "textBody": [{"partId": "t0"}],
            "bodyValues": {"t0": {"value": "Flags test"}},
        }
        rec = jmap_item_to_mail_record(sample_item)
        self.assertTrue(rec.flags["is_read"])
        self.assertTrue(rec.flags["is_flagged"])
        self.assertTrue(rec.flags["is_replied"])
        self.assertTrue(rec.flags["is_draft"])
        self.assertFalse(rec.flags["is_deleted"])

    def test_static_jmap_file_sniffing_and_ingestion(self) -> None:
        """Verify that samples/jmap/sample_jmap_export.json is sniffed as JMAP and streams cleanly."""
        sample_path = os.path.join(
            os.path.dirname(os.path.dirname(__file__)),
            "samples",
            "jmap",
            "sample_jmap_export.json",
        )
        self.assertTrue(os.path.isfile(sample_path), f"Missing {sample_path}")

        records = list(mail_formats.iter_mail_records(sample_path))
        self.assertEqual(len(records), 2)
        self.assertEqual(records[0].folder_parts, ["Inbox"])
        self.assertIn(b"Project Titan: JMAP RFC 8620 & RFC 8621 Verification", records[0].raw_bytes)
        self.assertTrue(records[0].flags["is_read"])
        self.assertTrue(records[0].flags["is_flagged"])

        self.assertIn(b"Re: Project Titan: JMAP RFC 8620 & RFC 8621 Verification", records[1].raw_bytes)
        self.assertTrue(records[1].flags["is_replied"])

    def test_jmap_live_sync_to_eml_and_mbox_with_checkpoint(self) -> None:
        """Test full live export lifecycle to EML and MBOX with resumable delta sync."""
        session_url = f"http://127.0.0.1:{self.port}/jmap/session"
        out_eml = os.path.join(self.temp_dir, "out_eml")
        chk_file = os.path.join(self.temp_dir, "checkpoint.json")

        # 1. Sync to EML directory
        count = mail2jmap.export_jmap(
            session_url=session_url,
            token="test-token",
            output=out_eml,
            checkpoint_file=chk_file,
        )
        self.assertEqual(count, 1)
        self.assertTrue(os.path.isdir(out_eml))
        self.assertTrue(os.path.isfile(chk_file))

        # Checkpoint validation
        with open(chk_file, "r", encoding="utf-8") as fp:
            chk = json.load(fp)
        self.assertIn("jmap_test_001", chk.get("migrated_message_ids", []))
        self.assertEqual(chk.get("total_messages_migrated"), 1)

        # 2. Re-running with checkpoint skips already migrated item
        out_mbox = os.path.join(self.temp_dir, "out.mbox")
        count2 = mail2jmap.export_jmap(
            session_url=session_url,
            token="test-token",
            output=out_mbox,
            checkpoint_file=chk_file,
        )
        self.assertEqual(count2, 0)  # Skipped because already in checkpoint

    def test_mail2mail_jmap_subcommand_cli(self) -> None:
        """Verify that mail2mail.py jmap ... dispatches properly."""
        import mail2mail

        session_url = f"http://127.0.0.1:{self.port}/jmap/session"
        out_mbox = os.path.join(self.temp_dir, "cli_out.mbox")

        # Invoke mail2mail jmap subcommand
        mail2mail.main([
            "jmap",
            "--session-url", session_url,
            "--token", "mock-token",
            "--out", out_mbox,
        ])

        self.assertTrue(os.path.isfile(out_mbox))
        self.assertGreater(os.path.getsize(out_mbox), 0)

        # Verify records inside output MBOX
        records = list(mail_formats.iter_mail_records(out_mbox))
        self.assertEqual(len(records), 1)
        self.assertIn(b"Mock JMAP Fastmail Test", records[0].raw_bytes)


if __name__ == "__main__":
    unittest.main()
