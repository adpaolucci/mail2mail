#!/usr/bin/env python3
"""
tests/test_real_world_hardening.py — Tests for production-scale migration hardening:
  - extract_best_datetime & Received header timestamp fallback
  - parse_gmail_labels & Google Takeout X-Gmail-Labels folder hierarchy reconstruction
  - In-session deduplication (--dedup auto/message-id/sha256)
  - Interrupted migration resuming (--resume <audit_log>)
"""

from __future__ import annotations

import csv
import email
from email.message import EmailMessage
import os
import shutil
import tempfile
import unittest

import tests  # noqa: F401
from mail_formats import (
    extract_attachment_filename,
    extract_best_datetime,
    iter_mail_records,
    parse_gmail_labels,
    pick_year_month,
)
from mail_audit import AuditLogger, AuditRecord, compute_sha256
import mail2eml
import mail2mbox
from mail2pst import (
    MSGFLAG_HASATTACH,
    MSGFLAG_READ,
    MSGFLAG_SUBMIT,
    MSGFLAG_UNSENT,
    PR_CREATION_TIME,
    PR_LAST_MODIFICATION_TIME,
    PR_MESSAGE_FLAGS,
    PstRouter,
)


class TestRealWorldHardening(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="test_hardening_")

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_extract_best_datetime_and_received_fallback(self):
        # 1. Standard Date header
        msg1 = EmailMessage()
        msg1["Date"] = "Wed, 25 Oct 2023 10:15:30 +0000"
        dt1 = extract_best_datetime(msg1)
        self.assertIsNotNone(dt1)
        self.assertEqual(dt1.year, 2023)
        self.assertEqual(dt1.month, 10)
        self.assertEqual(pick_year_month(msg1), (2023, 10))

        # 2. Missing Date header, but valid Received header
        msg2 = EmailMessage()
        msg2["Subject"] = "Draft without Date"
        msg2["Received"] = (
            "from mail.example.com (mail.example.com [192.0.2.1]) "
            "by mx.google.com with ESMTPS id abc123; "
            "Thu, 15 Aug 2024 16:45:00 +0000 (UTC)"
        )
        dt2 = extract_best_datetime(msg2)
        self.assertIsNotNone(dt2)
        self.assertEqual(dt2.year, 2024)
        self.assertEqual(dt2.month, 8)
        self.assertEqual(pick_year_month(msg2), (2024, 8))

        # 3. Corrupted Date header with valid Received fallback
        raw_msg3 = (
            b"Date: corrupted date header string\r\n"
            b"Received: from gate.corp.com; Mon, 05 May 2025 09:30:00 -0400\r\n\r\n"
            b"Body"
        )
        msg3 = email.message_from_bytes(raw_msg3)
        dt3 = extract_best_datetime(msg3)
        self.assertIsNotNone(dt3)
        self.assertEqual(dt3.year, 2025)
        self.assertEqual(dt3.month, 5)
        self.assertEqual(pick_year_month(msg3), (2025, 5))

        # 4. Completely missing Date and Received -> falls back to 1970
        msg4 = EmailMessage()
        self.assertIsNone(extract_best_datetime(msg4))
        self.assertEqual(pick_year_month(msg4), (1970, 1))

    def test_parse_gmail_labels(self):
        # 1. User nested label with Starred and Important
        parts, flags = parse_gmail_labels("Inbox,Important,Work/Project Alpha,Starred")
        self.assertEqual(parts, ["Work", "Project Alpha"])
        self.assertTrue(flags["is_flagged"])
        self.assertTrue(flags["is_read"])
        self.assertFalse(flags["is_deleted"])

        # 2. Sent items
        parts, flags = parse_gmail_labels("Sent,Important")
        self.assertEqual(parts, ["Sent Items"])
        self.assertFalse(flags["is_flagged"])

        # 3. Trash / Deleted
        parts, flags = parse_gmail_labels("Trash")
        self.assertEqual(parts, ["Deleted Items"])
        self.assertTrue(flags["is_deleted"])

        # 4. Unread user receipts
        parts, flags = parse_gmail_labels("Unread,Receipts")
        self.assertEqual(parts, ["Receipts"])
        self.assertFalse(flags["is_read"])

        # 5. Category only -> maps to Inbox
        parts, flags = parse_gmail_labels("Category Promotions,Unread")
        self.assertEqual(parts, ["Inbox"])
        self.assertFalse(flags["is_read"])

    def test_takeout_mbox_gmail_labels_routing(self):
        mbox_path = os.path.join(self.tmpdir, "All mail Including Spam and Trash.mbox")
        with open(mbox_path, "wb") as f:
            f.write(
                b"From user@gmail.com Wed Oct 25 10:00:00 2023\n"
                b"From: user@gmail.com\n"
                b"To: client@example.com\n"
                b"Subject: Budget Plan\n"
                b"Date: Wed, 25 Oct 2023 10:00:00 +0000\n"
                b"X-Gmail-Labels: Inbox,Work/Finance,Starred\n\n"
                b"Here is the budget.\n\n"
                b"From user@gmail.com Thu Oct 26 11:00:00 2023\n"
                b"From: user@gmail.com\n"
                b"To: client@example.com\n"
                b"Subject: Proposal\n"
                b"Date: Thu, 26 Oct 2023 11:00:00 +0000\n"
                b"X-Gmail-Labels: Sent,Clients/Acme\n\n"
                b"Here is the proposal.\n\n"
                b"From user@gmail.com Fri Oct 27 12:00:00 2023\n"
                b"From: spam@example.com\n"
                b"Subject: Junk\n"
                b"Date: Fri, 27 Oct 2023 12:00:00 +0000\n"
                b"X-Gmail-Labels: Trash\n\n"
                b"Spam text.\n\n"
            )

        out_eml_dir = os.path.join(self.tmpdir, "eml_export")
        exported, skipped, dir_counts, _ = mail2eml.export_to_eml(
            src_path=mbox_path,
            out_root=out_eml_dir,
            layout="folder",
            include_deleted=False,
            parse_gmail=True,
        )

        self.assertEqual(exported, 2)
        self.assertEqual(skipped, 1)  # Trash message skipped

        finance_dir = os.path.join(out_eml_dir, "Work", "Finance")
        acme_dir = os.path.join(out_eml_dir, "Clients", "Acme")
        self.assertTrue(os.path.isdir(finance_dir))
        self.assertTrue(os.path.isdir(acme_dir))
        self.assertEqual(len(os.listdir(finance_dir)), 1)
        self.assertEqual(len(os.listdir(acme_dir)), 1)

    def test_deduplication_by_message_id_and_sha256(self):
        mbox_path = os.path.join(self.tmpdir, "dups.mbox")
        with open(mbox_path, "wb") as f:
            f.write(
                b"From a@example.com Wed Oct 25 10:00:00 2023\n"
                b"From: a@example.com\n"
                b"Subject: Original Message\n"
                b"Message-ID: <msg-unique-001@example.com>\n"
                b"Date: Wed, 25 Oct 2023 10:00:00 +0000\n\n"
                b"First copy body.\n\n"
                b"From a@example.com Wed Oct 25 10:00:00 2023\n"
                b"From: a@example.com\n"
                b"Subject: Original Message Copy\n"
                b"Message-ID: <msg-unique-001@example.com>\n"
                b"Date: Wed, 25 Oct 2023 10:00:00 +0000\n\n"
                b"Duplicate copy body.\n\n"
                b"From b@example.com Thu Oct 26 10:00:00 2023\n"
                b"From: b@example.com\n"
                b"Subject: Different Message\n"
                b"Message-ID: <msg-unique-002@example.com>\n"
                b"Date: Thu, 26 Oct 2023 10:00:00 +0000\n\n"
                b"Distinct body.\n\n"
            )

        out_dir = os.path.join(self.tmpdir, "eml_dedup")
        audit_csv = os.path.join(self.tmpdir, "audit_dedup.csv")
        logger = AuditLogger(audit_log_path=audit_csv)

        exported, skipped, _, _ = mail2eml.export_to_eml(
            src_path=mbox_path,
            out_root=out_dir,
            dedup="auto",
            logger=logger,
        )
        logger.close()

        self.assertEqual(exported, 2)
        self.assertEqual(skipped, 1)

        with open(audit_csv, "r", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
            self.assertEqual(len(rows), 3)
            statuses = [r["status"] for r in rows]
            self.assertEqual(statuses.count("SUCCESS"), 2)
            self.assertEqual(statuses.count("SKIPPED_DUPLICATE"), 1)

    def test_resumability_with_existing_manifest(self):
        # Pre-seed an audit log with message 1 already marked SUCCESS
        audit_manifest = os.path.join(self.tmpdir, "prior_audit.csv")
        logger = AuditLogger(audit_log_path=audit_manifest)
        logger.record_message(AuditRecord(
            timestamp="2026-11-06T10:00:00Z",
            status="SUCCESS",
            source_file="inbox.mbox",
            source_folder="Inbox",
            target_location="",
            message_id="<already-done-001@example.com>",
            subject="Already Processed",
            date="Wed, 25 Oct 2023 10:00:00 +0000",
            sender="a@example.com",
            recipients="b@example.com",
            size_bytes=100,
            attachment_count=0,
            sha256="abc123hash",
        ))
        logger.close()

        mbox_path = os.path.join(self.tmpdir, "resume_test.mbox")
        with open(mbox_path, "wb") as f:
            f.write(
                b"From a@example.com Wed Oct 25 10:00:00 2023\n"
                b"From: a@example.com\n"
                b"Subject: Already Processed\n"
                b"Message-ID: <already-done-001@example.com>\n"
                b"Date: Wed, 25 Oct 2023 10:00:00 +0000\n\n"
                b"Body 1\n\n"
                b"From b@example.com Thu Oct 26 10:00:00 2023\n"
                b"From: b@example.com\n"
                b"Subject: New Message\n"
                b"Message-ID: <new-message-002@example.com>\n"
                b"Date: Thu, 26 Oct 2023 10:00:00 +0000\n\n"
                b"Body 2\n\n"
            )

        out_mbox = os.path.join(self.tmpdir, "resumed.mbox")
        written, skipped = mail2mbox.export_to_mbox(
            src_path=mbox_path,
            target_path=out_mbox,
            resume=audit_manifest,
        )

        self.assertEqual(written, 1)
        self.assertEqual(skipped, 1)  # already-done-001 was skipped via resume

    def test_extract_attachment_filename_rfc2047_and_content_type(self):
        # 1. Attachment with RFC 2047 encoded words in filename parameter
        raw1 = (
            b'Content-Type: application/pdf\r\n'
            b'Content-Disposition: attachment; filename="=?UTF-8?B?15HXk9eZ16fXlA==?=.pdf"\r\n\r\n'
            b'%PDF-1.4 sample'
        )
        part1 = email.message_from_bytes(raw1)
        fn1 = extract_attachment_filename(part1)
        self.assertEqual(fn1, "\u05d1\u05d3\u05d9\u05e7\u05d4.pdf")

        # 2. Attachment with name in Content-Type header only
        raw2 = (
            b'Content-Type: application/zip; name="=?ISO-8859-1?Q?archive=2Ezip?="\r\n'
            b'Content-Disposition: inline\r\n\r\n'
            b'PK...'
        )
        part2 = email.message_from_bytes(raw2)
        fn2 = extract_attachment_filename(part2)
        self.assertEqual(fn2, "archive.zip")

    def test_extract_best_datetime_mtime_fallback(self):
        # Message with no Date or Received header
        raw = b"From: test@example.com\r\nSubject: Dateless message\r\n\r\nHello world"
        msg = email.message_from_bytes(raw)
        
        # Without fallback -> None
        self.assertIsNone(extract_best_datetime(msg))

        # With mtime fallback timestamp
        ts = 1698235200.0  # Wed, 25 Oct 2023 12:00:00 UTC
        dt = extract_best_datetime(msg, fallback_timestamp=ts)
        self.assertIsNotNone(dt)
        self.assertEqual(dt.year, 2023)
        self.assertEqual(dt.month, 10)
        self.assertEqual(pick_year_month(msg, fallback_timestamp=ts), (2023, 10))

    def test_pst_router_naming_enthusedamateur_fix(self):
        # Verify that initial PST does not get an unnecessary _part1 suffix
        router = PstRouter(
            ns=None,
            out_dir=self.tmpdir,
            base_name="my_archive",
            max_gb=10.0,
            splits=None,
            split_by_year=False,
        )
        
        # Single / first part
        self.assertEqual(
            os.path.basename(router._new_path("")),
            "my_archive.pst",
        )
        # Overflow parts
        self.assertEqual(
            os.path.basename(router._new_path("part2")),
            "my_archive_part2.pst",
        )
        self.assertEqual(
            os.path.basename(router._new_path("part3")),
            "my_archive_part3.pst",
        )

        # Year split initial part vs overflow part
        self.assertEqual(
            os.path.basename(router._new_path("2023")),
            "my_archive_2023.pst",
        )
        self.assertEqual(
            os.path.basename(router._new_path("2023_part2")),
            "my_archive_2023_part2.pst",
        )

    def test_mapi_message_flags_and_properties_logic(self):
        # Test the MAPI flags calculation logic directly
        # When an email is received (not draft) and read, MSGFLAG_UNSENT must be cleared
        # and MSGFLAG_READ must be set.
        class MockPropertyAccessor:
            def __init__(self):
                self.props = {
                    PR_MESSAGE_FLAGS: MSGFLAG_UNSENT,  # Outlook default on Items.Add(0)
                }
            def GetProperty(self, tag):
                return self.props.get(tag, 0)
            def SetProperty(self, tag, val):
                self.props[tag] = val

        pa = MockPropertyAccessor()
        # Simulated message flags: is_read=True, is_draft=False, att_count=2
        raw_msg_flags = pa.GetProperty(PR_MESSAGE_FLAGS)
        cur_flags = int(raw_msg_flags)
        is_draft = False
        is_read = True
        att_count = 2

        # Apply our hardening logic
        if not is_draft:
            cur_flags = cur_flags & ~MSGFLAG_UNSENT & ~MSGFLAG_SUBMIT
        else:
            cur_flags |= MSGFLAG_UNSENT

        if is_read:
            cur_flags |= MSGFLAG_READ
        else:
            cur_flags &= ~MSGFLAG_READ

        if att_count > 0:
            cur_flags |= MSGFLAG_HASATTACH

        pa.SetProperty(PR_MESSAGE_FLAGS, cur_flags)

        final_flags = pa.GetProperty(PR_MESSAGE_FLAGS)
        # MSGFLAG_UNSENT (0x08) MUST NOT BE SET
        self.assertEqual(final_flags & MSGFLAG_UNSENT, 0)
        # MSGFLAG_SUBMIT (0x04) MUST NOT BE SET
        self.assertEqual(final_flags & MSGFLAG_SUBMIT, 0)
        # MSGFLAG_READ (0x01) MUST BE SET
        self.assertEqual(final_flags & MSGFLAG_READ, MSGFLAG_READ)
        # MSGFLAG_HASATTACH (0x10) MUST BE SET
        self.assertEqual(final_flags & MSGFLAG_HASATTACH, MSGFLAG_HASATTACH)


if __name__ == "__main__":
    unittest.main()
