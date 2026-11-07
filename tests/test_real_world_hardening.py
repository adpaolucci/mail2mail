#!/usr/bin/env python3
"""
tests/test_real_world_hardening.py — Tests for production-scale migration hardening:
  - extract_best_datetime & Received header timestamp fallback
  - parse_gmail_labels & Google Takeout X-Gmail-Labels folder hierarchy reconstruction
  - In-session deduplication (--dedup auto/message-id/sha256)
  - Interrupted migration resuming (--resume <audit_log>)
"""

from __future__ import annotations

import contextlib
import csv
import email
from email.message import EmailMessage
import io
import os
import shutil
import tempfile
import unittest

import gzip
import tests  # noqa: F401
from datetime import datetime, timezone
from mail_formats import (
    _iter_records_from_bytes,
    extract_attachment_filename,
    extract_best_datetime,
    inject_mozilla_status_bytes,
    iter_mail_records,
    normalize_eml_bytes_with_flags,
    parse_gmail_labels,
    parse_msg_bytes,
    pick_year_month,
    sanitize_name,
)
from mail_audit import AuditLogger, AuditRecord, compute_sha256, perform_dry_run
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
    add_attachments,
    to_mapi_utc,
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

    def test_fuzz_hardening_leap_year_clamping(self):
        # 1. to_mapi_utc with leap day before 1970
        dt_pre1970 = datetime(1968, 2, 29, 12, 0, 0, tzinfo=timezone.utc)
        mapi_dt1 = to_mapi_utc(None, dt_pre1970)
        self.assertEqual(mapi_dt1.year, 1970)
        self.assertEqual(mapi_dt1.month, 2)
        self.assertEqual(mapi_dt1.day, 28)

        # 2. to_mapi_utc with leap day after 2100
        dt_post2100 = datetime(2400, 2, 29, 12, 0, 0, tzinfo=timezone.utc)
        mapi_dt2 = to_mapi_utc(None, dt_post2100)
        self.assertEqual(mapi_dt2.year, 2100)
        self.assertEqual(mapi_dt2.month, 2)
        self.assertEqual(mapi_dt2.day, 28)

        # 3. mail2mbox.build_mbox_from_line with leap days
        from_line1 = mail2mbox.build_mbox_from_line(None, "Thu, 29 Feb 1888 12:00:00 +0000")
        self.assertIn(b"1970", from_line1)
        self.assertIn(b"Feb 28", from_line1)

        from_line2 = mail2mbox.build_mbox_from_line(None, "Sat, 29 Feb 2400 12:00:00 +0000")
        self.assertIn(b"2100", from_line2)
        self.assertIn(b"Feb 28", from_line2)

    def test_fuzz_hardening_null_flags_handling(self):
        raw = b"From: user@example.com\r\nSubject: Test\r\n\r\nMessage body"
        # normalize_eml_bytes_with_flags with flags=None
        norm = normalize_eml_bytes_with_flags(raw, None)
        self.assertIn(b"Subject: Test", norm)
        self.assertIn(b"Status: RO", norm)

        # inject_mozilla_status_bytes with flags=None
        injected = inject_mozilla_status_bytes(raw, None)
        self.assertIn(b"X-Mozilla-Status: 0001", injected)

        # format_mboxrd_entry with flags=None and msg=None
        entry = mail2mbox.format_mboxrd_entry(raw)
        self.assertTrue(entry.startswith(b"From "))
        self.assertIn(b"Status: RO", entry)

    def test_fuzz_hardening_mixed_timezone_audit(self):
        # Create temporary mbox with a mix of timezone-aware and naive Date headers
        mbox_path = os.path.join(self.tmpdir, "mixed_tz.mbox")
        with open(mbox_path, "wb") as f:
            f.write(b"From a@b Wed Oct 25 10:00:00 2023\n")
            f.write(b"Date: Wed, 25 Oct 2023 10:00:00 +0000\n")
            f.write(b"Subject: Aware Date\n\nMsg1\n\n")
            f.write(b"From a@b Thu Oct 26 10:00:00 2023\n")
            f.write(b"Date: 26 Oct 2023 11:00:00 -0000\n")  # Often parses as naive
            f.write(b"Subject: Naive Date\n\nMsg2\n\n")
            f.write(b"From a@b Fri Oct 27 10:00:00 2023\n")
            f.write(b"Date: 27 Oct 2023 12:00:00\n")  # Missing timezone offset
            f.write(b"Subject: Missing Offset\n\nMsg3\n\n")

        # perform_dry_run should safely compute earliest/latest dates without TypeError
        with io.StringIO() as buf, contextlib.redirect_stdout(buf):
            perform_dry_run(mbox_path)
            output = buf.getvalue()
        self.assertIn("Total Messages:          3", output)
        self.assertIn("Wed, 25 Oct 2023 10:00:00 +0000", output)
        self.assertIn("27 Oct 2023 12:00:00", output)

    def test_fuzz_hardening_corrupted_msg_bytes(self):
        # parse_msg_bytes with arbitrary garbage
        garbage = b"\x00\xff\xfe\x12\x34\x56\x78" * 50
        raw, flags = parse_msg_bytes(garbage, strict=False)
        self.assertEqual(raw, b"")
        self.assertIsInstance(flags, dict)

        # parse_msg_bytes with strict=True raises ValueError
        with self.assertRaises(ValueError):
            parse_msg_bytes(garbage, strict=True)

    def test_fuzz_hardening_sanitize_name_strict_length(self):
        for max_len in range(1, 20):
            res = sanitize_name("very_long_file_name_with_an_extension.tar.gz", max_len=max_len)
            self.assertLessEqual(len(res), max_len, f"Failed for max_len={max_len}: {res}")

    def test_zero_data_loss_nested_compressed_archive_stream(self):
        # Create an in-memory mbox.gz with 2 messages
        mbox_data = (
            b"From user1@example.com Wed Oct 25 10:00:00 2023\n"
            b"Subject: First Nested Message\n\nBody 1\n\n"
            b"From user2@example.com Wed Oct 25 11:00:00 2023\n"
            b"Subject: Second Nested Message\n\nBody 2\n\n"
        )
        gz_buf = io.BytesIO()
        with gzip.GzipFile(fileobj=gz_buf, mode="wb") as gz:
            gz.write(mbox_data)
        gz_bytes = gz_buf.getvalue()

        records = list(_iter_records_from_bytes(gz_bytes, ["ArchiveFolder"], "nested.mbox.gz"))
        self.assertEqual(len(records), 2, "Both messages inside nested mbox.gz must be extracted without loss")
        self.assertIn(b"Subject: First Nested Message", records[0].raw_bytes)
        self.assertIn(b"Subject: Second Nested Message", records[1].raw_bytes)

    def test_zero_data_loss_dedup_auto_retains_different_messages_with_same_mid(self):
        # Two distinct messages sharing Message-ID: <shared-id@example.com>
        # (e.g. from broken bot or automated notifications)
        msg1_bytes = (
            b"From: sender1@example.com\r\n"
            b"Message-ID: <shared-id@example.com>\r\n"
            b"Subject: Important Contract Document\r\n"
            b"Date: Wed, 25 Oct 2023 10:00:00 +0000\r\n\r\n"
            b"Contract terms..."
        )
        msg2_bytes = (
            b"From: sender2@example.com\r\n"
            b"Message-ID: <shared-id@example.com>\r\n"
            b"Subject: Urgent Server Security Alert\r\n"
            b"Date: Thu, 26 Oct 2023 12:00:00 +0000\r\n\r\n"
            b"Alert details..."
        )
        src_mbox = os.path.join(self.tmpdir, "shared_mid.mbox")
        with open(src_mbox, "wb") as f:
            f.write(b"From a@b Wed Oct 25 10:00:00 2023\n" + msg1_bytes + b"\n\n")
            f.write(b"From a@b Thu Oct 26 12:00:00 2023\n" + msg2_bytes + b"\n\n")

        # Export with --dedup auto to EML
        eml_out = os.path.join(self.tmpdir, "eml_out_shared_mid")
        exported, skipped, _, _ = mail2eml.export_to_eml(src_mbox, eml_out, dedup="auto")
        # BOTH messages must be retained since their subject/dates differ completely
        self.assertEqual(exported, 2, "Distinct messages with colliding Message-ID must both be retained")
        self.assertEqual(skipped, 0)

        # Export with --dedup auto to MBOX
        mbox_out = os.path.join(self.tmpdir, "mbox_out_shared_mid.mbox")
        written, skipped_m = mail2mbox.export_to_mbox(src_mbox, mbox_out, dedup="auto")
        self.assertEqual(written, 2, "Distinct messages with colliding Message-ID must both be retained in MBOX")
        self.assertEqual(skipped_m, 0)

    def test_zero_data_loss_eml_collision_prevention(self):
        # Create an existing file in destination folder
        dest_dir = os.path.join(self.tmpdir, "eml_collision_dest", "new_message")
        os.makedirs(dest_dir, exist_ok=True)
        existing_file = os.path.join(dest_dir, "000001.eml")
        with open(existing_file, "wb") as f:
            f.write(b"Original Existing EML Content")

        # Export a new message with naming='index'
        src_eml = os.path.join(self.tmpdir, "new_message.eml")
        with open(src_eml, "wb") as f:
            f.write(b"From: test@example.com\r\nSubject: New Msg\r\n\r\nNew body content")

        out_root = os.path.join(self.tmpdir, "eml_collision_dest")
        exported, _, _, _ = mail2eml.export_to_eml(src_eml, out_root, naming="index")
        self.assertEqual(exported, 1)

        # Verify existing file was NOT overwritten
        with open(existing_file, "rb") as f:
            self.assertEqual(f.read(), b"Original Existing EML Content")

        # Verify new file was written with clean incremented index
        new_file = os.path.join(dest_dir, "000002.eml")
        self.assertTrue(os.path.exists(new_file), "New file must be saved under clean next index")
        with open(new_file, "rb") as f:
            self.assertIn(b"New body content", f.read())

    def test_zero_data_loss_unattached_files_preserved_on_disk(self):
        # Test add_attachments emergency preservation to disk when COM raises
        raw = (
            b"Content-Type: multipart/mixed; boundary=\"BOUNDARY\"\r\n\r\n"
            b"--BOUNDARY\r\n"
            b"Content-Type: text/plain\r\n\r\n"
            b"Body\r\n"
            b"--BOUNDARY\r\n"
            b"Content-Type: application/octet-stream; name=\"large_report.pdf\"\r\n"
            b"Content-Disposition: attachment; filename=\"large_report.pdf\"\r\n\r\n"
            b"PAYLOAD_BYTES_HERE_FOR_EMERGENCY_PRESERVATION\r\n"
            b"--BOUNDARY--"
        )
        parsed = email.message_from_bytes(raw)
        
        # Mock mail object whose Attachments.Add raises a simulated COM error
        class MockAttachments:
            def Add(self, *args, **kwargs):
                raise RuntimeError("Simulated Outlook 4096: The attachment size exceeds the allowable limit.")

        class MockMail:
            Attachments = MockAttachments()

        out_dir = os.path.join(self.tmpdir, "pst_output_test")
        tmp_work = os.path.join(self.tmpdir, "work_tmp")
        os.makedirs(out_dir, exist_ok=True)
        os.makedirs(tmp_work, exist_ok=True)

        skipped_oversized = []
        count, has_cid = add_attachments(
            MockMail(),
            parsed,
            tmpdir=tmp_work,
            skipped_oversized=skipped_oversized,
            out_dir=out_dir,
        )
        self.assertEqual(len(skipped_oversized), 1)

        # Verify attachment was safely preserved on disk in _unattached_files
        unattached_dir = os.path.join(out_dir, "_unattached_files")
        self.assertTrue(os.path.isdir(unattached_dir))
        files = os.listdir(unattached_dir)
        self.assertIn("large_report.pdf", files)
        with open(os.path.join(unattached_dir, "large_report.pdf"), "rb") as f:
            self.assertEqual(f.read(), b"PAYLOAD_BYTES_HERE_FOR_EMERGENCY_PRESERVATION")


if __name__ == "__main__":
    unittest.main()
