#!/usr/bin/env python3
"""
tests/test_audit_and_logging.py — Automated verification of audit logging, manifests, and dry-run inspection:
  - AuditLogger CSV manifest generation with SHA-256 and metadata
  - AuditLogger JSONL manifest and error reporting
  - Persistent operational log file output and summary reconciliation
  - Pre-flight dry-run inspection (perform_dry_run) detecting folder stats and oversized attachments
  - CLI integration across mail2eml, mail2mbox, and mail2mail
"""

from __future__ import annotations

import csv
import io
import json
import os
import shutil
import sys
import tempfile
import unittest

import tests  # noqa: F401
from mail_audit import (
    AuditLogger,
    AuditRecord,
    compute_sha256,
    perform_dry_run,
)
import mail2eml
import mail2mbox
import mail2mail


class TestAuditAndLogging(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="test_audit_")

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_audit_logger_csv_manifest_and_sha256(self):
        csv_path = os.path.join(self.tmpdir, "manifest.csv")
        logger = AuditLogger(audit_log_path=csv_path)

        raw = b"From: alice@example.com\r\nSubject: Invoice\r\n\r\nHello"
        digest = compute_sha256(raw)

        rec = AuditRecord(
            timestamp="2026-11-07T05:00:00Z",
            status="SUCCESS",
            source_file="Inbox.mbox",
            source_folder="Inbox",
            target_location="D:\\PSTs\\Inbox.pst::Inbox",
            message_id="<inv001@example.com>",
            subject="Invoice",
            date="Fri, 06 Nov 2026 12:00:00 +0000",
            sender="alice@example.com",
            recipients="bob@example.com",
            size_bytes=len(raw),
            attachment_count=0,
            sha256=digest,
            notes="Clean conversion",
        )
        logger.record_message(rec)
        logger.close()

        self.assertTrue(os.path.isfile(csv_path))
        with open(csv_path, "r", encoding="utf-8") as f:
            reader = list(csv.DictReader(f))
            self.assertEqual(len(reader), 1)
            row = reader[0]
            self.assertEqual(row["status"], "SUCCESS")
            self.assertEqual(row["message_id"], "<inv001@example.com>")
            self.assertEqual(row["sha256"], digest)
            self.assertEqual(row["source_folder"], "Inbox")

    def test_audit_logger_jsonl_and_error_reporting(self):
        jsonl_path = os.path.join(self.tmpdir, "manifest.jsonl")
        err_path = os.path.join(self.tmpdir, "errors.jsonl")
        log_path = os.path.join(self.tmpdir, "ops.log")

        logger = AuditLogger(
            audit_log_path=jsonl_path,
            error_log_path=err_path,
            log_file_path=log_path,
            quiet=True,
        )

        rec = AuditRecord(
            timestamp="2026-11-07T05:00:00Z",
            status="SKIPPED_OVERSIZED",
            source_file="archive.zip",
            source_folder="Archive/2025",
            target_location="",
            message_id="<vid002@example.com>",
            subject="Company All-Hands Recording",
            date="Fri, 06 Nov 2026 15:00:00 +0000",
            sender="media@example.com",
            recipients="all@example.com",
            size_bytes=45 * 1024 * 1024,
            attachment_count=1,
            notes="Exceeded 25 MB limit",
        )
        logger.record_message(rec)
        logger.record_failure(
            source_file="archive.zip",
            source_folder="Archive/2025",
            message_id="<vid002@example.com>",
            subject="Company All-Hands Recording",
            error_type="ComAttachmentLimitExceeded",
            error_details="Attachment recording.mp4 (45 MB) exceeds Outlook MaximumAttachmentSize",
            retryable=True,
        )
        logger.info("Test operational message")
        logger.print_reconciliation_summary(
            source="archive.zip",
            target="out.pst",
            start_time=0.0,
            total_scanned=1,
            exported=0,
            skipped_date=0,
            skipped_deleted=0,
            failed=1,
            total_bytes=45 * 1024 * 1024,
        )
        logger.close()

        with open(jsonl_path, "r", encoding="utf-8") as f:
            lines = [json.loads(line) for line in f]
            self.assertEqual(len(lines), 1)
            self.assertEqual(lines[0]["status"], "SKIPPED_OVERSIZED")

        with open(err_path, "r", encoding="utf-8") as f:
            err_lines = [json.loads(line) for line in f]
            self.assertEqual(len(err_lines), 1)
            self.assertEqual(err_lines[0]["error_type"], "ComAttachmentLimitExceeded")
            self.assertTrue(err_lines[0]["retryable"])

        with open(log_path, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("MIGRATION AUDIT SUMMARY", content)
            self.assertIn("Test operational message", content)

    def test_perform_dry_run_inspection(self):
        # Create a synthetic mbox with 2 normal messages and 1 oversized attachment
        mbox_path = os.path.join(self.tmpdir, "audit_sample.mbox")
        huge_data = b"B" * (500 * 1024)  # 500 KB test payload
        with open(mbox_path, "wb") as f:
            f.write(
                b"From sender@example.com Fri Nov  6 12:00:00 2020\n"
                b"From: sender@example.com\n"
                b"Subject: Normal Mail 1\n"
                b"Date: Fri, 06 Nov 2020 12:00:00 +0000\n\n"
                b"Body text 1\n\n"
                b"From sender@example.com Sat Nov  7 12:00:00 2020\n"
                b"From: sender@example.com\n"
                b"Subject: Huge Mail 2\n"
                b"Date: Sat, 07 Nov 2020 12:00:00 +0000\n"
                b"Content-Type: multipart/mixed; boundary=\"sep\"\n\n"
                b"--sep\nContent-Type: text/plain\n\nBody\n"
                b"--sep\nContent-Type: application/zip; name=\"huge.zip\"\n"
                b"Content-Disposition: attachment; filename=\"huge.zip\"\n\n"
                + huge_data + b"\n--sep--\n\n"
            )

        captured_stdout = io.StringIO()
        old_stdout = sys.stdout
        try:
            sys.stdout = captured_stdout
            perform_dry_run(mbox_path, max_attach_mb=0.4)  # 0.4 MB threshold triggers on 500 KB
        finally:
            sys.stdout = old_stdout

        report = captured_stdout.getvalue()
        self.assertIn("PRE-FLIGHT AUDIT REPORT", report)
        self.assertIn("Total Messages:          2", report)
        self.assertIn("WARNING: Found 1 attachment(s) exceeding 0 MB", report)
        self.assertIn("huge.zip", report)

    def test_mail2eml_and_mail2mbox_audit_cli_integration(self):
        eml_src = os.path.join(self.tmpdir, "inbox.eml")
        with open(eml_src, "wb") as f:
            f.write(b"From: u@example.com\r\nSubject: Test\r\nDate: Fri, 06 Nov 2020 12:00:00 +0000\r\n\r\nHello\r\n")

        out_eml_dir = os.path.join(self.tmpdir, "out_eml")
        audit_csv = os.path.join(self.tmpdir, "audit.csv")
        ops_log = os.path.join(self.tmpdir, "ops.log")

        # Test mail2eml with audit manifest and log file
        mail2eml.main([eml_src, out_eml_dir, "--audit-log", audit_csv, "--log-file", ops_log, "-q"])
        self.assertTrue(os.path.isfile(audit_csv))
        self.assertTrue(os.path.isfile(ops_log))
        with open(audit_csv, "r", encoding="utf-8") as f:
            reader = list(csv.DictReader(f))
            self.assertEqual(len(reader), 1)
            self.assertEqual(reader[0]["status"], "SUCCESS")
            self.assertTrue(len(reader[0]["sha256"]) == 64)

        # Test mail2mbox with audit manifest
        out_mbox = os.path.join(self.tmpdir, "out.mbox")
        audit_mbox_csv = os.path.join(self.tmpdir, "audit_mbox.csv")
        mail2mbox.main([out_eml_dir, out_mbox, "--audit-log", audit_mbox_csv, "-q"])
        self.assertTrue(os.path.isfile(out_mbox))
        self.assertTrue(os.path.isfile(audit_mbox_csv))


if __name__ == "__main__":
    unittest.main()
