#!/usr/bin/env python3
"""
tests/test_advanced_features.py

Comprehensive tests for high-value enterprise features:
  1. Date-Range Filtering (--since / --after, --until / --before)
  2. Folder Inclusion and Exclusion Rules (--include-folder, --exclude-folder)
  3. Human-Friendly Size Parsing & Splitting (parse_size_bytes, --split-size)
  4. Persistent Incremental / Delta Migration State Cache (--state-file)
"""

import datetime
import os
import shutil
import tempfile
import unittest

import tests  # Ensure repo root is at sys.path[0]

import mail_filter
from mail_audit import AuditLogger
import mail2eml
import mail2mbox


class TestAdvancedFeatures(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="test_adv_features_")

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_parse_size_bytes(self):
        # Human strings
        self.assertEqual(mail_filter.parse_size_bytes("20GB"), 20 * 1024**3)
        self.assertEqual(mail_filter.parse_size_bytes("20 gb"), 20 * 1024**3)
        self.assertEqual(mail_filter.parse_size_bytes("500MB"), 500 * 1024**2)
        self.assertEqual(mail_filter.parse_size_bytes("1.5G"), int(1.5 * 1024**3))
        self.assertEqual(mail_filter.parse_size_bytes("100KB"), 100 * 1024)
        self.assertEqual(mail_filter.parse_size_bytes("1024B"), 1024)
        # Numbers
        self.assertEqual(mail_filter.parse_size_bytes(15.0), 15 * 1024**3)
        self.assertEqual(mail_filter.parse_size_bytes(0), 0)
        self.assertEqual(mail_filter.parse_size_bytes("unlimited"), 0)

    def test_date_filter_iso_and_rfc(self):
        df = mail_filter.DateFilter(since="2023-01-01", until="2023-12-31")
        # In range
        dt_mid = datetime.datetime(2023, 6, 15, 12, 0, 0, tzinfo=datetime.timezone.utc)
        self.assertTrue(df.matches(dt_mid))
        # Exact boundary: 2023-01-01 00:00:00 UTC
        dt_start = datetime.datetime(2023, 1, 1, 0, 0, 0, tzinfo=datetime.timezone.utc)
        self.assertTrue(df.matches(dt_start))
        # Exact boundary: 2023-12-31 23:59:59 UTC
        dt_end = datetime.datetime(2023, 12, 31, 23, 59, 59, tzinfo=datetime.timezone.utc)
        self.assertTrue(df.matches(dt_end))
        # Before range
        dt_before = datetime.datetime(2022, 12, 31, 23, 59, 59, tzinfo=datetime.timezone.utc)
        self.assertFalse(df.matches(dt_before))
        # After range
        dt_after = datetime.datetime(2024, 1, 1, 0, 0, 1, tzinfo=datetime.timezone.utc)
        self.assertFalse(df.matches(dt_after))

    def test_folder_filter_wildcards(self):
        # Exclusion only
        ff_ex = mail_filter.FolderFilter(exclude_patterns=["*Trash*", "*Spam*", "Junk*"])
        self.assertTrue(ff_ex.matches(["Inbox"]))
        self.assertTrue(ff_ex.matches(["Work", "Projects"]))
        self.assertFalse(ff_ex.matches(["Trash"]))
        self.assertFalse(ff_ex.matches(["[Gmail]", "Trash"]))
        self.assertFalse(ff_ex.matches(["Spam"]))
        self.assertFalse(ff_ex.matches(["Junk E-mail"]))

        # Inclusion + Exclusion
        ff_both = mail_filter.FolderFilter(
            include_patterns=["Work/*", "Archive/2023*"],
            exclude_patterns=["*Drafts*"],
        )
        self.assertTrue(ff_both.matches(["Work", "Marketing"]))
        self.assertTrue(ff_both.matches(["Archive", "2023-Q1"]))
        self.assertFalse(ff_both.matches(["Inbox"]))  # not in include
        self.assertFalse(ff_both.matches(["Work", "Drafts"]))  # excluded by drafts

    def test_migration_state_sqlite_incremental(self):
        db_path = os.path.join(self.tmpdir, "migration.db")
        state = mail_filter.MigrationState(db_path)
        self.assertTrue(state.is_active)
        self.assertFalse(state.is_completed(mid="<msg01@example.com>", sha256="hash01"))

        state.record_success(
            mid="<msg01@example.com>",
            sha256="hash01",
            source_file="inbox.eml",
            folder="Inbox",
        )
        state.close()

        # Reopen state and verify it remembers previously completed records
        state2 = mail_filter.MigrationState(db_path)
        self.assertTrue(state2.is_completed(mid="<msg01@example.com>"))
        self.assertTrue(state2.is_completed(sha256="hash01"))
        self.assertFalse(state2.is_completed(mid="<new_msg@example.com>", sha256="hash99"))
        state2.close()

    def test_mail2eml_date_filtering(self):
        # Create an mbox with 3 messages from 2021, 2023, and 2025
        mbox_path = os.path.join(self.tmpdir, "dates.mbox")
        with open(mbox_path, "wb") as f:
            f.write(
                b"From test@example.com Fri Jan 15 10:00:00 2021\r\n"
                b"From: test@example.com\r\n"
                b"Date: Fri, 15 Jan 2021 10:00:00 +0000\r\n"
                b"Subject: 2021 Email\r\n"
                b"Message-ID: <m2021@example.com>\r\n\r\n"
                b"Body 2021\r\n\r\n"
                b"From test@example.com Thu Jun 15 10:00:00 2023\r\n"
                b"From: test@example.com\r\n"
                b"Date: Thu, 15 Jun 2023 10:00:00 +0000\r\n"
                b"Subject: 2023 Email\r\n"
                b"Message-ID: <m2023@example.com>\r\n\r\n"
                b"Body 2023\r\n\r\n"
                b"From test@example.com Wed Jan 15 10:00:00 2025\r\n"
                b"From: test@example.com\r\n"
                b"Date: Wed, 15 Jan 2025 10:00:00 +0000\r\n"
                b"Subject: 2025 Email\r\n"
                b"Message-ID: <m2025@example.com>\r\n\r\n"
                b"Body 2025\r\n\r\n"
            )

        out_dir = os.path.join(self.tmpdir, "eml_dates_out")
        # Export only 2023
        exported, skipped, _, _ = mail2eml.export_to_eml(
            src_path=mbox_path,
            out_root=out_dir,
            since="2023-01-01",
            until="2023-12-31",
        )
        self.assertEqual(exported, 1)
        self.assertEqual(skipped, 2)

    def test_mail2mbox_folder_exclusion(self):
        # Create a directory tree with Inbox and Trash EMLs
        src_root = os.path.join(self.tmpdir, "src_mail")
        inbox_dir = os.path.join(src_root, "Inbox")
        trash_dir = os.path.join(src_root, "Trash")
        os.makedirs(inbox_dir, exist_ok=True)
        os.makedirs(trash_dir, exist_ok=True)

        with open(os.path.join(inbox_dir, "msg1.eml"), "wb") as f:
            f.write(b"From: a@b.com\r\nSubject: Keep\r\nMessage-ID: <k1@b.com>\r\n\r\nKeep body\r\n")
        with open(os.path.join(trash_dir, "msg2.eml"), "wb") as f:
            f.write(b"From: a@b.com\r\nSubject: Discard\r\nMessage-ID: <d1@b.com>\r\n\r\nDiscard body\r\n")

        out_mbox_dir = os.path.join(self.tmpdir, "out_mboxes")
        written, skipped = mail2mbox.export_to_mbox(
            src_path=src_root,
            target_path=out_mbox_dir,
            exclude_folders=["*Trash*"],
        )
        self.assertEqual(written, 1)
        self.assertEqual(skipped, 1)
        self.assertTrue(os.path.isfile(os.path.join(out_mbox_dir, "Inbox.mbox")))
        self.assertFalse(os.path.exists(os.path.join(out_mbox_dir, "Trash.mbox")))

    def test_mail2eml_incremental_delta_sync(self):
        # Create an mbox with 2 emails
        mbox_path = os.path.join(self.tmpdir, "delta.mbox")
        with open(mbox_path, "wb") as f:
            f.write(
                b"From test@example.com Mon Jan 01 10:00:00 2024\r\n"
                b"From: test@example.com\r\n"
                b"Subject: Initial 1\r\n"
                b"Message-ID: <delta1@example.com>\r\n\r\n"
                b"Body 1\r\n\r\n"
                b"From test@example.com Mon Jan 01 10:00:00 2024\r\n"
                b"From: test@example.com\r\n"
                b"Subject: Initial 2\r\n"
                b"Message-ID: <delta2@example.com>\r\n\r\n"
                b"Body 2\r\n\r\n"
            )

        state_db = os.path.join(self.tmpdir, "delta_state.sqlite")
        out_dir = os.path.join(self.tmpdir, "delta_eml_out")

        # First run: 2 exported, 0 skipped
        exp1, skip1, _, _ = mail2eml.export_to_eml(
            src_path=mbox_path,
            out_root=out_dir,
            state_file=state_db,
        )
        self.assertEqual(exp1, 2)
        self.assertEqual(skip1, 0)

        # Append a 3rd new email to the source mailbox
        with open(mbox_path, "ab") as f:
            f.write(
                b"From test@example.com Tue Jan 02 10:00:00 2024\r\n"
                b"From: test@example.com\r\n"
                b"Subject: Newly Arrived 3\r\n"
                b"Message-ID: <delta3@example.com>\r\n\r\n"
                b"Body 3\r\n\r\n"
            )

        # Second run (incremental delta sync): should only export the 1 new email!
        exp2, skip2, _, _ = mail2eml.export_to_eml(
            src_path=mbox_path,
            out_root=out_dir,
            state_file=state_db,
        )
        self.assertEqual(exp2, 1, "Delta sync must export only the new message")
        self.assertEqual(skip2, 2, "Delta sync must skip the 2 previously processed messages")

    def test_mail2mbox_incremental_delta_sync_jsonl(self):
        # Create an EML source folder with 2 emails
        src_dir = os.path.join(self.tmpdir, "delta_mbox_src")
        os.makedirs(src_dir, exist_ok=True)
        with open(os.path.join(src_dir, "msg1.eml"), "wb") as f:
            f.write(b"From: test@example.com\r\nSubject: MBox Delta 1\r\nMessage-ID: <mb1@example.com>\r\n\r\nBody 1\r\n")
        with open(os.path.join(src_dir, "msg2.eml"), "wb") as f:
            f.write(b"From: test@example.com\r\nSubject: MBox Delta 2\r\nMessage-ID: <mb2@example.com>\r\n\r\nBody 2\r\n")

        state_jsonl = os.path.join(self.tmpdir, "delta_state.jsonl")
        target_mbox = os.path.join(self.tmpdir, "out_delta.mbox")

        # Run 1
        w1, s1 = mail2mbox.export_to_mbox(
            src_path=src_dir,
            target_path=target_mbox,
            state_file=state_jsonl,
        )
        self.assertEqual(w1, 2)
        self.assertEqual(s1, 0)
        self.assertTrue(os.path.isfile(state_jsonl))

        # Add 3rd email
        with open(os.path.join(src_dir, "msg3.eml"), "wb") as f:
            f.write(b"From: test@example.com\r\nSubject: MBox Delta 3\r\nMessage-ID: <mb3@example.com>\r\n\r\nBody 3\r\n")

        # Run 2: incremental sync into the same mbox
        w2, s2 = mail2mbox.export_to_mbox(
            src_path=src_dir,
            target_path=target_mbox,
            state_file=state_jsonl,
        )
        self.assertEqual(w2, 1, "Delta sync must append only the 1 newly added message")
        self.assertEqual(s2, 2, "Delta sync must skip the 2 previously recorded messages")


if __name__ == "__main__":
    unittest.main()
