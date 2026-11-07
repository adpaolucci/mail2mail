#!/usr/bin/env python3
"""
tests/test_cross_compatibility.py

Automated Cross-Testing & Differential Conformance Suite between
mail2mail (Python) and mail2rust (Rust):
  - Protocol 1: Differential Output Parity (A vs B on identical input)
  - Protocol 2: Round-Trip Parity (Python MBOX -> EML -> Rust MBOX -> Python EML)
  - Protocol 3: Checkpoint Delta Hand-Off Python -> Rust
  - Protocol 4: Checkpoint Delta Hand-Off Rust -> Python
  - Protocol 5: Forensic Audit Manifest Hand-Off Rust -> Python
  - Protocol 6: Date Filtering Parity (--since, --until)
  - Protocol 7: Folder Inclusions and Exclusions Parity

When `mail2rust` is installed or discovered (via MAIL2RUST_BIN or sibling repo),
these tests execute full end-to-end differential verification against the compiled
Rust engine. When mail2rust is not present, tests are cleanly skipped so mail2mail's
CI runs autonomously without any build dependency.
"""

from datetime import datetime, timezone
import email
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

import tests  # Ensure repo root is on sys.path

import mail2eml
import mail2mbox
from mail_audit import AuditLogger


def find_mail2rust_binary() -> str | None:
    """Locate the mail2rust executable via env var, sibling directories, or PATH."""
    env_bin = os.environ.get("MAIL2RUST_BIN")
    if env_bin and os.path.isfile(env_bin):
        return os.path.abspath(env_bin)

    candidates = [
        os.path.join("..", "mail2rust", "target", "release", "mail2rust"),
        os.path.join("..", "mail2rust", "target", "debug", "mail2rust"),
        os.path.join("..", "..", "mail2rust", "target", "release", "mail2rust"),
        os.path.join("..", "..", "mail2rust", "target", "debug", "mail2rust"),
    ]
    if os.name == "nt":
        candidates = [c + ".exe" for c in candidates]

    for c in candidates:
        if os.path.isfile(c):
            return os.path.abspath(c)

    return shutil.which("mail2rust")


MAIL2RUST_BIN = find_mail2rust_binary()
HAS_MAIL2RUST = MAIL2RUST_BIN is not None


def count_emls(dir_path: str) -> int:
    """Count all .eml files recursively in a directory tree."""
    total = 0
    for _, _, files in os.walk(dir_path):
        for f in files:
            if f.lower().endswith(".eml"):
                total += 1
    return total


def extract_message_ids_from_dir(dir_path: str) -> set[str]:
    """Extract all Message-ID header values from all .eml files in directory."""
    mids = set()
    for root, _, files in os.walk(dir_path):
        for fn in files:
            if fn.lower().endswith(".eml"):
                with open(os.path.join(root, fn), "rb") as fh:
                    msg = email.message_from_binary_file(fh)
                    mid = msg.get("Message-ID", "").strip()
                    if mid:
                        mids.add(mid)
    return mids


def create_sample_mbox(path: str, count: int = 5) -> None:
    """Generate a reproducible test MBOX file with diverse MIME characteristics."""
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    with open(path, "wb") as f:
        for i in range(1, count + 1):
            dt = datetime(2024, 1, i, 10, 0, 0, tzinfo=timezone.utc)
            date_str = dt.strftime("%a, %d %b %Y %H:%M:%S +0000")
            asctime_str = dt.strftime("%a %b %d %H:%M:%S %Y")
            f.write(
                f"From test{i}@example.com {asctime_str}\r\n"
                f"From: Sender {i} <test{i}@example.com>\r\n"
                f"To: Recipient <recip@example.com>\r\n"
                f"Date: {date_str}\r\n"
                f"Subject: Cross-Test Subject {i} [Item-{i}]\r\n"
                f"Message-ID: <cross_compat_{i}@domain.test>\r\n"
                f"X-Status: {'R' if i % 2 == 0 else ''}\r\n"
                f"MIME-Version: 1.0\r\n"
                f"Content-Type: text/plain; charset=utf-8\r\n\r\n"
                f"This is body payload for message number {i}.\r\n"
                f"Testing cross-compatibility between Python and Rust.\r\n\r\n".encode("utf-8")
            )


class TestCrossCompatibility(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="test_cross_compat_")

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_cross_test_contract_helpers(self):
        """Verify the sample mbox generator and parser verification helpers work standalone."""
        mbox_path = os.path.join(self.tmpdir, "sample.mbox")
        create_sample_mbox(mbox_path, count=3)
        self.assertTrue(os.path.isfile(mbox_path))

        out_dir = os.path.join(self.tmpdir, "eml_out")
        exp, skip, _, _ = mail2eml.export_to_eml(mbox_path, out_dir, layout="flat")
        self.assertEqual(exp, 3)
        self.assertEqual(skip, 0)
        self.assertEqual(count_emls(out_dir), 3)

    @unittest.skipUnless(HAS_MAIL2RUST, "mail2rust binary not found (set MAIL2RUST_BIN=...)")
    def test_differential_eml_export(self):
        """Assert mail2mail and mail2rust produce identical EML counts and Message-IDs."""
        mbox_path = os.path.join(self.tmpdir, "diff.mbox")
        create_sample_mbox(mbox_path, count=5)

        py_out = os.path.join(self.tmpdir, "py_emls")
        rust_out = os.path.join(self.tmpdir, "rust_emls")

        # 1. Run Python
        mail2eml.export_to_eml(mbox_path, py_out, layout="flat")

        # 2. Run Rust
        subprocess.run(
            [MAIL2RUST_BIN, "eml", mbox_path, rust_out, "--layout", "flat"],
            check=True,
            capture_output=True,
        )

        # 3. Compare outputs
        self.assertEqual(count_emls(py_out), 5)
        self.assertEqual(count_emls(rust_out), 5)
        self.assertEqual(extract_message_ids_from_dir(py_out), extract_message_ids_from_dir(rust_out))

    @unittest.skipUnless(HAS_MAIL2RUST, "mail2rust binary not found (set MAIL2RUST_BIN=...)")
    def test_round_trip_parity(self):
        """Assert round-trip fidelity: Python (MBOX -> EML) -> Rust (EML -> MBOX) -> Python (MBOX -> EML)."""
        src_mbox = os.path.join(self.tmpdir, "source.mbox")
        create_sample_mbox(src_mbox, count=4)

        eml_pass1 = os.path.join(self.tmpdir, "eml_pass1")
        rust_mbox = os.path.join(self.tmpdir, "rust_reconstructed.mbox")
        eml_pass2 = os.path.join(self.tmpdir, "eml_pass2")

        # Pass 1: Python exports MBOX -> EML
        exp1, _, _, _ = mail2eml.export_to_eml(src_mbox, eml_pass1, layout="flat")
        self.assertEqual(exp1, 4)

        # Pass 2: Rust converts EML -> MBOX
        subprocess.run(
            [MAIL2RUST_BIN, "mbox", eml_pass1, rust_mbox],
            check=True,
            capture_output=True,
        )
        self.assertTrue(os.path.isfile(rust_mbox))

        # Pass 3: Python converts Rust's MBOX -> EML
        exp2, _, _, _ = mail2eml.export_to_eml(rust_mbox, eml_pass2, layout="flat")
        self.assertEqual(exp2, 4, "Round-trip message count must be 100% preserved")
        self.assertEqual(extract_message_ids_from_dir(eml_pass1), extract_message_ids_from_dir(eml_pass2))

    @unittest.skipUnless(HAS_MAIL2RUST, "mail2rust binary not found (set MAIL2RUST_BIN=...)")
    def test_checkpoint_state_handoff_python_to_rust(self):
        """Assert Rust resumes from Python's JSON checkpoint state and processes only delta messages."""
        mbox_path = os.path.join(self.tmpdir, "delta.mbox")
        create_sample_mbox(mbox_path, count=3)

        checkpoint_json = os.path.join(self.tmpdir, "checkpoint.json")
        out_dir = os.path.join(self.tmpdir, "delta_out")

        # Pass 1: Python exports 3 messages and writes checkpoint.json
        exp1, skip1, _, _ = mail2eml.export_to_eml(
            mbox_path, out_dir, state_file=checkpoint_json, layout="flat"
        )
        self.assertEqual(exp1, 3)
        self.assertTrue(os.path.isfile(checkpoint_json))

        # Append 2 new messages
        with open(mbox_path, "ab") as f:
            for i in (4, 5):
                f.write(
                    f"From test{i}@example.com Tue Jan 02 10:00:00 2024\r\n"
                    f"From: test{i}@example.com\r\n"
                    f"Subject: Appended {i}\r\n"
                    f"Message-ID: <cross_compat_{i}@domain.test>\r\n\r\n"
                    f"New body {i}\r\n\r\n".encode("utf-8")
                )

        # Pass 2: Rust runs with --resume checkpoint.json
        subprocess.run(
            [MAIL2RUST_BIN, "eml", mbox_path, out_dir, "--resume", checkpoint_json, "--layout", "flat"],
            capture_output=True,
            text=True,
            check=True,
        )

        # Total files in destination must be exactly 3 (from Python) + 2 (from Rust) = 5
        self.assertEqual(count_emls(out_dir), 5, "Total files after delta sync should be 3 + 2 = 5")
        all_mids = extract_message_ids_from_dir(out_dir)
        self.assertEqual(len(all_mids), 5)
        for i in range(1, 6):
            self.assertIn(f"<cross_compat_{i}@domain.test>", all_mids)

    @unittest.skipUnless(HAS_MAIL2RUST, "mail2rust binary not found (set MAIL2RUST_BIN=...)")
    def test_checkpoint_state_handoff_rust_to_python(self):
        """Assert Python resumes from Rust's JSON checkpoint state and processes only delta messages."""
        mbox_path = os.path.join(self.tmpdir, "rust_delta.mbox")
        create_sample_mbox(mbox_path, count=3)

        checkpoint_json = os.path.join(self.tmpdir, "rust_checkpoint.json")
        out_dir = os.path.join(self.tmpdir, "rust_delta_out")

        # Pass 1: Rust exports 3 messages and writes checkpoint JSON via --state-file
        subprocess.run(
            [MAIL2RUST_BIN, "eml", mbox_path, out_dir, "--state-file", checkpoint_json, "--layout", "flat"],
            capture_output=True,
            text=True,
            check=True,
        )
        self.assertEqual(count_emls(out_dir), 3)
        self.assertTrue(os.path.isfile(checkpoint_json))

        # Append 2 new messages
        with open(mbox_path, "ab") as f:
            for i in (4, 5):
                f.write(
                    f"From test{i}@example.com Tue Jan 02 10:00:00 2024\r\n"
                    f"From: test{i}@example.com\r\n"
                    f"Subject: Appended {i}\r\n"
                    f"Message-ID: <cross_compat_{i}@domain.test>\r\n\r\n"
                    f"New body {i}\r\n\r\n".encode("utf-8")
                )

        # Pass 2: Python runs with resume=checkpoint_json
        exp, skip, _, _ = mail2eml.export_to_eml(
            mbox_path, out_dir, resume=checkpoint_json, layout="flat"
        )
        self.assertEqual(exp, 2, "Python should export only the 2 newly appended messages")
        self.assertEqual(count_emls(out_dir), 5, "Total files after delta sync should be 3 + 2 = 5")

    @unittest.skipUnless(HAS_MAIL2RUST, "mail2rust binary not found (set MAIL2RUST_BIN=...)")
    def test_forensic_csv_manifest_handoff_rust_to_python(self):
        """Assert Python resumes from Rust's forensic audit manifest CSV and skips exported messages."""
        mbox_path = os.path.join(self.tmpdir, "manifest_handoff.mbox")
        create_sample_mbox(mbox_path, count=3)

        manifest_csv = os.path.join(self.tmpdir, "audit_manifest.csv")
        out_dir = os.path.join(self.tmpdir, "manifest_out")

        # Pass 1: Rust exports 3 messages and writes forensic CSV manifest
        subprocess.run(
            [MAIL2RUST_BIN, "eml", mbox_path, out_dir, "--audit-log", manifest_csv, "--layout", "flat"],
            capture_output=True,
            text=True,
            check=True,
        )
        self.assertEqual(count_emls(out_dir), 3)
        self.assertTrue(os.path.isfile(manifest_csv))

        # Verify completed keys can be loaded by Python from Rust's CSV
        loaded_keys = AuditLogger.load_completed_keys(manifest_csv)
        self.assertEqual(len(loaded_keys), 6)  # 3 mid:* + 3 sha256:*

        # Append 2 new messages
        with open(mbox_path, "ab") as f:
            for i in (4, 5):
                f.write(
                    f"From test{i}@example.com Tue Jan 02 10:00:00 2024\r\n"
                    f"From: test{i}@example.com\r\n"
                    f"Subject: Appended {i}\r\n"
                    f"Message-ID: <cross_compat_{i}@domain.test>\r\n\r\n"
                    f"New body {i}\r\n\r\n".encode("utf-8")
                )

        # Pass 2: Python resumes from Rust's audit CSV manifest
        exp, skip, _, _ = mail2eml.export_to_eml(
            mbox_path, out_dir, resume=manifest_csv, layout="flat"
        )
        self.assertEqual(exp, 2, "Python should export only the 2 un-manifested messages")
        self.assertEqual(count_emls(out_dir), 5)

    @unittest.skipUnless(HAS_MAIL2RUST, "mail2rust binary not found (set MAIL2RUST_BIN=...)")
    def test_date_range_parity(self):
        """Assert Python and Rust apply identical date boundaries (--since, --until)."""
        mbox_path = os.path.join(self.tmpdir, "dates.mbox")
        create_sample_mbox(mbox_path, count=5)  # dates are Jan 1, Jan 2, Jan 3, Jan 4, Jan 5 2024

        py_out = os.path.join(self.tmpdir, "py_date_filtered")
        rust_out = os.path.join(self.tmpdir, "rust_date_filtered")

        # Filter: Jan 2 to Jan 4 inclusive (Items 2, 3, 4)
        since_date = "2024-01-02"
        until_date = "2024-01-04"

        mail2eml.export_to_eml(
            mbox_path, py_out, since=since_date, until=until_date, layout="flat"
        )

        subprocess.run(
            [
                MAIL2RUST_BIN,
                "eml",
                mbox_path,
                rust_out,
                "--since",
                since_date,
                "--until",
                until_date,
                "--layout",
                "flat",
            ],
            capture_output=True,
            text=True,
            check=True,
        )

        py_count = count_emls(py_out)
        rust_count = count_emls(rust_out)
        self.assertEqual(py_count, 3, "Both should filter down to exactly 3 items (Jan 2, 3, 4)")
        self.assertEqual(rust_count, 3)
        self.assertEqual(
            extract_message_ids_from_dir(py_out),
            extract_message_ids_from_dir(rust_out),
            "Filtered Message-IDs must match exactly between Python and Rust",
        )

    @unittest.skipUnless(HAS_MAIL2RUST, "mail2rust binary not found (set MAIL2RUST_BIN=...)")
    def test_folder_filtering_parity(self):
        """Assert Python and Rust folder pattern inclusion and exclusion produce identical outputs."""
        src_root = os.path.join(self.tmpdir, "multi_folder_src")
        inbox_dir = os.path.join(src_root, "Inbox")
        archive_dir = os.path.join(src_root, "Archive")
        os.makedirs(inbox_dir, exist_ok=True)
        os.makedirs(archive_dir, exist_ok=True)

        # Write 2 messages in Inbox and 2 in Archive
        with open(os.path.join(inbox_dir, "inbox1.eml"), "wb") as f:
            f.write(b"From: a@b.c\r\nSubject: In1\r\nMessage-ID: <in1@b.c>\r\n\r\nBody\r\n")
        with open(os.path.join(inbox_dir, "inbox2.eml"), "wb") as f:
            f.write(b"From: a@b.c\r\nSubject: In2\r\nMessage-ID: <in2@b.c>\r\n\r\nBody\r\n")
        with open(os.path.join(archive_dir, "arch1.eml"), "wb") as f:
            f.write(b"From: a@b.c\r\nSubject: Arch1\r\nMessage-ID: <arch1@b.c>\r\n\r\nBody\r\n")
        with open(os.path.join(archive_dir, "arch2.eml"), "wb") as f:
            f.write(b"From: a@b.c\r\nSubject: Arch2\r\nMessage-ID: <arch2@b.c>\r\n\r\nBody\r\n")

        py_out = os.path.join(self.tmpdir, "py_filtered_folder")
        rust_out = os.path.join(self.tmpdir, "rust_filtered_folder")

        # Exclude Archive*
        mail2eml.export_to_eml(src_root, py_out, exclude_folders=["Archive*"], layout="flat")

        subprocess.run(
            [
                MAIL2RUST_BIN,
                "eml",
                src_root,
                rust_out,
                "--exclude-folder",
                "Archive*",
                "--layout",
                "flat",
            ],
            capture_output=True,
            text=True,
            check=True,
        )

        py_mids = extract_message_ids_from_dir(py_out)
        rust_mids = extract_message_ids_from_dir(rust_out)
        self.assertEqual(len(py_mids), 2)
        self.assertEqual(len(rust_mids), 2)
        self.assertEqual(py_mids, {"<in1@b.c>", "<in2@b.c>"})
        self.assertEqual(rust_mids, {"<in1@b.c>", "<in2@b.c>"})


if __name__ == "__main__":
    unittest.main()
