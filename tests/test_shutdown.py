#!/usr/bin/env python3
"""
tests/test_shutdown.py

Unit tests for graceful interruption, pause/stop controller, and signal handling:
  - ShutdownController signal handling and state transitions
  - Callback execution during shutdown
  - Graceful early exit in export_to_eml and export_to_mbox
  - State database and audit log persistence across interruptions
  - Resumption banner and reconciliation summary reporting
"""

import os
import shutil
import signal
import sys
import tempfile
import unittest

import tests  # Ensure repo root is at sys.path[0]

from mail_audit import AuditLogger
import mail2eml
import mail2mbox
from mail_shutdown import ShutdownController


class TestShutdownController(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="test_shutdown_")

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_controller_initial_state_and_request_stop(self):
        controller = ShutdownController(tool_name="test")
        self.assertFalse(controller.is_interrupted)
        self.assertFalse(controller.is_paused)

        callback_ran = []
        controller.register_callback(lambda: callback_ran.append(True))

        controller.request_stop("Testing manual halt")
        self.assertTrue(controller.is_interrupted)
        self.assertEqual(controller.interruption_reason, "Testing manual halt")
        self.assertEqual(callback_ran, [True])

    def test_controller_pause_and_resume(self):
        controller = ShutdownController(tool_name="test")
        self.assertFalse(controller.is_paused)
        controller.request_pause()
        self.assertTrue(controller.is_paused)
        controller.request_resume()
        self.assertFalse(controller.is_paused)

    def test_context_manager_registration(self):
        self.assertIsNone(ShutdownController.get_active())
        with ShutdownController(tool_name="ctx_test") as sc:
            self.assertIs(ShutdownController.get_active(), sc)
            self.assertTrue(sc.enabled)
        self.assertIsNone(ShutdownController.get_active())

    def test_signal_handling_first_strike(self):
        controller = ShutdownController(tool_name="sig_test")
        # Simulate SIGINT (Ctrl+C)
        controller._handle_signal(getattr(signal, "SIGINT", 2), None)
        self.assertTrue(controller.is_interrupted)
        self.assertIn("SIGINT", controller.interruption_reason)
        self.assertEqual(controller.interrupt_count, 1)

    def test_mail2eml_graceful_interruption(self):
        # Create an mbox with 5 messages
        mbox_path = os.path.join(self.tmpdir, "interrupted.mbox")
        with open(mbox_path, "wb") as f:
            for i in range(1, 6):
                f.write(
                    f"From test@example.com Mon Jan 01 10:00:00 2024\r\n"
                    f"From: test@example.com\r\n"
                    f"Subject: Message {i}\r\n"
                    f"Message-ID: <msg{i}@example.com>\r\n\r\n"
                    f"Body content {i}\r\n\r\n".encode("utf-8")
                )

        out_dir = os.path.join(self.tmpdir, "eml_out")
        state_file = os.path.join(self.tmpdir, "state.sqlite")
        audit_file = os.path.join(self.tmpdir, "audit.csv")

        logger = AuditLogger(audit_log_path=audit_file)
        controller = ShutdownController(tool_name="eml", logger=logger)

        # Trigger interruption after 2 messages are processed
        original_normalize = mail2eml.normalize_eml_bytes_with_flags
        processed_count = [0]
        def counting_normalize(*args, **kwargs):
            res = original_normalize(*args, **kwargs)
            processed_count[0] += 1
            if processed_count[0] >= 2:
                controller.request_stop("Simulated Ctrl+C after 2 messages")
            return res

        try:
            mail2eml.normalize_eml_bytes_with_flags = counting_normalize
            exp, skip, dir_counts, _ = mail2eml.export_to_eml(
                src_path=mbox_path,
                out_root=out_dir,
                state_file=state_file,
                logger=logger,
                shutdown=controller,
            )
        finally:
            mail2eml.normalize_eml_bytes_with_flags = original_normalize
            logger.close()

        # Should have stopped early (exported 2, not all 5)
        self.assertTrue(controller.is_interrupted)
        self.assertEqual(exp, 2, "Should have exported exactly 2 items before early interruption")
        self.assertTrue(os.path.isfile(state_file), "State file must be written and flushed")
        self.assertTrue(os.path.isfile(audit_file), "Audit log must be written and flushed")

        # Now resume the interrupted export with a new controller: should export the remaining 3!
        controller2 = ShutdownController(tool_name="eml")
        exp2, skip2, _, _ = mail2eml.export_to_eml(
            src_path=mbox_path,
            out_root=out_dir,
            state_file=state_file,
            shutdown=controller2,
        )
        self.assertEqual(exp2, 3, "Resumed run must export the remaining 3 messages")
        self.assertEqual(skip2, 2, "Resumed run must skip the 2 previously committed messages")

    def test_mail2mbox_graceful_interruption(self):
        # Create a directory with 4 EML files
        src_dir = os.path.join(self.tmpdir, "eml_source")
        os.makedirs(src_dir, exist_ok=True)
        for i in range(1, 5):
            with open(os.path.join(src_dir, f"mail_{i}.eml"), "wb") as f:
                f.write(
                    f"From: user@example.com\r\n"
                    f"Subject: Task {i}\r\n"
                    f"Message-ID: <task{i}@example.com>\r\n\r\n"
                    f"Body {i}\r\n".encode("utf-8")
                )

        target_mbox = os.path.join(self.tmpdir, "out.mbox")
        state_file = os.path.join(self.tmpdir, "mbox_state.sqlite")
        controller = ShutdownController(tool_name="mbox")

        # Request stop immediately after 1 message is written
        orig_write = mail2mbox.MboxWriterPool.write_entry
        written_count = [0]
        def counting_write(pool_self, *args, **kwargs):
            res = orig_write(pool_self, *args, **kwargs)
            written_count[0] += 1
            if written_count[0] >= 1:
                controller.request_stop("Simulated SIGTERM after 1 message")
            return res

        try:
            mail2mbox.MboxWriterPool.write_entry = counting_write
            w1, s1 = mail2mbox.export_to_mbox(
                src_path=src_dir,
                target_path=target_mbox,
                single_file_mode=True,
                state_file=state_file,
                shutdown=controller,
            )
        finally:
            mail2mbox.MboxWriterPool.write_entry = orig_write

        self.assertTrue(controller.is_interrupted)
        self.assertEqual(w1, 1, "Should have halted after writing 1 message")
        self.assertTrue(os.path.isfile(target_mbox))

        # Resume the MBOX export: should export remaining 3 messages into the mbox
        controller2 = ShutdownController(tool_name="mbox")
        w2, s2 = mail2mbox.export_to_mbox(
            src_path=src_dir,
            target_path=target_mbox,
            single_file_mode=True,
            state_file=state_file,
            overwrite=False,
            shutdown=controller2,
        )
        self.assertEqual(w2, 3, "Resumed MBOX export must write the 3 remaining messages")
        self.assertEqual(s2, 1, "Resumed MBOX export must skip the 1 previously written message")

    def test_resumption_banner_output(self):
        controller = ShutdownController(tool_name="pst")
        controller.request_stop("Testing banner")

        # Capture stdout
        from io import StringIO
        old_stdout = sys.stdout
        sys.stdout = StringIO()
        try:
            controller.print_resumption_banner(
                resume_file="D:\\Audit\\manifest.csv",
                state_file="D:\\Audit\\state.sqlite",
                tool="pst",
            )
            output = sys.stdout.getvalue()
        finally:
            sys.stdout = old_stdout

        self.assertIn("CONVERSION SAFELY PAUSED / STOPPED", output)
        self.assertIn("Zero data loss occurred", output)
        self.assertIn("--state-file", output)
        self.assertIn("--resume", output)


if __name__ == "__main__":
    unittest.main()
