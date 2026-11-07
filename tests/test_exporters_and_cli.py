#!/usr/bin/env python3
"""
tests/test_exporters_and_cli.py — Automated end-to-end tests for:
  - mail2eml.py (folder/year/month/flat layouts, subject/index naming, rollover, year filters)
  - mail2mbox.py (single .mbox, compressed .mbox.gz, multi-mbox tree, part rotation, round-trip)
  - mail2pst.py (populate_mailitem, MAPI properties after HTMLBody, MAX_PATH safety, com_retry, iter_pst_records)
  - mail2mail.py (master CLI mode inference across eml, mbox, and pst targets)
"""

from __future__ import annotations

from datetime import datetime, timezone
import email
from email import policy
import os
import shutil
import tempfile
import unittest

import tests  # noqa: F401
import mail_formats
import mail2eml
import mail2mbox
import mail2pst
import mail2mail


class TestMail2EmlAndMail2Mbox(unittest.TestCase):
    """Test mail2eml and mail2mbox conversion engines and round-trip fidelity."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="test_exporters_")

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_thunderbird_imapmail_to_eml_and_roundtrip_to_mbox_gz(self):
        imap_root = os.path.join(self.tmpdir, "ImapMail")
        acct_dir = os.path.join(imap_root, "imap.example.com")
        os.makedirs(os.path.join(acct_dir, "Notes &- Invoices.sbd"), exist_ok=True)
        tb_mbox = os.path.join(acct_dir, "Notes &- Invoices.sbd", "Caf&AOk-.")
        tb_mbox_bytes = (
            b"From - Mon Jan 01 00:00:00 2024\r\n"
            b"Subject: Empty body message\r\n"
            b"Date: Mon, 01 Jan 2024 00:00:00 +0000\r\n"
            b"X-Mozilla-Status: 0004\r\n"
            b"\r\n"
            b"From - Mon Jan 01 00:01:00 2024\r\n"
            b"Subject: Second message with escaped From\r\n"
            b"Date: Mon, 01 Jan 2024 00:01:00 +0000\r\n"
            b"X-Mozilla-Status: 1003\r\n"
            b"\r\n"
            b">From someone in body\r\n"
            b">>From nested quote\r\n"
            b"\r\n"
            b"From - Mon Jan 01 00:02:00 2024\r\n"
            b"Subject: Expunged message\r\n"
            b"X-Mozilla-Status: 0008\r\n"
            b"\r\n"
            b"Should be skipped\r\n"
        )
        with open(tb_mbox, "wb") as f:
            f.write(tb_mbox_bytes)

        eml_out = os.path.join(self.tmpdir, "out_eml")
        mail2eml.main([imap_root, eml_out, "--naming", "index"])

        expected_folder = os.path.join(eml_out, "imap.example.com", "Notes & Invoices", "Café")
        self.assertTrue(os.path.isdir(expected_folder))
        self.assertEqual(sorted(os.listdir(expected_folder)), ["000001.eml", "000002.eml"])

        # Round-trip EML folder tree -> compressed .mbox.gz via mail2mbox
        mbox_gz_out = os.path.join(self.tmpdir, "backup.mbox.gz")
        mail2mbox.main([eml_out, mbox_gz_out])
        self.assertTrue(os.path.isfile(mbox_gz_out))

        rt_recs = list(mail_formats.iter_mail_records(mbox_gz_out))
        self.assertEqual(len(rt_recs), 2)
        m2 = email.message_from_bytes(rt_recs[1].raw_bytes, policy=policy.default)
        self.assertIn("From someone in body", m2.get_content())
        self.assertIn(">From nested quote", m2.get_content())
        self.assertTrue(rt_recs[1].flags["is_read"])
        self.assertTrue(rt_recs[1].flags["is_replied"])
        self.assertTrue(rt_recs[1].flags["is_forwarded"])

    def test_mail2mbox_folder_layout_and_part_rotation(self):
        src_mbox = os.path.join(self.tmpdir, "Inbox.mbox")
        with open(src_mbox, "wb") as f:
            for i in range(1, 5):
                f.write(
                    f"From user{i}@example.com Fri Nov  6 12:0{i}:00 2020\r\n"
                    f"From: user{i}@example.com\r\n"
                    f"Subject: Msg {i}\r\n"
                    f"Date: Fri, 06 Nov 2020 12:0{i}:00 +0000\r\n\r\n"
                    f"{'X' * 500}\r\n\r\n".encode("utf-8")
                )

        mbox_tree = os.path.join(self.tmpdir, "mbox_tree")
        # Force rotation after ~1 KB (1e-6 GB)
        written, skipped = mail2mbox.export_to_mbox(
            src_path=src_mbox,
            target_path=mbox_tree,
            single_file_mode=False,
            layout="year",
            max_mbox_gb=1e-6,
        )
        self.assertEqual(written, 4)
        self.assertEqual(skipped, 0)
        year_dir = os.path.join(mbox_tree, "Inbox")
        files = sorted(os.listdir(year_dir))
        self.assertIn("2020.mbox", files)
        self.assertIn("2020__part2.mbox", files)


class TestMail2PstAndPstReader(unittest.TestCase):
    """Test mail2pst MAPI population, COM retry, MAX_PATH protection, and PST input reader."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="test_pst_")

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_populate_mailitem_and_iter_pst_records(self):
        long_att_name = ("very_long_attachment_filename_" * 6) + ".pdf"
        mime_complex = (
            b"From: \"Last; First\" <sender@example.com>\r\n"
            b"To: \"Recip; One\" <r1@example.com>\r\n"
            b"To: Second Recip <r2@example.com>\r\n"
            b"Subject: Complex MIME\r\n"
            b"Date: Fri, 06 Nov 2020 15:30:00 -0500\r\n"
            b"Message-ID: <complex@example.com>\r\n"
            b"Status: RO\r\n"
            b"X-Status: FA\r\n"
            b"Content-Type: multipart/mixed; boundary=\"mix\"\r\n\r\n"
            b"--mix\r\n"
            b"Content-Type: text/html; charset=\"utf-8\"\r\n\r\n"
            b"<html><body>Hello <img src=\"cid:img1\"></body></html>\r\n"
            b"--mix\r\n"
            b"Content-Type: image/png\r\n"
            b"Content-ID: <img1>\r\n"
            b"Content-Transfer-Encoding: base64\r\n\r\n"
            b"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==\r\n"
            b"--mix\r\n"
            + f"Content-Type: application/pdf; name=\"{long_att_name}\"\r\n".encode()
            + f"Content-Disposition: attachment; filename=\"{long_att_name}\"\r\n\r\n".encode()
            + b"%PDF-1.4 dummy\r\n"
            b"--mix--\r\n"
        )

        class DummyPA:
            def __init__(self):
                self.props = {}
            def SetProperty(self, tag, val):
                self.props[tag] = val
            def GetProperty(self, tag):
                return self.props.get(tag)

        class DummyAtt:
            def __init__(self, path, disp_name):
                self.path = path
                self.FileName = os.path.basename(path)
                self.PropertyAccessor = DummyPA()
            def SaveAsFile(self, dst):
                with open(dst, "wb") as f:
                    f.write(b"dummy_att_bytes")

        class DummyAttachments:
            def __init__(self, test_case):
                self.items = []
                self.tc = test_case
            @property
            def Count(self):
                return len(self.items)
            def Item(self, idx):
                return self.items[idx - 1]
            def Add(self, path, typ=1, pos=1, disp_name=None):
                self.tc.assertTrue(os.path.isfile(path))
                self.tc.assertLessEqual(len(path), 240)
                att = DummyAtt(path, disp_name)
                self.items.append(att)
                return att

        class DummyMail:
            def __init__(self, test_case):
                self.Class = 43  # olMail
                self.PropertyAccessor = DummyPA()
                self.Attachments = DummyAttachments(test_case)
                self.Subject = ""
                self.To = ""
                self.CC = ""
                self.BCC = ""
                self.SenderName = "Last, First"
                self.SenderEmailAddress = "sender@example.com"
                self.ReceivedTime = datetime(2020, 11, 6, 20, 30, 0, tzinfo=timezone.utc)
                self.Body = ""
                self._html_body = ""
                self.UnRead = True
                self.FlagStatus = 0
                self.FlagRequest = ""
                self.saved = False
            @property
            def HTMLBody(self):
                return self._html_body
            @HTMLBody.setter
            def HTMLBody(self, val):
                # Simulate OOM resetting transport headers if HTMLBody is set after MAPI props
                self.PropertyAccessor.props.pop(mail2pst.PR_TRANSPORT_MESSAGE_HEADERS_W, None)
                self.PropertyAccessor.props.pop(mail2pst.PR_SENDER_EMAIL_ADDRESS_W, None)
                self._html_body = val
            def Save(self):
                self.saved = True

        parsed = email.message_from_bytes(mime_complex, policy=policy.default)
        mail = DummyMail(self)
        mail2pst.populate_mailitem(mail, parsed, mime_complex, self.tmpdir)

        self.assertTrue(mail.saved)
        self.assertFalse(mail.UnRead)
        self.assertEqual(mail.FlagStatus, 2)
        self.assertEqual(mail.To, "Recip, One <r1@example.com>; Second Recip <r2@example.com>")
        self.assertIn(mail2pst.PR_TRANSPORT_MESSAGE_HEADERS_W, mail.PropertyAccessor.props)
        self.assertIn(mail2pst.PR_SENDER_EMAIL_ADDRESS_W, mail.PropertyAccessor.props)
        self.assertEqual(
            mail.PropertyAccessor.props[mail2pst.PR_MESSAGE_DELIVERY_TIME],
            datetime(2020, 11, 6, 20, 30, 0),
        )

        # Now verify iter_pst_records reading this mail item back from a mocked PST Store
        class DummySubFolders:
            Count = 0
            def Item(self, idx):
                return None

        class DummyPstRoot:
            def __init__(self, item):
                self._items = [item]
                self.Folders = DummySubFolders()
                self.Items = self
            @property
            def Count(self):
                return len(self._items)
            def Item(self, idx):
                return self._items[idx - 1]

        class DummyStore:
            def __init__(self, path, root):
                self.StoreID = "store1"
                self.FilePath = path
                self._root = root
            def GetRootFolder(self):
                return self._root

        class DummyNS:
            def __init__(self, store):
                self.Stores = [store]
            def RemoveStore(self, root):
                pass

        fake_pst = os.path.join(self.tmpdir, "input.pst")
        pst_recs = list(mail_formats.iter_pst_records(fake_pst, namespace=DummyNS(DummyStore(fake_pst, DummyPstRoot(mail)))))
        self.assertEqual(len(pst_recs), 1)
        self.assertTrue(pst_recs[0].flags["is_read"])
        self.assertTrue(pst_recs[0].flags["is_flagged"])
        self.assertTrue(pst_recs[0].flags["is_replied"])

    def test_com_retry_recovers_from_rpc_e_call_rejected(self):
        attempts = {"count": 0}

        @mail2pst.com_retry(max_attempts=4, base_delay=0.01)
        def flaky_com_call():
            attempts["count"] += 1
            if attempts["count"] < 3:
                err = Exception("Call was rejected by callee.")
                err.hresult = -2147418111  # RPC_E_CALL_REJECTED
                raise err
            return "ok"

        self.assertEqual(flaky_com_call(), "ok")
        self.assertEqual(attempts["count"], 3)


class TestMail2MailRouter(unittest.TestCase):
    """Test mail2mail._infer_mode CLI routing across subcommands, flags, and file extensions."""

    def test_infer_mode_subcommands_and_extensions(self):
        self.assertEqual(mail2mail._infer_mode(["eml", "src", "dst"])[0], "eml")
        self.assertEqual(mail2mail._infer_mode(["mbox", "src", "dst"])[0], "mbox")
        self.assertEqual(mail2mail._infer_mode(["pst", "src", "dst"])[0], "pst")
        self.assertEqual(mail2mail._infer_mode(["src", "dst", "--to", "mbox"])[0], "mbox")
        self.assertEqual(mail2mail._infer_mode(["src", "dst", "--to=pst"])[0], "pst")
        # PST source -> EML directory destination must infer 'eml', not 'pst'!
        self.assertEqual(mail2mail._infer_mode([r"D:\Archive.pst", r"D:\Export_EML"])[0], "eml")
        # PST source -> .mbox.gz destination must infer 'mbox'
        self.assertEqual(mail2mail._infer_mode([r"D:\Archive.pst", r"D:\Archive.mbox.gz"])[0], "mbox")
        # MBOX source -> .pst destination must infer 'pst'
        self.assertEqual(mail2mail._infer_mode([r"D:\Inbox.mbox", r"D:\Out.pst"])[0], "pst")


if __name__ == "__main__":
    unittest.main()
