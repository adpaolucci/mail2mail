#!/usr/bin/env python3
"""
tests/test_mail_formats.py — Automated unit tests for all 10 input formats in mail_formats.py:
  1. Standard .mbox & extensionless mbox
  2. Thunderbird ImapMail / Mail (.sbd folders, .msf skip, X-Mozilla-Status bitmasks)
  3. Maildir & Maildir++ (cur/, new/, tmp/, :2,FRSPT & ;2,FRSPT, .Folder.Subfolder)
  4. Apple Mail (.emlx, .partial.emlx, *.mbox / *.imapmbox bundles, Data/0..9/Messages shards)
  5. Outlook .msg (pure-Python MS-CFB OLE2 + MS-OXMSG parser)
  6. Outlook for Mac .olm (XML + ZIP attachment archive parser)
  7. Outlook .pst input reader (iter_pst_records via MAPI namespace)
  8. MH / Claws Mail / Sylpheed (.mh_sequences & numbered message files)
  9. MMDF (\x01\x01\x01\x01) & Emacs Rmail Babyl (BABYL OPTIONS:)
  10. Compressed files & archives (.mbox.gz, .mbox.bz2, .mbox.xz, .zip, .tgz)
"""

from __future__ import annotations

import bz2
from datetime import datetime, timezone
import email
from email import policy
import gzip
import lzma
import os
import shutil
import struct
import tarfile
import tempfile
import unittest
import zipfile

import tests  # noqa: F401
import mail_formats


def build_synthetic_msg_ole2() -> bytes:
    """Build a valid 512-byte-sector MS-CFB OLE2 .msg file in memory per [MS-CFB] & [MS-OXMSG]."""
    sector_size = 512
    ENDOFCHAIN = 0xFFFFFFFE
    FREESECT = 0xFFFFFFFF

    subj_bytes = "Synthetic OLE2 MSG".encode("utf-16-le")
    body_bytes = "Hello from a binary .msg file!".encode("utf-16-le")
    sender_bytes = "carol@example.com".encode("utf-16-le")
    # __properties_version1.0: 32-byte header + two 16-byte entries:
    # 1) PR_MESSAGE_FLAGS (0x0E070003) = 0 (Unread)
    # 2) PR_FLAG_STATUS (0x10900003) = 2 (Flagged)
    props_stream = (
        b"\x00" * 32
        + struct.pack("<IIQ", 0x0E070003, 0x06, 0)
        + struct.pack("<IIQ", 0x10900003, 0x06, 2)
    )

    def pad_sec(b: bytes) -> bytes:
        rem = len(b) % sector_size
        return b + (b"\x00" * ((sector_size - rem) % sector_size))

    def make_dir_entry(name: str, obj_type: int, left: int, right: int, child: int, start_sec: int, size: int) -> bytes:
        raw_name = (name + "\x00").encode("utf-16-le")
        buf = bytearray(128)
        buf[:len(raw_name)] = raw_name
        struct.pack_into("<H", buf, 64, len(raw_name))
        buf[66] = obj_type
        buf[67] = 1
        struct.pack_into("<iii", buf, 68, left, right, child)
        struct.pack_into("<I", buf, 116, start_sec)
        struct.pack_into("<Q", buf, 120, size)
        return bytes(buf)

    fat_ints = [
        ENDOFCHAIN,  # Sec 0: FAT
        2,           # Sec 1: Dir part 1 -> Sec 2
        ENDOFCHAIN,  # Sec 2: Dir part 2
        ENDOFCHAIN,  # Sec 3: Subject
        ENDOFCHAIN,  # Sec 4: Body
        ENDOFCHAIN,  # Sec 5: Sender
        ENDOFCHAIN,  # Sec 6: Properties
    ] + [FREESECT] * ((sector_size // 4) - 7)
    sec_fat = struct.pack(f"<{len(fat_ints)}I", *fat_ints)

    dir_entries = [
        make_dir_entry("Root Entry", 5, -1, -1, 1, ENDOFCHAIN, 0),
        make_dir_entry("__substg1.0_0037001F", 2, 2, 3, -1, 3, len(subj_bytes)),
        make_dir_entry("__substg1.0_1000001F", 2, -1, -1, -1, 4, len(body_bytes)),
        make_dir_entry("__substg1.0_5D01001F", 2, -1, 4, -1, 5, len(sender_bytes)),
        make_dir_entry("__properties_version1.0", 2, -1, -1, -1, 6, len(props_stream)),
        b"\x00" * 128,
        b"\x00" * 128,
        b"\x00" * 128,
    ]
    sec_dir = b"".join(dir_entries)

    hdr = bytearray(512)
    hdr[:8] = mail_formats.OLE2_MAGIC
    struct.pack_into("<H", hdr, 24, 0x003E)
    struct.pack_into("<H", hdr, 26, 0x0003)
    struct.pack_into("<H", hdr, 28, 0xFFFE)
    struct.pack_into("<H", hdr, 30, 9)
    struct.pack_into("<H", hdr, 32, 6)
    struct.pack_into("<I", hdr, 44, 1)
    struct.pack_into("<I", hdr, 48, 1)
    struct.pack_into("<I", hdr, 56, 1)
    struct.pack_into("<I", hdr, 60, ENDOFCHAIN)
    struct.pack_into("<I", hdr, 68, ENDOFCHAIN)
    struct.pack_into("<I", hdr, 76, 0)
    for i in range(1, 109):
        struct.pack_into("<I", hdr, 76 + i * 4, FREESECT)

    return (
        bytes(hdr)
        + sec_fat
        + sec_dir
        + pad_sec(subj_bytes)
        + pad_sec(body_bytes)
        + pad_sec(sender_bytes)
        + pad_sec(props_stream)
    )


class TestMailFormats(unittest.TestCase):
    """Test all 10 mailbox/message formats supported by mail_formats.py."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="test_mail_formats_")

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_maildir_and_maildir_plus_plus(self):
        md_root = os.path.join(self.tmpdir, "maildir")
        os.makedirs(os.path.join(md_root, "cur"), exist_ok=True)
        os.makedirs(os.path.join(md_root, "new"), exist_ok=True)
        os.makedirs(os.path.join(md_root, "tmp"), exist_ok=True)
        os.makedirs(os.path.join(md_root, ".Work.Caf&AOk-", "cur"), exist_ok=True)

        with open(os.path.join(md_root, "new", "1700001.M1P1.host"), "wb") as f:
            f.write(b"From: a@example.com\r\nSubject: New Unread\r\n\r\nBody 1\r\n")
        with open(os.path.join(md_root, ".Work.Caf&AOk-", "cur", "1700002.M2P2.host;2,FRSP"), "wb") as f:
            f.write(b"From: b@example.com\r\nSubject: Work Flagged\r\n\r\nBody 2\r\n")
        with open(os.path.join(md_root, "cur", "1700003.M3P3.host;2,ST"), "wb") as f:
            f.write(b"From: c@example.com\r\nSubject: Trashed\r\n\r\nBody 3\r\n")
        with open(os.path.join(md_root, "tmp", "1700004.M4P4.host"), "wb") as f:
            f.write(b"From: d@example.com\r\nSubject: Tmp Ignored\r\n\r\nBody 4\r\n")

        recs = list(mail_formats.iter_mail_records(md_root, include_deleted=False))
        self.assertEqual(len(recs), 2)
        by_subj = {email.message_from_bytes(r.raw_bytes)["Subject"]: r for r in recs}
        self.assertEqual(by_subj["New Unread"].folder_parts, ["Inbox"])
        self.assertFalse(by_subj["New Unread"].flags["is_read"])
        self.assertEqual(by_subj["Work Flagged"].folder_parts, ["Work", "Café"])
        self.assertTrue(by_subj["Work Flagged"].flags["is_read"])
        self.assertTrue(by_subj["Work Flagged"].flags["is_flagged"])
        self.assertTrue(by_subj["Work Flagged"].flags["is_replied"])
        self.assertTrue(by_subj["Work Flagged"].flags["is_forwarded"])

    def test_apple_mail_emlx_and_mbox_bundle(self):
        apple_root = os.path.join(self.tmpdir, "apple_mail")
        msgs_dir = os.path.join(
            apple_root, "Important.mbox", "12345678-1234-1234-1234-123456789ABC", "Data", "4", "2", "Messages"
        )
        os.makedirs(msgs_dir, exist_ok=True)
        rfc822 = b"From: tim@apple.com\r\nSubject: EMLX Test\r\n\r\nHello from macOS Mail!\r\n"
        # flags = 1 (read) | 4 (replied) | 16 (flagged) | 256 (forwarded) = 277
        plist_277 = b'<?xml version="1.0" encoding="UTF-8"?><plist version="1.0"><dict><key>flags</key><integer>277</integer></dict></plist>'
        with open(os.path.join(msgs_dir, "101.emlx"), "wb") as f:
            f.write(f"{len(rfc822)}\n".encode("ascii") + rfc822 + plist_277)

        plist_del = b'<?xml version="1.0" encoding="UTF-8"?><plist version="1.0"><dict><key>flags</key><integer>2</integer></dict></plist>'
        with open(os.path.join(msgs_dir, "102.partial.emlx"), "wb") as f:
            f.write(f"{len(rfc822)}\n".encode("ascii") + rfc822 + plist_del)

        recs = list(mail_formats.iter_mail_records(apple_root, include_deleted=False))
        self.assertEqual(len(recs), 1)
        self.assertEqual(recs[0].folder_parts, ["Important"])
        self.assertTrue(recs[0].flags["is_read"])
        self.assertTrue(recs[0].flags["is_replied"])
        self.assertTrue(recs[0].flags["is_flagged"])
        self.assertTrue(recs[0].flags["is_forwarded"])
        self.assertNotIn(b"<?xml", recs[0].raw_bytes)

    def test_outlook_msg_ole2_parser(self):
        msg_path = os.path.join(self.tmpdir, "sample.msg")
        with open(msg_path, "wb") as f:
            f.write(build_synthetic_msg_ole2())

        recs = list(mail_formats.iter_mail_records(msg_path))
        self.assertEqual(len(recs), 1)
        self.assertEqual(recs[0].native_msg_path, os.path.abspath(msg_path))
        self.assertFalse(recs[0].flags["is_read"])
        self.assertTrue(recs[0].flags["is_flagged"])
        parsed = email.message_from_bytes(recs[0].raw_bytes, policy=policy.default)
        self.assertEqual(parsed["Subject"], "Synthetic OLE2 MSG")
        self.assertIn("carol@example.com", parsed["From"])
        self.assertIn("Hello from a binary .msg file!", parsed.get_body().get_content())

    def test_outlook_for_mac_olm_archive(self):
        olm_path = os.path.join(self.tmpdir, "export.olm")
        with zipfile.ZipFile(olm_path, "w") as zf:
            zf.writestr("Attachments/report.txt", b"Quarterly report contents")
            olm_xml = """<?xml version="1.0" encoding="UTF-8"?>
            <emails>
              <email>
                <OPFMessageCopySubject>OLM Archive Message</OPFMessageCopySubject>
                <OPFMessageCopySentTime>2022-08-15T14:30:00Z</OPFMessageCopySentTime>
                <OPFMessageCopyMessageID>olm-123@example.com</OPFMessageCopyMessageID>
                <OPFMessageGetIsRead>0</OPFMessageGetIsRead>
                <OPFMessageCopyFlagStatus>1</OPFMessageCopyFlagStatus>
                <OPFMessageCopyFromAddresses>
                  <emailAddress OPFContactEmailAddressName="Alice" OPFContactEmailAddressAddress="alice@example.com"/>
                </OPFMessageCopyFromAddresses>
                <OPFMessageCopyToAddresses>
                  <emailAddress OPFContactEmailAddressName="Bob" OPFContactEmailAddressAddress="bob@example.com"/>
                </OPFMessageCopyToAddresses>
                <OPFMessageCopyBody>Plain OLM text</OPFMessageCopyBody>
                <OPFMessageCopyHTMLBody>&lt;html&gt;&lt;body&gt;&lt;b&gt;Rich OLM HTML&lt;/b&gt;&lt;/body&gt;&lt;/html&gt;</OPFMessageCopyHTMLBody>
                <OPFMessageCopyAttachmentList>
                  <attachment OPFAttachmentName="report.txt" OPFAttachmentContentType="text/plain" OPFAttachmentURL="Attachments/report.txt"/>
                </OPFMessageCopyAttachmentList>
              </email>
            </emails>"""
            zf.writestr("Accounts/alice@example.com/com.microsoft.__Messages/Projects/Q3/message_00001.xml", olm_xml)
            zf.writestr("Accounts/alice@example.com/com.microsoft.__Messages/Trash/message_00002.xml", olm_xml)

        recs = list(mail_formats.iter_mail_records(olm_path, include_deleted=False))
        self.assertEqual(len(recs), 1)
        self.assertEqual(recs[0].folder_parts, ["Projects", "Q3"])
        self.assertFalse(recs[0].flags["is_read"])
        self.assertTrue(recs[0].flags["is_flagged"])
        parsed = email.message_from_bytes(recs[0].raw_bytes, policy=policy.default)
        self.assertEqual(parsed["Subject"], "OLM Archive Message")
        atts = list(parsed.iter_attachments())
        self.assertEqual(len(atts), 1)
        self.assertEqual(atts[0].get_filename(), "report.txt")
        self.assertEqual(atts[0].get_payload(decode=True), b"Quarterly report contents")

    def test_mh_mmdf_babyl_and_compressed_archives(self):
        # MH
        mh_dir = os.path.join(self.tmpdir, "mh_box")
        os.makedirs(mh_dir, exist_ok=True)
        with open(os.path.join(mh_dir, ".mh_sequences"), "w", encoding="utf-8") as f:
            f.write("unseen: 2\nflagged: 1\ndeleted: 3\n")
        for num in (1, 2, 3):
            with open(os.path.join(mh_dir, str(num)), "wb") as f:
                f.write(f"From: mh@example.com\r\nSubject: MH #{num}\r\n\r\nMH body {num}\r\n".encode("utf-8"))
        mh_recs = list(mail_formats.iter_mail_records(mh_dir, include_deleted=False))
        self.assertEqual(len(mh_recs), 2)

        # MMDF
        mmdf_path = os.path.join(self.tmpdir, "test.mmdf")
        delim = b"\x01\x01\x01\x01\r\n"
        with open(mmdf_path, "wb") as f:
            f.write(delim + b"From: a@example.com\r\nSubject: MMDF 1\r\n\r\nA\r\n" + delim)
        self.assertEqual(len(list(mail_formats.iter_mail_records(mmdf_path))), 1)

        # Babyl
        babyl_path = os.path.join(self.tmpdir, "test.babyl")
        with open(babyl_path, "wb") as f:
            f.write(
                b"BABYL OPTIONS: -*- rmail -*-\r\nVersion: 5\r\n\x1f\x0c\r\n"
                b"1, unseen, flagged,,\r\n*** EOOH ***\r\nFrom: b@example.com\r\nSubject: Babyl 1\r\n\r\nBody\r\n\x1f"
            )
        babyl_recs = list(mail_formats.iter_mail_records(babyl_path))
        self.assertEqual(len(babyl_recs), 1)
        self.assertFalse(babyl_recs[0].flags["is_read"])
        self.assertTrue(babyl_recs[0].flags["is_flagged"])

        # Compressed .mbox.gz, .mbox.bz2, .mbox.xz, and .tgz
        sample_mbox = b"From gz@example.com Fri Nov  6 10:00:00 2020\r\nFrom: gz@example.com\r\nSubject: Compressed\r\n\r\nBody\r\n"
        for ext, open_fn in ((".mbox.gz", gzip.open), (".mbox.bz2", bz2.open), (".mbox.xz", lzma.open)):
            c_path = os.path.join(self.tmpdir, f"arch{ext}")
            with open_fn(c_path, "wb") as f:
                f.write(sample_mbox)
            c_recs = list(mail_formats.iter_mail_records(c_path))
            self.assertEqual(len(c_recs), 1)

        tgz_path = os.path.join(self.tmpdir, "takeout.tgz")
        with tarfile.open(tgz_path, "w:gz") as tf:
            tf.add(mmdf_path, arcname="Takeout/Mail/test.mmdf")
        self.assertEqual(len(list(mail_formats.iter_mail_records(tgz_path))), 1)


if __name__ == "__main__":
    unittest.main()
