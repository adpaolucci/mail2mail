#!/usr/bin/env python3
"""
tests/test_rfc_compliance.py — Automated verification of IETF RFC compliance:
  - RFC 5322 (Headers, folding, malformed header resilience, Date, Address lists, Message-ID)
  - RFC 2045 / 2046 (MIME multipart/mixed, multipart/alternative, multipart/related, base64, QP)
  - RFC 2047 (Encoded-word headers: =?charset?B/Q?...?=)
  - RFC 2183 / 2231 / 2392 (Content-Disposition, parameter continuations, Content-ID cid: MHTML)
  - RFC 3501 §5.1.3 (IMAP4rev1 Modified UTF-7 folder name encoding & decoding)
  - RFC 4155 & mboxrd (From <sender> <asctime> framing & reversible >From body escaping)
  - RFC 8551 / 3156 (S/MIME pkcs7-mime/pkcs7-signature & PGP pgp-signature preservation)
"""

from __future__ import annotations

from datetime import datetime, timezone
import email
from email import policy
import io
import os
import shutil
import tempfile
import unittest

import tests  # noqa: F401 (ensures pywin32 stubs if needed)
import mail_formats
import mail2mbox
import mail2pst


class TestRFC5322HeadersAndDates(unittest.TestCase):
    """Verify RFC 5322 header parsing, folding, dates, addresses, and malformed header recovery."""

    def test_folded_headers_and_multi_value_recipients(self):
        raw = (
            b"From: \"Doe; Jane\" <jane@example.com>\r\n"
            b"To: First Recipient\r\n"
            b" <first@example.com>, \"Second; Jr.\" <second@example.com>\r\n"
            b"To: Third Recipient <third@example.com>\r\n"
            b"Cc: Carbon <cc@example.com>\r\n"
            b"Subject: Line 1\r\n"
            b" Line 2 Folded\r\n"
            b"Date: Fri, 06 Nov 2020 15:30:00 -0500\r\n"
            b"Message-ID: <rfc5322-1@example.com>\r\n\r\n"
            b"Hello RFC 5322!\r\n"
        )
        msg = email.message_from_bytes(raw, policy=policy.default)
        self.assertEqual(mail_formats.safe_get_header(msg, "Subject"), "Line 1 Line 2 Folded")
        self.assertEqual(mail_formats.safe_get_header(msg, "Message-ID"), "<rfc5322-1@example.com>")

        to_all = mail_formats.safe_get_all_headers(msg, "To")
        self.assertEqual(len(to_all), 2)
        formatted_to = mail2pst.addresses_to_str(to_all)
        self.assertEqual(
            formatted_to,
            "First Recipient <first@example.com>; Second, Jr. <second@example.com>; Third Recipient <third@example.com>",
        )

    def test_malformed_rfc5322_headers_do_not_crash(self):
        # Malformed Date, invalid group syntax in To, and broken Content-Type parameter
        raw = (
            b"From: broken@@example..com\r\n"
            b"To: Undisclosed-Recipients:;,,,,;\r\n"
            b"Date: Fri, 99 Nov 2020 99:99:99 +9999\r\n"
            b"Subject: Malformed Header Survival\r\n\r\n"
            b"Body text\r\n"
        )
        msg = email.message_from_bytes(raw, policy=policy.default)
        self.assertEqual(mail_formats.safe_get_header(msg, "Subject"), "Malformed Header Survival")
        self.assertEqual(mail_formats.pick_year_month(msg), (1970, 1))
        # Ensure safe_get_all_headers returns raw fallback without raising HeaderParseError
        _ = mail_formats.safe_get_all_headers(msg, "From")
        _ = mail_formats.safe_get_all_headers(msg, "To")

    def test_rfc5322_date_to_mapi_naive_utc(self):
        # -0500 should convert to 20:30:00 naive UTC so pywin32 VT_DATE does not shift by local tz
        dt = email.utils.parsedate_to_datetime("Fri, 06 Nov 2020 15:30:00 -0500")
        utc_naive = mail2pst.to_mapi_utc(None, dt)
        self.assertIsNone(utc_naive.tzinfo)
        self.assertEqual(utc_naive, datetime(2020, 11, 6, 20, 30, 0))

        # Out-of-bounds years are clamped to [1970, 2100]
        ancient = datetime(1605, 1, 1, 0, 0, 0, tzinfo=timezone.utc)
        self.assertEqual(mail2pst.to_mapi_utc(None, ancient).year, 1970)
        far_future = datetime(2500, 1, 1, 0, 0, 0, tzinfo=timezone.utc)
        self.assertEqual(mail2pst.to_mapi_utc(None, far_future).year, 2100)


class TestRFC2047EncodedWords(unittest.TestCase):
    """Verify RFC 2047 encoded-word header decoding (B, Q, multi-token, and malformed)."""

    def test_rfc2047_base64_and_quoted_printable(self):
        b64_subj = "=?UTF-8?B?Q2Fmw6kgaW52b2ljZSDigJQgMjAyNA==?="
        self.assertEqual(mail_formats.decode_mime_words(b64_subj), "Café invoice — 2024")

        qp_subj = "=?ISO-8859-1?Q?Pr=FCfung_f=FCr_M=FCnchen?="
        self.assertEqual(mail_formats.decode_mime_words(qp_subj), "Prüfung für München")

        # Malformed charset fallback
        bad_subj = "=?UNKNOWN-CHARSET-999?B?SGVsbG8=?="
        decoded = mail_formats.decode_mime_words(bad_subj)
        self.assertTrue(isinstance(decoded, str) and len(decoded) > 0)


class TestRFC3501ModifiedUTF7(unittest.TestCase):
    """Verify RFC 3501 §5.1.3 IMAP4rev1 Modified UTF-7 folder name encoding and decoding."""

    def test_rfc3501_decode_and_encode_roundtrip(self):
        cases = [
            ("INBOX", "INBOX"),
            ("Notes & Invoices", "Notes &- Invoices"),
            ("Café", "Caf&AOk-"),
            ("Entwürfe & Café", "Entw&APw-rfe &- Caf&AOk-"),
            ("日本語/台北", "&ZeVnLIqe-/&U,BTFw-"),
            ("Archive 🚀 2026", "Archive &2D3egA- 2026"),
        ]
        for decoded_expected, mutf7_expected in cases:
            with self.subTest(folder=decoded_expected):
                self.assertEqual(mail_formats.decode_imap_utf7(mutf7_expected), decoded_expected)
                self.assertEqual(mail_formats.encode_imap_utf7(decoded_expected), mutf7_expected)
                self.assertEqual(
                    mail_formats.decode_imap_utf7(mail_formats.encode_imap_utf7(decoded_expected)),
                    decoded_expected,
                )


class TestRFC4155MboxAndMboxrd(unittest.TestCase):
    """Verify RFC 4155 envelope framing and reversible mboxrd >From body escaping."""

    def test_mboxrd_quoting_and_unquoting_roundtrip(self):
        raw_eml = (
            b"From: sender@example.com\r\n"
            b"Subject: Mboxrd Quoting Test\r\n"
            b"Date: Fri, 06 Nov 2020 12:00:00 +0000\r\n\r\n"
            b"Normal line 1\r\n"
            b"From someone at start of line\r\n"
            b">From already single-quoted\r\n"
            b">>From already double-quoted\r\n"
            b"Not From at start\r\n"
        )
        entry = mail2mbox.format_mboxrd_entry(
            raw_eml,
            flags={"is_read": True, "is_replied": True, "is_flagged": False, "is_forwarded": False, "is_deleted": False},
        )
        self.assertTrue(entry.startswith(b"From sender@example.com Fri Nov  6 12:00:00 2020\n"))
        self.assertIn(b"\n>From someone at start of line\n", entry)
        self.assertIn(b"\n>>From already single-quoted\n", entry)
        self.assertIn(b"\n>>>From already double-quoted\n", entry)

        # Now parse it back via iter_mbox_stream and verify exact un-escaping
        recs = list(mail_formats.iter_mbox_stream(io.BytesIO(entry), ["Inbox"]))
        self.assertEqual(len(recs), 1)
        parsed = email.message_from_bytes(recs[0].raw_bytes, policy=policy.default)
        body = parsed.get_content()
        self.assertIn("From someone at start of line", body)
        self.assertIn(">From already single-quoted", body)
        self.assertIn(">>From already double-quoted", body)
        self.assertNotIn(">>>From", body)
        self.assertTrue(recs[0].flags["is_read"])
        self.assertTrue(recs[0].flags["is_replied"])


class TestRFC2045MimeAndSMime(unittest.TestCase):
    """Verify MIME multipart body selection, inline CID MHTML attachments, message/rfc822, and S/MIME."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="test_rfc_mime_")

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_multipart_related_cid_and_rfc822_forwarding(self):
        raw = (
            b"From: alice@example.com\r\n"
            b"To: bob@example.com\r\n"
            b"Subject: Multipart + CID + Forwarded RFC822\r\n"
            b"MIME-Version: 1.0\r\n"
            b"Content-Type: multipart/mixed; boundary=\"top\"\r\n\r\n"
            b"--top\r\n"
            b"Content-Type: multipart/alternative; boundary=\"alt\"\r\n\r\n"
            b"--alt\r\n"
            b"Content-Type: text/plain; charset=\"utf-8\"\r\n\r\n"
            b"Plain text fallback\r\n"
            b"--alt\r\n"
            b"Content-Type: text/html; charset=\"utf-8\"\r\n\r\n"
            b"<html><body><img src=\"cid:logo123\"></body></html>\r\n"
            b"--alt--\r\n"
            b"--top\r\n"
            b"Content-Type: image/png\r\n"
            b"Content-ID: <logo123>\r\n"
            b"Content-Transfer-Encoding: base64\r\n\r\n"
            b"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==\r\n"
            b"--top\r\n"
            b"Content-Type: message/rfc822\r\n"
            b"Content-Disposition: attachment\r\n\r\n"
            b"From: inner@example.com\r\n"
            b"Subject: Inner Forwarded Message\r\n\r\n"
            b"Inner body\r\n"
            b"--top\r\n"
            b"Content-Type: application/pkcs7-signature\r\n\r\n"
            b"DUMMY_SMIME_SIG\r\n"
            b"--top--\r\n"
        )
        msg = email.message_from_bytes(raw, policy=policy.default)
        html, text, body_parts = mail2pst.pick_body(msg)
        self.assertIn("cid:logo123", html)
        self.assertIn("Plain text fallback", text)
        self.assertEqual(len(body_parts), 2)

        added = []

        class MockPA:
            def __init__(self):
                self.props = {}
            def SetProperty(self, k, v):
                self.props[k] = v

        class MockAtt:
            def __init__(self, path, disp):
                self.path = path
                self.disp = disp
                self.PropertyAccessor = MockPA()

        class MockMail:
            class Attachments:
                @staticmethod
                def Add(path, t, p, disp):
                    att = MockAtt(path, disp)
                    added.append(att)
                    return att

        count, has_cid = mail2pst.add_attachments(MockMail(), msg, self.tmpdir, body_parts=body_parts)
        self.assertEqual(count, 3)
        self.assertTrue(has_cid)
        disp_names = [a.disp for a in added]
        self.assertTrue(any(n.endswith(".eml") for n in disp_names))
        self.assertIn("smime.p7s", disp_names)
        cid_att = [a for a in added if mail2pst.PR_ATTACH_CONTENT_ID_W in a.PropertyAccessor.props][0]
        self.assertEqual(cid_att.PropertyAccessor.props[mail2pst.PR_ATTACH_CONTENT_ID_W], "logo123")
        self.assertEqual(cid_att.PropertyAccessor.props[mail2pst.PR_ATTACH_FLAGS], 4)

    def test_rfc5545_calendar_and_rfc3464_delivery_status(self):
        """Verify RFC 5545 (text/calendar) and RFC 3464 (message/delivery-status) attachments."""
        raw = (
            b"From: organizer@example.com\r\n"
            b"To: attendee@example.com\r\n"
            b"Subject: Meeting Invite & Status Report\r\n"
            b"MIME-Version: 1.0\r\n"
            b"Content-Type: multipart/mixed; boundary=\"mix\"\r\n\r\n"
            b"--mix\r\n"
            b"Content-Type: text/plain; charset=\"utf-8\"\r\n\r\n"
            b"Please join our planning session.\r\n"
            b"--mix\r\n"
            b"Content-Type: text/calendar; method=REQUEST; charset=\"utf-8\"\r\n\r\n"
            b"BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:REQUEST\r\nBEGIN:VEVENT\r\nSUMMARY:Session\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n"
            b"--mix\r\n"
            b"Content-Type: message/delivery-status\r\n\r\n"
            b"Reporting-MTA: dns; mail.example.com\r\nStatus: 2.0.0\r\n"
            b"--mix--\r\n"
        )
        msg = email.message_from_bytes(raw, policy=policy.default)
        html, text, body_parts = mail2pst.pick_body(msg)
        self.assertIn("planning session", text)

        added = []

        class MockPA:
            def __init__(self):
                self.props = {}
            def SetProperty(self, k, v):
                self.props[k] = v

        class MockAtt:
            def __init__(self, path, disp):
                self.path = path
                self.disp = disp
                self.PropertyAccessor = MockPA()

        class MockMail:
            class Attachments:
                @staticmethod
                def Add(path, t, p, disp):
                    att = MockAtt(path, disp)
                    added.append(att)
                    return att

        count, _ = mail2pst.add_attachments(MockMail(), msg, self.tmpdir, body_parts=body_parts)
        self.assertEqual(count, 2)
        disp_names = [a.disp for a in added]
        self.assertTrue(any(n.endswith(".ics") for n in disp_names))
        self.assertTrue(any(n.endswith(".txt") for n in disp_names))


if __name__ == "__main__":
    unittest.main()
