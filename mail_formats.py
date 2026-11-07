#!/usr/bin/env python3
"""
mail_formats.py

Unified multi-format email & mailbox reader and shared utilities for mail2mail
(mail2eml, mail2mbox, mail2pst).

Supports reading and streaming messages with folder hierarchy and status flags from:
  1. Standard MBOX files (mboxo / mboxrd) & Thunderbird ImapMail/Mail (.sbd, mUTF-7)
  2. Maildir & Maildir++ (cur/, new/, tmp/, :2,FRS flags, .Folder.Subfolder, Thunderbird Maildir)
  3. Apple Mail (.emlx, .partial.emlx, and macOS *.mbox / *.imapmbox directory bundles)
  4. Outlook .msg files (pure-Python MS-CFB OLE2 + MS-OXMSG parser, plus native path tracking)
  5. Outlook for Mac .olm archives (ZIP + XML message & attachment reconstruction)
  6. Outlook .pst archives (via Outlook MAPI on Windows)
  7. MH / Sylpheed / Claws Mail directories (numbered files + .mh_sequences)
  8. MMDF (\x01\x01\x01\x01 delimited) and Emacs Babyl (BABYL OPTIONS:) mailboxes
  9. Standard .eml files & directory trees
 10. Compressed files & archives (.gz, .bz2, .xz, .zip, .tar, .tar.gz, .tgz, .tar.bz2, .tar.xz)
"""

from __future__ import annotations

import base64
import bz2
import email
import email.utils
import gzip
import io
import lzma
import mimetypes
import os
import plistlib
import re
import shutil
import struct
import tarfile
import tempfile
import urllib.parse
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from email import policy
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import parsedate_to_datetime
from typing import BinaryIO, Dict, Iterator, List, Optional, Set, Tuple
import xml.etree.ElementTree as ET


# ----------------------- Constants & Regexes -----------------------

MSG_FLAG_READ      = 0x0001
MSG_FLAG_REPLIED   = 0x0002
MSG_FLAG_MARKED    = 0x0004  # Flagged / Starred in Thunderbird
MSG_FLAG_EXPUNGED  = 0x0008  # Deleted (awaiting folder compaction)
MSG_FLAG_FORWARDED = 0x1000  # Forwarded

FORBIDDEN_CHARS = r'\/:*?"<>|'
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

THUNDERBIRD_SKIP_SUFFIXES = (
    ".msf", ".dat", ".wdseml", ".mozmsgs", ".html", ".json", ".sqlite", ".mab", ".db", ".log", ".ini"
)
THUNDERBIRD_SKIP_FILES = {
    "msgfilterrules.dat",
    "filterlog.html",
    "popstate.dat",
    "folderlock",
    "foldertree.json",
    "virtualfolders.dat",
    "info.plist",
    "table_of_contents",
    "content_index",
    ".ds_store",
    "thumbs.db",
    "desktop.ini",
    ".mh_sequences",
    ".sylpheed_mark",
    ".sylpheed_cache",
    ".claws_mark",
    ".claws_cache",
    ".xmhcache",
    "dovecot.index",
    "dovecot.index.cache",
    "dovecot.index.log",
    "dovecot-uidlist",
    "dovecot-keywords",
    "maildirsize",
    "subscriptions",
}

APPLE_UUID_RE = re.compile(
    r"^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$"
)
MBOX_FROM_RE = re.compile(
    rb"^From \S+ (?:Mon|Tue|Wed|Thu|Fri|Sat|Sun) "
    rb"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) "
    rb"(?: \d|\d\d) \d\d:\d\d:\d\d \d{4}"
)
HEADER_START_RE = re.compile(
    rb"^(?:From|To|Cc|Bcc|Date|Subject|Message-ID|Received|Return-Path|"
    rb"MIME-Version|Content-Type|Delivered-To|X-[A-Za-z0-9-]+|Status)\s*:",
    re.IGNORECASE,
)
MAILDIR_FLAG_RE = re.compile(r"[:;!✓]2,([A-Za-z]*)$")
OLE2_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
PST_MAGIC = b"!BDN"


# ----------------------- Shared Path, Name & Header Helpers -----------------------

def win_long_path(path: str) -> str:
    """Prefix Windows paths with \\\\?\\ when approaching MAX_PATH (260 chars)."""
    if os.name != "nt" or not path:
        return path
    if path.startswith("\\\\?\\"):
        return path
    abs_p = os.path.abspath(path)
    if len(abs_p) >= 240:
        if abs_p.startswith("\\\\"):
            return "\\\\?\\UNC\\" + abs_p[2:]
        return "\\\\?\\" + abs_p
    return path


def decode_imap_utf7(s: str) -> str:
    """
    Decode IMAP Modified UTF-7 folder names (RFC 3501 Section 5.1.3).
    Example: 'Notes &- Invoices' -> 'Notes & Invoices', 'Caf&AOk-' -> 'Café'
    """
    if not s or "&" not in s:
        return s
    out: List[str] = []
    i = 0
    n = len(s)
    while i < n:
        if s[i] == "&":
            end = s.find("-", i + 1)
            if end == -1:
                out.append(s[i:])
                break
            if end == i + 1:
                out.append("&")
            else:
                b64 = s[i + 1 : end].replace(",", "/")
                pad = (4 - len(b64) % 4) % 4
                try:
                    raw = base64.b64decode(b64 + ("=" * pad))
                    out.append(raw.decode("utf-16-be", errors="replace"))
                except Exception:
                    out.append(s[i : end + 1])
            i = end + 1
        else:
            out.append(s[i])
            i += 1
    return "".join(out)


decode_modified_utf7 = decode_imap_utf7


def clean_folder_segment(name: str) -> str:
    """Decode IMAP mUTF-7 and strip container suffixes (.sbd, .mbox, etc.)."""
    s = name.strip()
    for suffix in (".sbd", ".mbox", ".imapmbox", ".popmbox"):
        if s.lower().endswith(suffix) and len(s) > len(suffix):
            s = s[: -len(suffix)]
            break
    s = decode_imap_utf7(s).strip()
    return s or "Folder"


def sanitize_name(s: str, default: str = "item", max_len: int = 120) -> str:
    """
    Decode IMAP Modified UTF-7 and return a Windows/POSIX/Outlook-safe folder or filename.
    Preserves file extensions when truncating to max_len.
    """
    if not isinstance(s, str):
        s = str(s or "")
    s = decode_imap_utf7(s)
    s = re.sub(r'[\x00-\x1f\x7f\\/:*?"<>|]+', "_", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    if not s:
        s = default
    if s.upper().split(".")[0] in WINDOWS_RESERVED_NAMES:
        s = f"_{s}"
    if len(s) > max_len:
        base, ext = os.path.splitext(s)
        if ext and len(ext) <= 12:
            keep = max(1, max_len - len(ext))
            s = base[:keep].rstrip(" .") + ext
        else:
            s = s[:max_len].strip(" .")
    return s or default


def decode_mime_words(s) -> str:
    """Safely decode RFC 2047 encoded-word headers into a clean Unicode string."""
    if not s:
        return ""
    try:
        return str(make_header(decode_header(str(s))))
    except Exception:
        return str(s)


def safe_get_header(msg, name: str, default: str = "") -> str:
    """
    Safely read and decode a single header from an email.message.Message without
    failing on malformed RFC 5322 headers (HeaderParseError / IndexError / ValueError).
    """
    try:
        v = msg.get(name)
        if v is not None:
            return decode_mime_words(v)
    except Exception:
        pass
    target = name.lower()
    try:
        for k, raw_v in msg.raw_items():
            if k.lower() == target:
                return decode_mime_words(raw_v)
    except Exception:
        pass
    return default


def safe_get_all_headers(msg, name: str) -> List[str]:
    """Safely read all instances of a header from an email.message.Message."""
    try:
        vals = msg.get_all(name)
        if vals:
            return [decode_mime_words(v) for v in vals if v is not None]
    except Exception:
        pass
    target = name.lower()
    out: List[str] = []
    try:
        for k, raw_v in msg.raw_items():
            if k.lower() == target and raw_v is not None:
                out.append(decode_mime_words(raw_v))
    except Exception:
        pass
    return out


def pick_year_month(msg) -> Tuple[int, int]:
    """Extract (year, month) from the Date header; default to (1970, 1) if missing/invalid."""
    year = 1970
    month = 1
    try:
        d = safe_get_header(msg, "Date", "").strip()
        if d:
            dt = parsedate_to_datetime(d)
            if dt is not None and 1900 <= dt.year <= 2100:
                year = dt.year
                month = dt.month
    except Exception:
        pass
    return year, month


# ----------------------- Data Model & Status Flags -----------------------

@dataclass
class MailRecord:
    """Unified representation of an email extracted from any supported format."""
    raw_bytes: bytes
    folder_parts: List[str] = field(default_factory=list)
    source_name: str = "message.eml"
    is_deleted: bool = False
    flags: Dict[str, bool] = field(default_factory=dict)
    native_msg_path: Optional[str] = None


def default_flags(
    is_read: bool = True,
    is_replied: bool = False,
    is_flagged: bool = False,
    is_forwarded: bool = False,
    is_deleted: bool = False,
) -> Dict[str, bool]:
    return {
        "is_read": is_read,
        "is_replied": is_replied,
        "is_flagged": is_flagged,
        "is_forwarded": is_forwarded,
        "is_deleted": is_deleted,
    }


def flags_to_mozilla_status_hex(flags: Dict[str, bool]) -> str:
    """Convert normalized boolean flags into a 4-digit hex X-Mozilla-Status value."""
    bits = 0
    if flags.get("is_read", True):
        bits |= MSG_FLAG_READ
    if flags.get("is_replied", False):
        bits |= MSG_FLAG_REPLIED
    if flags.get("is_flagged", False):
        bits |= MSG_FLAG_MARKED
    if flags.get("is_deleted", False):
        bits |= MSG_FLAG_EXPUNGED
    if flags.get("is_forwarded", False):
        bits |= MSG_FLAG_FORWARDED
    return f"{bits:04x}"


def extract_flags_from_headers(raw_bytes: bytes) -> Tuple[Dict[str, bool], bool]:
    """
    Inspect raw RFC 822 header block for X-Mozilla-Status, Status, X-Status, or X-Message-Flag.
    Returns (flags_dict, found_explicit_header).
    """
    sep_crlf = raw_bytes.find(b"\r\n\r\n")
    sep_lf = raw_bytes.find(b"\n\n")
    if sep_crlf != -1 and (sep_lf == -1 or sep_crlf < sep_lf):
        hdr_block = raw_bytes[:sep_crlf]
    elif sep_lf != -1:
        hdr_block = raw_bytes[:sep_lf]
    else:
        hdr_block = raw_bytes[:4096]

    text = hdr_block.decode("latin1", errors="replace")
    moz_match = re.search(r"(?im)^X-Mozilla-Status\s*:\s*([0-9A-Fa-f]{1,8})", text)
    if moz_match:
        try:
            val = int(moz_match.group(1), 16)
            return (
                default_flags(
                    is_read=bool(val & MSG_FLAG_READ),
                    is_replied=bool(val & MSG_FLAG_REPLIED),
                    is_flagged=bool(val & MSG_FLAG_MARKED),
                    is_deleted=bool(val & MSG_FLAG_EXPUNGED),
                    is_forwarded=bool(val & MSG_FLAG_FORWARDED),
                ),
                True,
            )
        except ValueError:
            pass

    status_match = re.search(r"(?im)^Status\s*:\s*([^\r\n]+)", text)
    xstatus_match = re.search(r"(?im)^X-Status\s*:\s*([^\r\n]+)", text)
    flag_match = re.search(r"(?im)^X-Message-Flag\s*:\s*([^\r\n]+)", text)
    if status_match or xstatus_match or flag_match:
        st = (status_match.group(1) if status_match else "").upper()
        xst = (xstatus_match.group(1) if xstatus_match else "").upper()
        mfl = (flag_match.group(1) if flag_match else "").strip()
        is_read = ("R" in st) if st else True
        if "U" in st:
            is_read = False
        return (
            default_flags(
                is_read=is_read,
                is_replied=("A" in xst),
                is_flagged=("F" in xst) or bool(mfl),
                is_deleted=("D" in xst or "D" in st),
                is_forwarded=("W" in xst),
            ),
            True,
        )

    return default_flags(is_read=True), False


def extract_message_flags(msg) -> Tuple[bool, bool, bool, bool, bool]:
    """
    Inspect X-Mozilla-Status, Status, X-Status, and X-Message-Flag on a parsed Message.
    Returns (is_read, is_flagged, is_replied, is_forwarded, is_expunged).
    """
    is_read = True
    is_flagged = False
    is_replied = False
    is_forwarded = False
    is_expunged = False

    tb_status = safe_get_header(msg, "X-Mozilla-Status", "").strip()
    if tb_status:
        try:
            val = int(tb_status, 16)
            return (
                bool(val & MSG_FLAG_READ),
                bool(val & MSG_FLAG_MARKED),
                bool(val & MSG_FLAG_REPLIED),
                bool(val & MSG_FLAG_FORWARDED),
                bool(val & MSG_FLAG_EXPUNGED),
            )
        except ValueError:
            pass

    status_hdr = safe_get_header(msg, "Status", "").upper()
    if status_hdr:
        is_read = ("R" in status_hdr)

    x_status = safe_get_header(msg, "X-Status", "").upper()
    if "F" in x_status:
        is_flagged = True
    if "A" in x_status:
        is_replied = True
    if "W" in x_status:
        is_forwarded = True
    if "D" in x_status or "D" in status_hdr:
        is_expunged = True

    if safe_get_header(msg, "X-Message-Flag", "").strip():
        is_flagged = True

    return is_read, is_flagged, is_replied, is_forwarded, is_expunged


def update_status_headers(
    raw_header_bytes: bytes,
    is_read: bool,
    is_flagged: bool,
    is_replied: bool,
    is_forwarded: bool = False,
) -> bytes:
    """
    Directly filter and update Status, X-Status, and X-Message-Flag headers at the byte level.
    Preserves 100% of MIME structure, DKIM signatures, and line folding.
    """
    lines = raw_header_bytes.splitlines(keepends=True)
    filtered_lines = []
    skip_continuation = False
    target_prefixes = (b"status:", b"x-status:", b"x-message-flag:")

    for line in lines:
        lower = line.lower()
        if any(lower.startswith(prefix) for prefix in target_prefixes):
            skip_continuation = True
            continue
        if skip_continuation and (line.startswith(b" ") or line.startswith(b"\t")):
            continue
        skip_continuation = False
        filtered_lines.append(line)

    if filtered_lines and not filtered_lines[-1].endswith((b"\r\n", b"\n", b"\r")):
        filtered_lines[-1] = filtered_lines[-1] + b"\r\n"

    status_hdr = b"Status: RO\r\n" if is_read else b"Status: O\r\n"
    filtered_lines.append(status_hdr)

    x_status_val = b""
    if is_flagged:
        x_status_val += b"F"
    if is_replied:
        x_status_val += b"A"
    if is_forwarded:
        x_status_val += b"W"
    if x_status_val:
        filtered_lines.append(b"X-Status: " + x_status_val + b"\r\n")

    if is_flagged:
        filtered_lines.append(b"X-Message-Flag: Follow up\r\n")

    return b"".join(filtered_lines)


def normalize_eml_bytes_with_flags(raw_bytes: bytes, flags: Dict[str, bool]) -> bytes:
    """Split raw RFC 822 message into headers + body and apply update_status_headers."""
    sep_crlf = raw_bytes.find(b"\r\n\r\n")
    sep_lf = raw_bytes.find(b"\n\n")
    if sep_crlf != -1 and (sep_lf == -1 or sep_crlf < sep_lf):
        hdr = raw_bytes[:sep_crlf]
        body = raw_bytes[sep_crlf + 4 :]
    elif sep_lf != -1:
        hdr = raw_bytes[:sep_lf]
        body = raw_bytes[sep_lf + 2 :]
    else:
        hdr = raw_bytes
        body = b""

    hdr_flags, has_hdr = extract_flags_from_headers(hdr)
    is_read = flags.get("is_read", hdr_flags["is_read"]) if not has_hdr else hdr_flags["is_read"]
    is_flagged = flags.get("is_flagged", False) or hdr_flags["is_flagged"]
    is_replied = flags.get("is_replied", False) or hdr_flags["is_replied"]
    is_forwarded = flags.get("is_forwarded", False) or hdr_flags["is_forwarded"]

    norm_hdr = update_status_headers(hdr, is_read, is_flagged, is_replied, is_forwarded)
    return norm_hdr.rstrip(b"\r\n") + b"\r\n\r\n" + body


def inject_mozilla_status_bytes(raw_bytes: bytes, flags: Dict[str, bool]) -> bytes:
    """
    Ensure X-Mozilla-Status (and Status/X-Status) headers exist in raw_bytes
    so downstream EML, MBOX, and PST tools preserve Read/Replied/Flagged/Forwarded state.
    """
    if not raw_bytes:
        return raw_bytes
    _, has_hdr = extract_flags_from_headers(raw_bytes)
    if has_hdr:
        return raw_bytes

    nl = b"\r\n" if b"\r\n" in raw_bytes[:512] else b"\n"
    moz_hex = flags_to_mozilla_status_hex(flags).encode("ascii")
    status_val = b"RO" if flags.get("is_read", True) else b"O"
    xstatus_chars = []
    if flags.get("is_replied"):
        xstatus_chars.append("A")
    if flags.get("is_flagged"):
        xstatus_chars.append("F")
    if flags.get("is_forwarded"):
        xstatus_chars.append("W")
    if flags.get("is_deleted"):
        xstatus_chars.append("D")

    prefix = b"X-Mozilla-Status: " + moz_hex + nl + b"Status: " + status_val + nl
    if xstatus_chars:
        prefix += b"X-Status: " + "".join(xstatus_chars).encode("ascii") + nl

    if raw_bytes.startswith(b"From "):
        first_nl = raw_bytes.find(b"\n")
        if first_nl != -1:
            return raw_bytes[: first_nl + 1] + prefix + raw_bytes[first_nl + 1 :]
    return prefix + raw_bytes


# ----------------------- 1. Maildir & Maildir++ -----------------------

def parse_maildir_flags(filename: str, in_new: bool = False) -> Tuple[Dict[str, bool], bool]:
    """
    Parse Maildir filename flags (e.g. '1700000.M123P456.host:2,FRS').
    Returns (flags, has_explicit_flags).
    """
    m = MAILDIR_FLAG_RE.search(filename)
    if m:
        fstr = m.group(1)
        return (
            default_flags(
                is_read=("S" in fstr),
                is_replied=("R" in fstr),
                is_flagged=("F" in fstr),
                is_forwarded=("P" in fstr),
                is_deleted=("T" in fstr),
            ),
            True,
        )
    if in_new:
        return default_flags(is_read=False), True
    return default_flags(is_read=True), False


def normalize_maildir_folder_parts(rel_parts: List[str]) -> List[str]:
    """
    Convert relative path segments inside a Maildir / Maildir++ / Thunderbird Maildir
    into logical folder parts.
    """
    out: List[str] = []
    for part in rel_parts:
        low = part.lower()
        if low in ("cur", "new", "tmp"):
            continue
        if part.startswith(".") and len(part) > 1:
            subparts = [clean_folder_segment(p) for p in part.lstrip(".").split(".") if p]
            if len(subparts) > 1 and subparts[0].upper() == "INBOX":
                subparts = subparts[1:]
            out.extend(subparts)
        else:
            out.append(clean_folder_segment(part))
    return out if out else ["Inbox"]


# ----------------------- 2. Apple Mail (.emlx & .mbox Bundles) -----------------------

def parse_emlx_bytes(data: bytes) -> Tuple[bytes, Dict[str, bool]]:
    """
    Parse an Apple Mail .emlx or .partial.emlx byte buffer.
    Strips the leading byte-count line and trailing XML <plist>, returning
    (rfc822_bytes, flags_dict). Resilient against CRLF/LF line-ending shifts.
    """
    flags = default_flags(is_read=True)
    if not data:
        return b"", flags

    first_nl = data.find(b"\n")
    if first_nl == -1:
        return data, flags

    first_line = data[:first_nl].strip()
    if not first_line.isdigit():
        hdr_flags, _ = extract_flags_from_headers(data)
        return data, hdr_flags

    declared_len = int(first_line)
    payload_start = first_nl + 1
    candidate_end = payload_start + declared_len

    plist_bytes = b""
    rfc822_bytes = b""

    if candidate_end <= len(data):
        remainder = data[candidate_end:].lstrip()
        if remainder.startswith((b"<?xml", b"<!DOCTYPE", b"<plist")):
            rfc822_bytes = data[payload_start:candidate_end]
            plist_bytes = remainder

    if not rfc822_bytes:
        tail_search_start = max(payload_start, len(data) - 16384)
        tail = data[tail_search_start:]
        idx = -1
        for marker in (b"<?xml", b"<!DOCTYPE plist", b"<plist"):
            pos = tail.rfind(marker)
            if pos != -1:
                idx = tail_search_start + pos
                break
        if idx != -1 and idx > payload_start:
            rfc822_bytes = data[payload_start:idx].rstrip() + b"\r\n"
            plist_bytes = data[idx:]
        else:
            rfc822_bytes = data[payload_start : min(len(data), candidate_end)]

    if plist_bytes:
        apple_flags_int: Optional[int] = None
        try:
            pl = plistlib.loads(plist_bytes)
            if isinstance(pl, dict) and "flags" in pl:
                apple_flags_int = int(pl["flags"])
        except Exception:
            m = re.search(rb"<key>\s*flags\s*</key>\s*<integer>\s*(\d+)\s*</integer>", plist_bytes)
            if m:
                apple_flags_int = int(m.group(1))

        if apple_flags_int is not None:
            flags = default_flags(
                is_read=bool(apple_flags_int & (1 << 0)),
                is_deleted=bool(apple_flags_int & (1 << 1)),
                is_replied=bool(apple_flags_int & (1 << 2)),
                is_flagged=bool(apple_flags_int & (1 << 4)),
                is_forwarded=bool(apple_flags_int & (1 << 8)),
            )
            rfc822_bytes = inject_mozilla_status_bytes(rfc822_bytes, flags)
            return rfc822_bytes, flags

    hdr_flags, _ = extract_flags_from_headers(rfc822_bytes)
    return rfc822_bytes, hdr_flags


def normalize_apple_mail_folder_parts(rel_parts: List[str]) -> List[str]:
    """
    Collapse internal Apple Mail V2-V10 bundle directories (such as
    `Messages`, `Data`, numeric shard directories `0`..`9`, `Attachments`,
    and UUID folders inside `.mbox` bundles) into clean folder paths.
    """
    out: List[str] = []
    for part in rel_parts:
        low = part.lower()
        if low in ("messages", "data", "attachments", "contents", "resources"):
            continue
        if part.isdigit() and len(part) <= 2:
            continue
        if APPLE_UUID_RE.match(part):
            continue
        out.append(clean_folder_segment(part))
    return out if out else ["Inbox"]


# ----------------------- 3. Pure-Python Outlook .msg (MS-CFB + MS-OXMSG) -----------------------

class _Ole2Reader:
    """Minimal pure-Python MS-CFB (OLE2 Compound Document) reader for .msg files."""

    ENDOFCHAIN = 0xFFFFFFFE
    FREESECT = 0xFFFFFFFF

    def __init__(self, data: bytes):
        if len(data) < 512 or data[:8] != OLE2_MAGIC:
            raise ValueError("Not a valid OLE2 Compound Document (.msg)")
        self.data = data
        u16 = lambda off: struct.unpack_from("<H", data, off)[0]
        u32 = lambda off: struct.unpack_from("<I", data, off)[0]

        sector_shift = u16(30)
        mini_shift = u16(32)
        self.sector_size = 1 << sector_shift
        self.mini_sector_size = 1 << mini_shift
        self.first_dir_sector = u32(48)
        self.mini_cutoff = u32(56) or 4096
        self.first_minifat_sector = u32(60)
        self.num_minifat_sectors = u32(64)
        self.first_difat_sector = u32(68)
        self.num_difat_sectors = u32(72)

        difat: List[int] = []
        for i in range(109):
            sec = u32(76 + i * 4)
            if sec < self.ENDOFCHAIN:
                difat.append(sec)

        curr_difat = self.first_difat_sector
        entries_per_sector = (self.sector_size // 4) - 1
        for _ in range(self.num_difat_sectors):
            if curr_difat >= self.ENDOFCHAIN:
                break
            off = (curr_difat + 1) * self.sector_size
            if off + self.sector_size > len(data):
                break
            for i in range(entries_per_sector):
                sec = u32(off + i * 4)
                if sec < self.ENDOFCHAIN:
                    difat.append(sec)
            curr_difat = u32(off + entries_per_sector * 4)

        self.fat: List[int] = []
        ints_per_sec = self.sector_size // 4
        for sec in difat:
            off = (sec + 1) * self.sector_size
            if off + self.sector_size <= len(data):
                self.fat.extend(struct.unpack_from(f"<{ints_per_sec}I", data, off))

        dir_bytes = self._read_sector_chain(self.first_dir_sector)
        self.entries: List[Dict] = []
        for pos in range(0, len(dir_bytes), 128):
            chunk = dir_bytes[pos : pos + 128]
            if len(chunk) < 128:
                break
            name_len = struct.unpack_from("<H", chunk, 64)[0]
            obj_type = chunk[66]
            if obj_type == 0 or name_len < 2:
                self.entries.append({"name": "", "type": 0, "left": -1, "right": -1, "child": -1, "start": 0, "size": 0})
                continue
            name = chunk[: max(0, name_len - 2)].decode("utf-16-le", errors="replace")
            left, right, child = struct.unpack_from("<iii", chunk, 68)
            start_sec = struct.unpack_from("<I", chunk, 116)[0]
            size = struct.unpack_from("<Q", chunk, 120)[0]
            if self.sector_size == 512:
                size &= 0xFFFFFFFF
            self.entries.append({
                "name": name,
                "type": obj_type,
                "left": left,
                "right": right,
                "child": child,
                "start": start_sec,
                "size": size,
            })

        self.minifat: List[int] = []
        if self.first_minifat_sector < self.ENDOFCHAIN and self.num_minifat_sectors > 0:
            mf_bytes = self._read_sector_chain(self.first_minifat_sector, self.num_minifat_sectors * self.sector_size)
            self.minifat = list(struct.unpack_from(f"<{len(mf_bytes) // 4}I", mf_bytes, 0))

        if self.entries and self.entries[0]["type"] == 5:
            self.mini_stream = self._read_sector_chain(self.entries[0]["start"], self.entries[0]["size"])
        else:
            self.mini_stream = b""

        self.streams: Dict[str, Dict] = {}
        if self.entries:
            self._walk_tree(self.entries[0]["child"], "")

    def _walk_tree(self, sid: int, prefix: str, visited: Optional[Set[int]] = None):
        if visited is None:
            visited = set()
        if sid < 0 or sid >= len(self.entries) or sid in visited:
            return
        visited.add(sid)
        ent = self.entries[sid]
        self._walk_tree(ent["left"], prefix, visited)
        full_name = f"{prefix}/{ent['name']}" if prefix else ent["name"]
        if ent["type"] == 2:
            self.streams[full_name] = ent
        elif ent["type"] == 1:
            self._walk_tree(ent["child"], full_name, visited)
        self._walk_tree(ent["right"], prefix, visited)

    def _read_sector_chain(self, start_sec: int, max_size: Optional[int] = None) -> bytes:
        out = bytearray()
        curr = start_sec
        seen: Set[int] = set()
        while curr < self.ENDOFCHAIN and curr not in seen:
            seen.add(curr)
            off = (curr + 1) * self.sector_size
            if off + self.sector_size > len(self.data):
                out.extend(self.data[off:])
                break
            out.extend(self.data[off : off + self.sector_size])
            if max_size is not None and len(out) >= max_size:
                return bytes(out[:max_size])
            if curr < len(self.fat):
                curr = self.fat[curr]
            else:
                break
        return bytes(out[:max_size] if max_size is not None else out)

    def read_stream(self, path: str) -> bytes:
        ent = self.streams.get(path)
        if not ent or ent["size"] == 0:
            return b""
        size = ent["size"]
        start = ent["start"]
        if size < self.mini_cutoff and self.mini_stream:
            out = bytearray()
            curr = start
            seen: Set[int] = set()
            while curr < self.ENDOFCHAIN and curr not in seen and len(out) < size:
                seen.add(curr)
                off = curr * self.mini_sector_size
                out.extend(self.mini_stream[off : off + self.mini_sector_size])
                if curr < len(self.minifat):
                    curr = self.minifat[curr]
                else:
                    break
            return bytes(out[:size])
        return self._read_sector_chain(start, size)


def _get_msg_prop_str(ole: _Ole2Reader, prefix: str, prop_id_hex: str) -> str:
    base = f"{prefix}/__substg1.0_{prop_id_hex.upper()}" if prefix else f"__substg1.0_{prop_id_hex.upper()}"
    for p in ole.streams:
        p_up = p.upper()
        if p_up == f"{base.upper()}001F":
            raw = ole.read_stream(p)
            return raw.decode("utf-16-le", errors="replace").rstrip("\x00")
        if p_up == f"{base.upper()}001E":
            raw = ole.read_stream(p)
            try:
                return raw.decode("utf-8").rstrip("\x00")
            except UnicodeDecodeError:
                return raw.decode("cp1252", errors="replace").rstrip("\x00")
    return ""


def _get_msg_prop_bin(ole: _Ole2Reader, prefix: str, prop_id_hex: str) -> bytes:
    target = f"{prefix}/__substg1.0_{prop_id_hex.upper()}0102" if prefix else f"__substg1.0_{prop_id_hex.upper()}0102"
    for p in ole.streams:
        if p.upper() == target.upper():
            return ole.read_stream(p)
    return b""


def _parse_msg_fixed_props(ole: _Ole2Reader, prefix: str = "") -> Dict[int, int]:
    stream_name = f"{prefix}/__properties_version1.0" if prefix else "__properties_version1.0"
    raw = b""
    for p in ole.streams:
        if p.upper() == stream_name.upper():
            raw = ole.read_stream(p)
            break
    if not raw:
        return {}
    hdr_size = 32 if not prefix else 24
    if len(raw) < hdr_size:
        return {}
    props: Dict[int, int] = {}
    for off in range(hdr_size, len(raw) - 15, 16):
        prop_tag, _flags, val = struct.unpack_from("<IIQ", raw, off)
        props[prop_tag] = val
    return props


def parse_msg_bytes(data: bytes) -> Tuple[bytes, Dict[str, bool]]:
    """
    Parse an Outlook .msg (OLE2 MS-OXMSG) binary buffer into an RFC 5322 EmailMessage
    byte string and normalized status flags.
    """
    ole = _Ole2Reader(data)
    fixed_props = _parse_msg_fixed_props(ole, "")

    msg_flags = fixed_props.get(0x0E070003, 0x01) & 0xFFFFFFFF
    flag_status = fixed_props.get(0x10900003, 0) & 0xFFFFFFFF
    last_verb = fixed_props.get(0x10810003, 0) & 0xFFFFFFFF

    flags = default_flags(
        is_read=bool(msg_flags & 0x01),
        is_replied=(last_verb in (102, 103)),
        is_flagged=(flag_status == 2),
        is_forwarded=(last_verb == 104),
        is_deleted=False,
    )

    transport_headers = _get_msg_prop_str(ole, "", "007D")
    subject = _get_msg_prop_str(ole, "", "0037")
    body_text = _get_msg_prop_str(ole, "", "1000")
    body_html = _get_msg_prop_str(ole, "", "1013")
    if not body_html:
        html_bin = _get_msg_prop_bin(ole, "", "1013")
        if html_bin:
            try:
                body_html = html_bin.decode("utf-8")
            except UnicodeDecodeError:
                body_html = html_bin.decode("cp1252", errors="replace")

    sender_name = _get_msg_prop_str(ole, "", "0C1A") or _get_msg_prop_str(ole, "", "0042")
    sender_email = (
        _get_msg_prop_str(ole, "", "5D01")
        or _get_msg_prop_str(ole, "", "0C1F")
        or _get_msg_prop_str(ole, "", "5D02")
        or _get_msg_prop_str(ole, "", "0065")
    )
    msg_id = _get_msg_prop_str(ole, "", "1035")

    recip_storages = sorted({
        p.split("/")[0]
        for p in ole.streams
        if p.upper().startswith("__RECIP_VERSION1.0_#")
    })
    to_list: List[str] = []
    cc_list: List[str] = []
    bcc_list: List[str] = []
    for r_prefix in recip_storages:
        r_props = _parse_msg_fixed_props(ole, r_prefix)
        r_type = r_props.get(0x0C150003, 1) & 0xFFFFFFFF
        r_name = _get_msg_prop_str(ole, r_prefix, "3001")
        r_email = _get_msg_prop_str(ole, r_prefix, "39FE") or _get_msg_prop_str(ole, r_prefix, "3003")
        formatted = email.utils.formataddr((r_name, r_email)) if r_email else r_name
        if not formatted:
            continue
        if r_type == 2:
            cc_list.append(formatted)
        elif r_type == 3:
            bcc_list.append(formatted)
        else:
            to_list.append(formatted)

    if not to_list:
        disp_to = _get_msg_prop_str(ole, "", "0E04")
        if disp_to:
            to_list = [disp_to]
    if not cc_list:
        disp_cc = _get_msg_prop_str(ole, "", "0E03")
        if disp_cc:
            cc_list = [disp_cc]
    if not bcc_list:
        disp_bcc = _get_msg_prop_str(ole, "", "0E02")
        if disp_bcc:
            bcc_list = [disp_bcc]

    date_str = ""
    ft = fixed_props.get(0x0E060040) or fixed_props.get(0x00390040)
    if ft and ft > 0:
        try:
            dt = datetime(1601, 1, 1, tzinfo=timezone.utc) + timedelta(microseconds=ft // 10)
            if 1970 <= dt.year <= 2100:
                date_str = email.utils.format_datetime(dt)
        except Exception:
            pass

    if not date_str and transport_headers:
        m_date = re.search(r"(?im)^Date\s*:\s*([^\r\n]+)", transport_headers)
        if m_date:
            date_str = m_date.group(1).strip()

    msg = EmailMessage()
    if subject:
        msg["Subject"] = subject
    if sender_email or sender_name:
        msg["From"] = email.utils.formataddr((sender_name, sender_email)) if "@" in sender_email else (sender_name or sender_email)
    elif transport_headers:
        m_from = re.search(r"(?im)^From\s*:\s*([^\r\n]+)", transport_headers)
        if m_from:
            msg["From"] = m_from.group(1).strip()

    if to_list:
        msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    if bcc_list:
        msg["Bcc"] = ", ".join(bcc_list)
    if date_str:
        msg["Date"] = date_str
    if msg_id:
        msg["Message-ID"] = msg_id
    msg["X-Mozilla-Status"] = flags_to_mozilla_status_hex(flags)

    if body_text and body_html:
        msg.set_content(body_text)
        msg.add_alternative(body_html, subtype="html")
    elif body_html:
        msg.set_content(body_html, subtype="html")
    else:
        msg.set_content(body_text or "")

    attach_storages = sorted({
        p.split("/")[0]
        for p in ole.streams
        if p.upper().startswith("__ATTACH_VERSION1.0_#")
    })
    for a_prefix in attach_storages:
        att_data = _get_msg_prop_bin(ole, a_prefix, "3701")
        if not att_data:
            continue
        att_name = (
            _get_msg_prop_str(ole, a_prefix, "3707")
            or _get_msg_prop_str(ole, a_prefix, "3704")
            or "attachment.bin"
        )
        att_mime = _get_msg_prop_str(ole, a_prefix, "370E")
        if not att_mime or "/" not in att_mime:
            guessed, _ = mimetypes.guess_type(att_name)
            att_mime = guessed or "application/octet-stream"
        maintype, subtype = att_mime.split("/", 1)
        att_cid = _get_msg_prop_str(ole, a_prefix, "3712").strip("<>")
        if att_cid:
            msg.add_attachment(
                att_data,
                maintype=maintype,
                subtype=subtype,
                filename=att_name,
                cid=f"<{att_cid}>",
                disposition="inline",
            )
        else:
            msg.add_attachment(
                att_data,
                maintype=maintype,
                subtype=subtype,
                filename=att_name,
            )

    return msg.as_bytes(policy=policy.SMTP), flags


# ----------------------- 4. Outlook for Mac .olm Archives -----------------------

def _parse_olm_addresses(elem: Optional[ET.Element]) -> List[str]:
    if elem is None:
        return []
    addrs: List[str] = []
    for addr_el in elem.findall(".//emailAddress"):
        name = (addr_el.attrib.get("OPFContactEmailAddressName") or "").strip()
        email_addr = (addr_el.attrib.get("OPFContactEmailAddressAddress") or "").strip()
        if email_addr:
            addrs.append(email.utils.formataddr((name, email_addr)))
        elif name:
            addrs.append(name)
    return addrs


def _parse_olm_iso_date(val: str) -> str:
    val = (val or "").strip()
    if not val:
        return ""
    try:
        dt = datetime.fromisoformat(val.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return email.utils.format_datetime(dt)
    except Exception:
        return val


def iter_olm_records(
    zf: zipfile.ZipFile,
    base_folder_parts: Optional[List[str]] = None,
    include_deleted: bool = False,
) -> Iterator[MailRecord]:
    """
    Iterate over all email messages inside an Outlook for Mac (.olm) ZIP archive,
    reconstructing RFC 5322 EmailMessages with HTML/plain bodies, attachments,
    folder hierarchy, and Read/Flagged status.
    """
    base_parts = list(base_folder_parts) if base_folder_parts else []
    namelist = zf.namelist()
    name_lookup = {n: n for n in namelist}
    name_lookup_lower = {n.lower(): n for n in namelist}

    def _read_olm_attachment(url: str) -> Optional[bytes]:
        if not url:
            return None
        candidates = [
            url,
            url.lstrip("/"),
            urllib.parse.unquote(url),
            urllib.parse.unquote(url).lstrip("/"),
        ]
        for cand in candidates:
            if cand in name_lookup:
                return zf.read(name_lookup[cand])
            if cand.lower() in name_lookup_lower:
                return zf.read(name_lookup_lower[cand.lower()])
        return None

    for member in namelist:
        if member.endswith("/") or not member.lower().endswith(".xml"):
            continue
        base_fname = os.path.basename(member).lower()
        if base_fname in ("categories.xml", "accounts.xml", "account.xml"):
            continue
        if "com.microsoft.__messages" not in member.lower() and not base_fname.startswith("message_"):
            continue

        raw_parts = [p for p in member.replace("\\", "/").split("/") if p]
        msg_marker_idx = -1
        for idx, p in enumerate(raw_parts):
            if p.lower() == "com.microsoft.__messages":
                msg_marker_idx = idx
                break
        if msg_marker_idx != -1:
            sub_folders = [clean_folder_segment(p) for p in raw_parts[msg_marker_idx + 1 : -1]]
        else:
            sub_folders = [clean_folder_segment(p) for p in raw_parts[:-1]]
        folder_parts = base_parts + (sub_folders if sub_folders else ["Inbox"])

        if not include_deleted and any(
            p.lower() in ("trash", "deleted items", "deleted messages") for p in sub_folders
        ):
            continue

        try:
            xml_bytes = zf.read(member)
            root = ET.fromstring(xml_bytes)
        except Exception:
            continue

        email_elems = root.findall(".//email") if root.tag != "email" else [root]
        for idx_em, em_el in enumerate(email_elems):
            def _txt(tag: str) -> str:
                node = em_el.find(tag)
                return (node.text or "").strip() if node is not None else ""

            subject = _txt("OPFMessageCopySubject")
            body_text = _txt("OPFMessageCopyBody")
            body_html = _txt("OPFMessageCopyHTMLBody")
            msg_id = _txt("OPFMessageCopyMessageID")
            date_str = _parse_olm_iso_date(
                _txt("OPFMessageCopySentTime") or _txt("OPFMessageCopyReceivedTime")
            )

            is_read_str = _txt("OPFMessageGetIsRead")
            flag_status_str = _txt("OPFMessageCopyFlagStatus")
            flags = default_flags(
                is_read=(is_read_str != "0"),
                is_replied=False,
                is_flagged=(flag_status_str in ("1", "2")),
                is_forwarded=False,
                is_deleted=False,
            )

            from_addrs = _parse_olm_addresses(em_el.find("OPFMessageCopyFromAddresses"))
            if not from_addrs:
                from_addrs = _parse_olm_addresses(em_el.find("OPFMessageCopySenderAddress"))
            to_addrs = _parse_olm_addresses(em_el.find("OPFMessageCopyToAddresses"))
            cc_addrs = _parse_olm_addresses(em_el.find("OPFMessageCopyCCAddresses"))
            bcc_addrs = _parse_olm_addresses(em_el.find("OPFMessageCopyBCCAddresses"))

            msg = EmailMessage()
            if subject:
                msg["Subject"] = subject
            if from_addrs:
                msg["From"] = ", ".join(from_addrs)
            if to_addrs:
                msg["To"] = ", ".join(to_addrs)
            if cc_addrs:
                msg["Cc"] = ", ".join(cc_addrs)
            if bcc_addrs:
                msg["Bcc"] = ", ".join(bcc_addrs)
            if date_str:
                msg["Date"] = date_str
            if msg_id:
                msg["Message-ID"] = msg_id if msg_id.startswith("<") else f"<{msg_id}>"
            msg["X-Mozilla-Status"] = flags_to_mozilla_status_hex(flags)

            if body_text and body_html:
                msg.set_content(body_text)
                msg.add_alternative(body_html, subtype="html")
            elif body_html:
                msg.set_content(body_html, subtype="html")
            else:
                msg.set_content(body_text or "")

            att_list_el = em_el.find("OPFMessageCopyAttachmentList")
            if att_list_el is not None:
                for att_el in att_list_el.findall(".//attachment"):
                    att_url = att_el.attrib.get("OPFAttachmentURL") or ""
                    att_data = _read_olm_attachment(att_url)
                    if att_data is None:
                        continue
                    att_name = (
                        att_el.attrib.get("OPFAttachmentName")
                        or os.path.basename(att_url)
                        or "attachment.bin"
                    )
                    att_ctype = att_el.attrib.get("OPFAttachmentContentType") or ""
                    if "/" not in att_ctype:
                        guessed, _ = mimetypes.guess_type(att_name)
                        att_ctype = guessed or "application/octet-stream"
                    maintype, subtype = att_ctype.split("/", 1)
                    att_cid = (att_el.attrib.get("OPFAttachmentContentID") or "").strip("<>")
                    if att_cid:
                        msg.add_attachment(
                            att_data,
                            maintype=maintype,
                            subtype=subtype,
                            filename=att_name,
                            cid=f"<{att_cid}>",
                            disposition="inline",
                        )
                    else:
                        msg.add_attachment(
                            att_data,
                            maintype=maintype,
                            subtype=subtype,
                            filename=att_name,
                        )

            src_name = os.path.splitext(os.path.basename(member))[0]
            if len(email_elems) > 1:
                src_name = f"{src_name}_{idx_em + 1}"
            yield MailRecord(
                raw_bytes=msg.as_bytes(policy=policy.SMTP),
                folder_parts=folder_parts,
                source_name=f"{src_name}.eml",
                is_deleted=False,
                flags=flags,
            )


# ----------------------- 5. Outlook .pst Reader (via MAPI on Windows) -----------------------

def iter_pst_records(
    pst_path: str,
    base_folder_parts: Optional[List[str]] = None,
    include_deleted: bool = False,
    namespace=None,
) -> Iterator[MailRecord]:
    """
    Read and yield MailRecords from a Microsoft Outlook .pst file via Outlook MAPI.
    Preserves folder hierarchy, HTML/plain bodies, attachments, inline CIDs, and status flags.
    """
    abs_pst = os.path.abspath(pst_path)
    base_parts = list(base_folder_parts) if base_folder_parts else []

    own_com = False
    if namespace is None:
        try:
            import pythoncom
            import win32com.client as win32
            pythoncom.CoInitialize()
            own_com = True
            app = win32.Dispatch("Outlook.Application")
            namespace = app.GetNamespace("MAPI")
            namespace.Logon("", "", False, True)
        except Exception as exc:
            raise RuntimeError(
                f"Reading .pst files requires Windows with Classic Outlook and pywin32: {exc}"
            ) from exc

    norm_target = os.path.normcase(os.path.normpath(abs_pst))
    target_store = None
    already_mounted = False
    existing_ids = set()

    for s in namespace.Stores:
        try:
            existing_ids.add(s.StoreID)
            if s.FilePath and os.path.normcase(os.path.normpath(s.FilePath)) == norm_target:
                target_store = s
                already_mounted = True
                break
        except Exception:
            pass

    if target_store is None:
        try:
            namespace.AddStoreEx(abs_pst, 2)
        except Exception:
            namespace.AddStore(abs_pst)
        for s in namespace.Stores:
            try:
                if s.FilePath and os.path.normcase(os.path.normpath(s.FilePath)) == norm_target:
                    target_store = s
                    break
                if s.StoreID not in existing_ids:
                    target_store = s
                    break
            except Exception:
                pass

    if target_store is None:
        return

    root_folder = target_store.GetRootFolder()
    tmpdir = tempfile.mkdtemp(prefix="pst_read_")

    def _walk_pst_folder(folder, rel_parts: List[str]) -> Iterator[MailRecord]:
        # Skip Deleted Items / Trash unless include_deleted=True
        if rel_parts and not include_deleted:
            if rel_parts[-1].lower() in ("deleted items", "trash", "deleted messages"):
                return

        try:
            items = folder.Items
            count = items.Count
        except Exception:
            count = 0

        for idx in range(1, count + 1):
            try:
                item = items.Item(idx)
                if getattr(item, "Class", None) != 43:  # 43 = olMail
                    continue

                is_read = not bool(getattr(item, "UnRead", False))
                is_flagged = getattr(item, "FlagStatus", 0) == 2
                is_replied = False
                is_forwarded = False
                pa = getattr(item, "PropertyAccessor", None)
                if pa is not None:
                    try:
                        verb = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
                        is_replied = verb in (102, 103)
                        is_forwarded = verb == 104
                    except Exception:
                        pass

                flags = default_flags(
                    is_read=is_read,
                    is_replied=is_replied,
                    is_flagged=is_flagged,
                    is_forwarded=is_forwarded,
                    is_deleted=False,
                )

                msg = EmailMessage()
                subj = getattr(item, "Subject", "") or ""
                if subj:
                    msg["Subject"] = subj

                sender_name = getattr(item, "SenderName", "") or ""
                sender_email = getattr(item, "SenderEmailAddress", "") or ""
                if sender_email or sender_name:
                    msg["From"] = (
                        email.utils.formataddr((sender_name, sender_email))
                        if "@" in sender_email
                        else (sender_name or sender_email)
                    )

                to_val = getattr(item, "To", "") or ""
                if to_val:
                    msg["To"] = to_val.replace(";", ",")
                cc_val = getattr(item, "CC", "") or ""
                if cc_val:
                    msg["Cc"] = cc_val.replace(";", ",")
                bcc_val = getattr(item, "BCC", "") or ""
                if bcc_val:
                    msg["Bcc"] = bcc_val.replace(";", ",")

                dt = getattr(item, "SentOn", None) or getattr(item, "ReceivedTime", None)
                if isinstance(dt, datetime):
                    if dt.tzinfo is None:
                        dt = dt.replace(tzinfo=timezone.utc)
                    msg["Date"] = email.utils.format_datetime(dt)

                if pa is not None:
                    try:
                        mid = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
                        if mid:
                            msg["Message-ID"] = str(mid)
                    except Exception:
                        pass

                msg["X-Mozilla-Status"] = flags_to_mozilla_status_hex(flags)

                body_text = getattr(item, "Body", "") or ""
                body_html = getattr(item, "HTMLBody", "") or ""
                if body_text and body_html:
                    msg.set_content(body_text)
                    msg.add_alternative(body_html, subtype="html")
                elif body_html:
                    msg.set_content(body_html, subtype="html")
                else:
                    msg.set_content(body_text)

                atts = getattr(item, "Attachments", None)
                att_count = getattr(atts, "Count", 0) if atts is not None else 0
                for a_idx in range(1, att_count + 1):
                    try:
                        att = atts.Item(a_idx)
                        att_fn = sanitize_name(getattr(att, "FileName", "") or f"att_{a_idx}.bin", max_len=80)
                        tmp_att = os.path.join(tmpdir, f"{idx}_{a_idx}_{att_fn}")
                        att.SaveAsFile(tmp_att)
                        with open(tmp_att, "rb") as af:
                            att_data = af.read()
                        os.remove(tmp_att)
                        guessed, _ = mimetypes.guess_type(att_fn)
                        maintype, subtype = (guessed or "application/octet-stream").split("/", 1)
                        cid = ""
                        att_pa = getattr(att, "PropertyAccessor", None)
                        if att_pa is not None:
                            try:
                                cid = str(att_pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F") or "").strip("<>")
                            except Exception:
                                cid = ""
                        if cid:
                            msg.add_attachment(att_data, maintype=maintype, subtype=subtype, filename=att_fn, cid=f"<{cid}>", disposition="inline")
                        else:
                            msg.add_attachment(att_data, maintype=maintype, subtype=subtype, filename=att_fn)
                    except Exception:
                        pass

                yield MailRecord(
                    raw_bytes=msg.as_bytes(policy=policy.SMTP),
                    folder_parts=base_parts + (rel_parts if rel_parts else ["Inbox"]),
                    source_name=f"pst_msg_{idx:06d}.eml",
                    is_deleted=False,
                    flags=flags,
                )
            except Exception:
                continue

        try:
            subfolders = folder.Folders
            sub_count = subfolders.Count
        except Exception:
            sub_count = 0

        for s_idx in range(1, sub_count + 1):
            try:
                sub = subfolders.Item(s_idx)
                sub_name = clean_folder_segment(getattr(sub, "Name", "") or f"Folder_{s_idx}")
                yield from _walk_pst_folder(sub, rel_parts + [sub_name])
            except Exception:
                continue

    try:
        yield from _walk_pst_folder(root_folder, [])
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
        if not already_mounted:
            try:
                namespace.RemoveStore(root_folder)
            except Exception:
                pass
        if own_com:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ----------------------- 6. MH / Claws / Sylpheed Mailboxes -----------------------

def parse_mh_sequences(seq_file_path: str) -> Dict[str, Set[int]]:
    """
    Parse an MH `.mh_sequences` file into a dict mapping sequence names
    (e.g. 'unseen', 'flagged', 'replied') to sets of message numbers.
    """
    sequences: Dict[str, Set[int]] = {}
    try:
        with open(win_long_path(seq_file_path), "r", encoding="utf-8", errors="replace") as f:
            for line in f:
                line = line.strip()
                if not line or ":" not in line:
                    continue
                name, rest = line.split(":", 1)
                name = name.strip().lower()
                nums = sequences.setdefault(name, set())
                for token in rest.split():
                    if "-" in token:
                        parts = token.split("-", 1)
                        if parts[0].isdigit() and parts[1].isdigit():
                            start, end = int(parts[0]), int(parts[1])
                            if start <= end and (end - start) <= 1_000_000:
                                nums.update(range(start, end + 1))
                    elif token.isdigit():
                        nums.add(int(token))
    except OSError:
        pass
    return sequences


def flags_for_mh_message(msg_num: int, sequences: Dict[str, Set[int]]) -> Tuple[Dict[str, bool], bool]:
    """Compute message flags for an MH message number given parsed .mh_sequences."""
    if not sequences:
        return default_flags(is_read=True), False
    is_read = True
    if "unseen" in sequences:
        is_read = msg_num not in sequences["unseen"]
    is_flagged = (
        msg_num in sequences.get("flagged", set())
        or msg_num in sequences.get("tick", set())
        or msg_num in sequences.get("mark", set())
    )
    is_replied = msg_num in sequences.get("replied", set()) or msg_num in sequences.get("answered", set())
    is_forwarded = msg_num in sequences.get("forwarded", set())
    is_deleted = msg_num in sequences.get("deleted", set())
    return (
        default_flags(
            is_read=is_read,
            is_replied=is_replied,
            is_flagged=is_flagged,
            is_forwarded=is_forwarded,
            is_deleted=is_deleted,
        ),
        True,
    )


# ----------------------- 7. MBOX, MMDF, and Babyl Stream Readers -----------------------

def is_mbox_from_line(line: bytes) -> bool:
    if not line.startswith(b"From "):
        return False
    if MBOX_FROM_RE.match(line):
        return True
    parts = line.split()
    return len(parts) >= 3 and len(parts[-1]) == 4 and parts[-1].isdigit()


def _finalize_stream_message(
    lines: List[bytes],
    folder_parts: List[str],
    source_name: str,
    ext_flags: Optional[Dict[str, bool]] = None,
) -> MailRecord:
    while lines and lines[-1] in (b"\n", b"\r\n"):
        lines.pop()
    raw = b"".join(lines)
    hdr_flags, has_hdr = extract_flags_from_headers(raw)
    if ext_flags is not None and not has_hdr:
        flags = ext_flags
        raw = inject_mozilla_status_bytes(raw, flags)
    else:
        flags = hdr_flags
    return MailRecord(
        raw_bytes=raw,
        folder_parts=list(folder_parts),
        source_name=source_name,
        is_deleted=flags.get("is_deleted", False),
        flags=flags,
    )


def iter_mbox_stream(
    fp: BinaryIO,
    folder_parts: List[str],
    source_prefix: str = "msg",
    include_deleted: bool = False,
) -> Iterator[MailRecord]:
    """
    Stream messages from an mbox / mboxrd file-like object with O(1) memory per message.
    Un-escapes mboxrd '>+From ' lines in message bodies.
    """
    msg_lines: List[bytes] = []
    in_msg = False
    in_headers = False
    idx = 0

    for line in fp:
        if not in_msg:
            if is_mbox_from_line(line):
                in_msg = True
                in_headers = True
                msg_lines = []
            elif HEADER_START_RE.match(line):
                in_msg = True
                in_headers = not (line in (b"\n", b"\r\n"))
                msg_lines = [line]
        else:
            if in_headers:
                if is_mbox_from_line(line):
                    if msg_lines:
                        idx += 1
                        rec = _finalize_stream_message(msg_lines, folder_parts, f"{source_prefix}_{idx:06d}.eml")
                        if include_deleted or not rec.is_deleted:
                            yield rec
                    in_headers = True
                    msg_lines = []
                    continue
                if line in (b"\n", b"\r\n"):
                    in_headers = False
                msg_lines.append(line)
            else:
                if is_mbox_from_line(line):
                    if msg_lines:
                        idx += 1
                        rec = _finalize_stream_message(msg_lines, folder_parts, f"{source_prefix}_{idx:06d}.eml")
                        if include_deleted or not rec.is_deleted:
                            yield rec
                    in_headers = True
                    msg_lines = []
                else:
                    stripped = line.lstrip(b">")
                    if stripped.startswith(b"From ") and line.startswith(b">"):
                        line = line[1:]
                    msg_lines.append(line)

    if in_msg and msg_lines:
        idx += 1
        rec = _finalize_stream_message(msg_lines, folder_parts, f"{source_prefix}_{idx:06d}.eml")
        if include_deleted or not rec.is_deleted:
            yield rec


def iter_mmdf_stream(
    fp: BinaryIO,
    folder_parts: List[str],
    source_prefix: str = "mmdf",
    include_deleted: bool = False,
) -> Iterator[MailRecord]:
    """Stream messages from an MMDF mailbox file (delimited by \\x01\\x01\\x01\\x01 lines)."""
    msg_lines: List[bytes] = []
    idx = 0
    for line in fp:
        if line.rstrip(b"\r\n") == b"\x01\x01\x01\x01":
            if msg_lines:
                if is_mbox_from_line(msg_lines[0]):
                    msg_lines = msg_lines[1:]
                if msg_lines:
                    idx += 1
                    rec = _finalize_stream_message(msg_lines, folder_parts, f"{source_prefix}_{idx:06d}.eml")
                    if include_deleted or not rec.is_deleted:
                        yield rec
                msg_lines = []
        else:
            msg_lines.append(line)

    if msg_lines:
        if is_mbox_from_line(msg_lines[0]):
            msg_lines = msg_lines[1:]
        if msg_lines:
            idx += 1
            rec = _finalize_stream_message(msg_lines, folder_parts, f"{source_prefix}_{idx:06d}.eml")
            if include_deleted or not rec.is_deleted:
                yield rec


def iter_babyl_stream(
    fp: BinaryIO,
    folder_parts: List[str],
    source_prefix: str = "babyl",
    include_deleted: bool = False,
) -> Iterator[MailRecord]:
    """Parse messages from an Emacs Rmail Babyl mailbox (BABYL OPTIONS: ... \\x1f\\x0c ...)."""
    data = fp.read()
    blocks = data.split(b"\x1f")
    idx = 0
    for block in blocks[1:]:
        block = block.lstrip(b"\x0c\r\n")
        if not block.strip():
            continue
        lines = block.splitlines(keepends=True)
        if not lines:
            continue

        attr_line = lines[0].decode("latin1", errors="replace").lower()
        flags = default_flags(
            is_read=("unseen" not in attr_line),
            is_replied=("answered" in attr_line),
            is_flagged=("flagged" in attr_line),
            is_forwarded=("forwarded" in attr_line),
            is_deleted=("deleted" in attr_line),
        )
        if flags["is_deleted"] and not include_deleted:
            continue

        eooh_idx = -1
        for i, ln in enumerate(lines[1:], start=1):
            if ln.strip() == b"*** EOOH ***":
                eooh_idx = i
                break

        if eooh_idx != -1:
            before_eooh = lines[1:eooh_idx]
            after_eooh = lines[eooh_idx + 1 :]
            if after_eooh and HEADER_START_RE.match(after_eooh[0]):
                msg_lines = after_eooh
            else:
                msg_lines = before_eooh + after_eooh
        else:
            msg_lines = lines[1:]

        if msg_lines:
            idx += 1
            yield _finalize_stream_message(msg_lines, folder_parts, f"{source_prefix}_{idx:06d}.eml", ext_flags=flags)


# ----------------------- 8. Format Detection & Unified Dispatcher -----------------------

def _sniff_stream_format(header_bytes: bytes) -> str:
    """Sniff the format of a stream from its initial bytes."""
    if header_bytes.startswith(PST_MAGIC):
        return "pst"
    if header_bytes.startswith(OLE2_MAGIC):
        return "msg"
    if header_bytes.startswith(b"PK\x03\x04"):
        return "zip"
    if header_bytes.startswith(b"\x1f\x8b"):
        return "gz"
    if header_bytes.startswith(b"BZh"):
        return "bz2"
    if header_bytes.startswith(b"\xfd7zXZ\x00"):
        return "xz"
    if header_bytes.startswith(b"\x01\x01\x01\x01"):
        return "mmdf"
    if header_bytes.lstrip().startswith(b"BABYL OPTIONS:"):
        return "babyl"
    first_nl = header_bytes.find(b"\n")
    if first_nl != -1 and header_bytes[:first_nl].strip().isdigit():
        return "emlx"
    if is_mbox_from_line(header_bytes[:first_nl + 1] if first_nl != -1 else header_bytes):
        return "mbox"
    if re.search(rb"(?m)^From \S+ .*\d{4}\r?$", header_bytes):
        return "mbox"
    if HEADER_START_RE.match(header_bytes.lstrip()):
        return "eml"
    return "unknown"


def _iter_records_from_bytes(
    data: bytes,
    folder_parts: List[str],
    filename: str,
    include_deleted: bool = False,
    ext_flags: Optional[Dict[str, bool]] = None,
    native_msg_path: Optional[str] = None,
) -> Iterator[MailRecord]:
    low = filename.lower()
    fmt = _sniff_stream_format(data[:4096])

    if low.endswith((".emlx", ".partial.emlx")) or fmt == "emlx":
        raw, flags = parse_emlx_bytes(data)
        if raw and (include_deleted or not flags.get("is_deleted", False)):
            yield MailRecord(
                raw_bytes=raw,
                folder_parts=list(folder_parts),
                source_name=filename,
                is_deleted=flags.get("is_deleted", False),
                flags=flags,
            )
        return

    if low.endswith(".msg") or fmt == "msg":
        try:
            raw, flags = parse_msg_bytes(data)
            if raw and (include_deleted or not flags.get("is_deleted", False)):
                yield MailRecord(
                    raw_bytes=raw,
                    folder_parts=list(folder_parts),
                    source_name=filename,
                    is_deleted=flags.get("is_deleted", False),
                    flags=flags,
                    native_msg_path=native_msg_path,
                )
        except Exception:
            pass
        return

    if fmt == "mmdf":
        yield from iter_mmdf_stream(io.BytesIO(data), folder_parts, os.path.splitext(filename)[0], include_deleted)
        return

    if fmt == "babyl":
        yield from iter_babyl_stream(io.BytesIO(data), folder_parts, os.path.splitext(filename)[0], include_deleted)
        return

    if low.endswith(".mbox") or fmt == "mbox":
        yield from iter_mbox_stream(io.BytesIO(data), folder_parts, os.path.splitext(filename)[0], include_deleted)
        return

    hdr_flags, has_hdr = extract_flags_from_headers(data)
    if ext_flags is not None and not has_hdr:
        flags = ext_flags
        data = inject_mozilla_status_bytes(data, flags)
    else:
        flags = hdr_flags
    if include_deleted or not flags.get("is_deleted", False):
        yield MailRecord(
            raw_bytes=data,
            folder_parts=list(folder_parts),
            source_name=filename,
            is_deleted=flags.get("is_deleted", False),
            flags=flags,
        )


def _derive_archive_member_folder_parts(
    member_path: str,
    base_folder_parts: List[str],
) -> Tuple[List[str], Optional[Dict[str, bool]]]:
    norm = member_path.replace("\\", "/").strip("/")
    parts = [p for p in norm.split("/") if p]
    if not parts:
        return list(base_folder_parts) or ["Inbox"], None

    fname = parts[-1]
    dir_parts = parts[:-1]

    if dir_parts and dir_parts[-1].lower() in ("cur", "new"):
        in_new = dir_parts[-1].lower() == "new"
        md_flags, has_md_flags = parse_maildir_flags(fname, in_new=in_new)
        folder_parts = list(base_folder_parts) + normalize_maildir_folder_parts(dir_parts)
        return folder_parts, (md_flags if has_md_flags else None)

    if fname.lower().endswith((".emlx", ".partial.emlx")) or any(
        p.lower().endswith((".mbox", ".imapmbox", ".popmbox")) for p in dir_parts
    ):
        folder_parts = list(base_folder_parts) + normalize_apple_mail_folder_parts(dir_parts)
        return folder_parts, None

    cleaned = [clean_folder_segment(p) for p in dir_parts]
    folder_parts = list(base_folder_parts) + (cleaned if cleaned else ["Inbox"])
    return folder_parts, None


def iter_zip_or_olm_records(
    zip_path_or_fileobj,
    base_folder_parts: Optional[List[str]] = None,
    include_deleted: bool = False,
) -> Iterator[MailRecord]:
    """Iterate records from a .zip or .olm archive."""
    base_parts = list(base_folder_parts) if base_folder_parts else []
    with zipfile.ZipFile(zip_path_or_fileobj, "r") as zf:
        names = zf.namelist()
        is_olm = any("com.microsoft.__messages" in n.lower() for n in names)
        if is_olm:
            yield from iter_olm_records(zf, base_parts, include_deleted=include_deleted)
            return

        for member in names:
            if member.endswith("/"):
                continue
            norm = member.replace("\\", "/")
            parts = [p for p in norm.split("/") if p]
            if not parts:
                continue
            if any(p == "__MACOSX" or p.startswith("._") or p.lower() == "tmp" for p in parts):
                continue
            fname = parts[-1]
            if fname.lower() in THUNDERBIRD_SKIP_FILES or fname.lower().endswith(THUNDERBIRD_SKIP_SUFFIXES):
                continue

            folder_parts, ext_flags = _derive_archive_member_folder_parts(norm, base_parts)
            if ext_flags and ext_flags.get("is_deleted") and not include_deleted:
                continue

            try:
                data = zf.read(member)
            except Exception:
                continue
            if not data:
                continue

            yield from _iter_records_from_bytes(
                data,
                folder_parts,
                fname,
                include_deleted=include_deleted,
                ext_flags=ext_flags,
            )


def iter_tar_records(
    tar_path_or_fileobj,
    base_folder_parts: Optional[List[str]] = None,
    include_deleted: bool = False,
) -> Iterator[MailRecord]:
    """Iterate records from a .tar / .tar.gz / .tgz / .tar.bz2 / .tar.xz archive."""
    base_parts = list(base_folder_parts) if base_folder_parts else []
    if isinstance(tar_path_or_fileobj, str):
        tf = tarfile.open(win_long_path(tar_path_or_fileobj), "r:*")
    else:
        tf = tarfile.open(fileobj=tar_path_or_fileobj, mode="r:*")

    with tf:
        for member in tf:
            if not member.isfile() or member.size == 0:
                continue
            norm = member.name.replace("\\", "/").lstrip("./")
            parts = [p for p in norm.split("/") if p]
            if not parts:
                continue
            if any(p == "__MACOSX" or p.startswith("._") or p.lower() == "tmp" for p in parts):
                continue
            fname = parts[-1]
            if fname.lower() in THUNDERBIRD_SKIP_FILES or fname.lower().endswith(THUNDERBIRD_SKIP_SUFFIXES):
                continue

            folder_parts, ext_flags = _derive_archive_member_folder_parts(norm, base_parts)
            if ext_flags and ext_flags.get("is_deleted") and not include_deleted:
                continue

            fobj = tf.extractfile(member)
            if fobj is None:
                continue
            with fobj:
                data = fobj.read()
            if not data:
                continue

            yield from _iter_records_from_bytes(
                data,
                folder_parts,
                fname,
                include_deleted=include_deleted,
                ext_flags=ext_flags,
            )


def iter_file_records(
    file_path: str,
    folder_parts: Optional[List[str]] = None,
    include_deleted: bool = False,
    ext_flags: Optional[Dict[str, bool]] = None,
) -> Iterator[MailRecord]:
    """
    Yield MailRecords from a single file on disk (supports .mbox, extensionless mbox,
    .eml, .emlx, .msg, .olm, .pst, MMDF, Babyl, .zip, .tar*, .gz, .bz2, .xz).
    """
    lp = win_long_path(file_path)
    try:
        if not os.path.isfile(lp) or os.path.getsize(lp) == 0:
            return
    except OSError:
        return

    fname = os.path.basename(file_path)
    low = fname.lower()
    if low in THUNDERBIRD_SKIP_FILES or low.endswith(THUNDERBIRD_SKIP_SUFFIXES):
        return

    default_folder = list(folder_parts) if folder_parts is not None else [clean_folder_segment(os.path.splitext(fname)[0] or "Inbox")]

    if low.endswith(".pst"):
        yield from iter_pst_records(file_path, folder_parts or [], include_deleted=include_deleted)
        return

    if low.endswith((".tar", ".tar.gz", ".tgz", ".tar.bz2", ".tbz2", ".tar.xz", ".txz")):
        yield from iter_tar_records(file_path, folder_parts or [], include_deleted=include_deleted)
        return

    if low.endswith((".zip", ".olm")):
        yield from iter_zip_or_olm_records(lp, folder_parts or [], include_deleted=include_deleted)
        return

    if low.endswith((".gz", ".bz2", ".xz")):
        open_fn = gzip.open if low.endswith(".gz") else (bz2.open if low.endswith(".bz2") else lzma.open)
        inner_name = os.path.splitext(fname)[0]
        inner_low = inner_name.lower()
        try:
            with open_fn(lp, "rb") as zfp:
                head = zfp.read(4096)
                fmt = _sniff_stream_format(head)
            if inner_low.endswith(".mbox") or fmt == "mbox":
                f_parts = list(folder_parts) if folder_parts is not None else [clean_folder_segment(os.path.splitext(inner_name)[0] or "Inbox")]
                with open_fn(lp, "rb") as zfp:
                    yield from iter_mbox_stream(zfp, f_parts, os.path.splitext(inner_name)[0], include_deleted)
                return
            if fmt == "mmdf":
                with open_fn(lp, "rb") as zfp:
                    yield from iter_mmdf_stream(zfp, default_folder, os.path.splitext(inner_name)[0], include_deleted)
                return
            with open_fn(lp, "rb") as zfp:
                data = zfp.read()
            yield from _iter_records_from_bytes(
                data,
                default_folder,
                inner_name,
                include_deleted=include_deleted,
                ext_flags=ext_flags,
            )
        except Exception:
            pass
        return

    try:
        with open(lp, "rb") as fp:
            head = fp.read(4096)
    except OSError:
        return

    fmt = _sniff_stream_format(head)

    if fmt == "pst":
        yield from iter_pst_records(file_path, folder_parts or [], include_deleted=include_deleted)
        return

    if fmt == "zip":
        yield from iter_zip_or_olm_records(lp, folder_parts or [], include_deleted=include_deleted)
        return

    if low.endswith(".msg") or fmt == "msg":
        try:
            with open(lp, "rb") as fp:
                data = fp.read()
            yield from _iter_records_from_bytes(
                data,
                default_folder,
                fname,
                include_deleted=include_deleted,
                ext_flags=ext_flags,
                native_msg_path=os.path.abspath(file_path),
            )
        except OSError:
            pass
        return

    if low.endswith((".emlx", ".partial.emlx")) or fmt == "emlx":
        try:
            with open(lp, "rb") as fp:
                data = fp.read()
            yield from _iter_records_from_bytes(
                data,
                default_folder,
                fname,
                include_deleted=include_deleted,
                ext_flags=ext_flags,
            )
        except OSError:
            pass
        return

    if fmt == "mmdf":
        with open(lp, "rb") as fp:
            yield from iter_mmdf_stream(fp, default_folder, os.path.splitext(fname)[0], include_deleted)
        return

    if fmt == "babyl":
        with open(lp, "rb") as fp:
            yield from iter_babyl_stream(fp, default_folder, os.path.splitext(fname)[0], include_deleted)
        return

    is_single_msg_ext = low.endswith(".eml") or (ext_flags is not None) or fname.isdigit()
    if not is_single_msg_ext and (low.endswith(".mbox") or fmt == "mbox" or "." not in fname):
        mbox_folder = list(folder_parts) if folder_parts is not None else [clean_folder_segment(fname)]
        with open(lp, "rb") as fp:
            yield from iter_mbox_stream(fp, mbox_folder, clean_folder_segment(fname), include_deleted)
        return

    if fmt in ("eml", "mbox") or low.endswith(".eml"):
        try:
            with open(lp, "rb") as fp:
                data = fp.read()
            yield from _iter_records_from_bytes(
                data,
                default_folder,
                fname,
                include_deleted=include_deleted,
                ext_flags=ext_flags,
            )
        except OSError:
            pass


def iter_mail_records(
    src_path: str,
    include_deleted: bool = False,
) -> Iterator[MailRecord]:
    """
    Universal generator yielding MailRecords from any supported file or directory hierarchy.
    """
    abs_src = os.path.abspath(src_path)
    if os.path.isfile(win_long_path(abs_src)):
        yield from iter_file_records(abs_src, folder_parts=None, include_deleted=include_deleted)
        return

    for root, dirs, files in os.walk(win_long_path(abs_src)):
        dirs[:] = sorted([
            d for d in dirs
            if d.lower() != "tmp" and d != "__MACOSX" and not d.startswith("._")
        ])
        files_sorted = sorted(files)

        clean_root = root[4:] if root.startswith("\\\\?\\") and not abs_src.startswith("\\\\?\\") else root
        rel_dir = os.path.relpath(clean_root, abs_src)
        rel_parts = [] if rel_dir in (".", "") else [p for p in rel_dir.replace("\\", "/").split("/") if p]

        if rel_parts and rel_parts[-1].lower() in ("cur", "new"):
            in_new = rel_parts[-1].lower() == "new"
            folder_parts = normalize_maildir_folder_parts(rel_parts)
            for fname in files_sorted:
                if fname.startswith("."):
                    continue
                md_flags, _ = parse_maildir_flags(fname, in_new=in_new)
                if md_flags.get("is_deleted") and not include_deleted:
                    continue
                fpath = os.path.join(clean_root, fname)
                yield from iter_file_records(
                    fpath,
                    folder_parts=folder_parts,
                    include_deleted=include_deleted,
                    ext_flags=md_flags,
                )
            continue

        mh_seq_path = os.path.join(clean_root, ".mh_sequences")
        mh_sequences = parse_mh_sequences(mh_seq_path) if os.path.exists(win_long_path(mh_seq_path)) else {}
        in_apple_bundle = any(p.lower().endswith((".mbox", ".imapmbox", ".popmbox")) for p in rel_parts)

        for fname in files_sorted:
            low = fname.lower()
            if fname.startswith(".") or low in THUNDERBIRD_SKIP_FILES or low.endswith(THUNDERBIRD_SKIP_SUFFIXES):
                continue

            fpath = os.path.join(clean_root, fname)

            if low.endswith((".emlx", ".partial.emlx")) or in_apple_bundle:
                folder_parts = normalize_apple_mail_folder_parts(rel_parts)
                yield from iter_file_records(
                    fpath,
                    folder_parts=folder_parts,
                    include_deleted=include_deleted,
                )
                continue

            if fname.isdigit():
                folder_parts = [clean_folder_segment(p) for p in rel_parts] or ["Inbox"]
                mh_flags, has_mh = flags_for_mh_message(int(fname), mh_sequences)
                if mh_flags.get("is_deleted") and not include_deleted:
                    continue
                yield from iter_file_records(
                    fpath,
                    folder_parts=folder_parts,
                    include_deleted=include_deleted,
                    ext_flags=(mh_flags if has_mh else None),
                )
                continue

            base_no_ext, ext = os.path.splitext(fname)
            if ext == "" or low.endswith(".mbox") or low.endswith(".mbox.gz"):
                mbox_leaf = clean_folder_segment(
                    fname[:-8] if low.endswith(".mbox.gz") else (base_no_ext if ext else fname)
                )
                folder_parts = [clean_folder_segment(p) for p in rel_parts] + [mbox_leaf]
                yield from iter_file_records(
                    fpath,
                    folder_parts=folder_parts,
                    include_deleted=include_deleted,
                )
                continue

            folder_parts = [clean_folder_segment(p) for p in rel_parts]
            yield from iter_file_records(
                fpath,
                folder_parts=folder_parts,
                include_deleted=include_deleted,
            )
