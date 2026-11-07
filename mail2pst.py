#!/usr/bin/env python3
"""
mail2pst.py — Universal Mail/Mailbox -> Outlook PST Importer (part of mail2mail)

Directly converts and imports any supported email or mailbox format:
  - Standard .mbox & Thunderbird ImapMail/Mail (mbox and Maildir modes)
  - Maildir & Maildir++ (cur/, new/, .Folder.Subfolder)
  - Apple Mail (.emlx, .partial.emlx, *.mbox / *.imapmbox bundles)
  - Outlook .msg (native MAPI OpenSharedItem transfer + pure-Python MS-CFB fallback)
  - Outlook for Mac .olm archives
  - Outlook .pst Personal Folders (for splitting, merging, or filtering PSTs by year/size)
  - MH / Claws Mail / Sylpheed directories
  - MMDF & Emacs Babyl files
  - Standard .eml files & directory trees
  - Compressed archives (.gz, .bz2, .xz, .zip, .tar, .tar.gz, .tgz, .tar.bz2, .tar.xz)
into one or more Microsoft Outlook .pst files via MAPI/COM (pywin32).

Supports both:
  1) Direct single-PST output:
       python mail2pst.py "D:\\Mail\\ImapMail" "D:\\PSTs\\Mailbox.pst"
  2) Multi-PST split / auto-rotated output:
       python mail2pst.py --src "D:\\Export_EML" --out-dir "D:\\PSTs" --split-by year --max-pst-gb 15
"""

from __future__ import annotations

import argparse
import functools
import gc
import mimetypes
import os
import shutil
import sys
import tempfile
import time
from datetime import datetime, timezone
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from typing import Dict, List, Optional, Set, Tuple

try:
    import pythoncom
    import pywintypes
    import win32com.client as win32
    HAS_PYWIN32 = True
except Exception:
    pythoncom = None
    pywintypes = None
    win32 = None
    HAS_PYWIN32 = False

from mail_formats import (
    FORBIDDEN_CHARS,
    MSG_FLAG_EXPUNGED,
    MSG_FLAG_FORWARDED,
    MSG_FLAG_MARKED,
    MSG_FLAG_READ,
    MSG_FLAG_REPLIED,
    WINDOWS_RESERVED_NAMES,
    decode_imap_utf7,
    decode_mime_words,
    decode_modified_utf7,
    extract_message_flags,
    iter_mail_records,
    pick_year_month,
    safe_get_all_headers,
    safe_get_header,
    sanitize_name,
    win_long_path,
)

parse_message_flags = extract_message_flags
sanitize_folder_name = sanitize_name

# ----------------------- MAPI Property Tags (DASL) -----------------------

MAPI_PROPTAG_BASE = "http://schemas.microsoft.com/mapi/proptag/"

# Transport headers & Message-ID
PR_TRANSPORT_MESSAGE_HEADERS_W = MAPI_PROPTAG_BASE + "0x007D001F"
PR_TRANSPORT_MESSAGE_HEADERS_A = MAPI_PROPTAG_BASE + "0x007D001E"
PR_INTERNET_MESSAGE_ID_W       = MAPI_PROPTAG_BASE + "0x1035001F"
PR_INTERNET_MESSAGE_ID_A       = MAPI_PROPTAG_BASE + "0x1035001E"

# Sender & Sent-Representing (Unicode PT_UNICODE = 001F)
PR_SENDER_NAME_W               = MAPI_PROPTAG_BASE + "0x0C1A001F"
PR_SENDER_EMAIL_ADDRESS_W      = MAPI_PROPTAG_BASE + "0x0C1F001F"
PR_SENDER_ADDRTYPE_W           = MAPI_PROPTAG_BASE + "0x0C1E001F"
PR_SENT_REPRESENTING_NAME_W    = MAPI_PROPTAG_BASE + "0x0042001F"
PR_SENT_REPRESENTING_EMAIL_W   = MAPI_PROPTAG_BASE + "0x0065001F"
PR_SENT_REPRESENTING_ADDRTYPE_W = MAPI_PROPTAG_BASE + "0x0064001F"

# Timestamps (PT_SYSTIME = 0040)
PR_CLIENT_SUBMIT_TIME          = MAPI_PROPTAG_BASE + "0x00390040"
PR_MESSAGE_DELIVERY_TIME       = MAPI_PROPTAG_BASE + "0x0E060040"

# Inline HTML Attachment properties
PR_ATTACH_CONTENT_ID_W         = MAPI_PROPTAG_BASE + "0x3712001F"
PR_ATTACH_CONTENT_ID_A         = MAPI_PROPTAG_BASE + "0x3712001E"
PR_ATTACH_FLAGS                = MAPI_PROPTAG_BASE + "0x37140003"

# Replied / Forwarded Verb & Icon properties (PT_LONG = 0003)
PR_LAST_VERB_EXECUTED          = MAPI_PROPTAG_BASE + "0x10810003"
PR_ICON_INDEX                  = MAPI_PROPTAG_BASE + "0x10800003"


# ----------------------- COM Resilience & Helpers -----------------------

def com_retry(max_attempts: int = 5, base_delay: float = 0.25):
    """
    Retry wrapper for transient Outlook COM errors such as RPC_E_CALL_REJECTED
    (-2147418111 / 0x80010001) or SERVERCALL_RETRYLATER (-2147417846 / 0x8001010A).
    """
    def decorator(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            delay = base_delay
            for attempt in range(1, max_attempts + 1):
                try:
                    return fn(*args, **kwargs)
                except Exception as exc:
                    hr = getattr(exc, "hresult", None)
                    transient = hr in (-2147418111, -2147417846) or "call was rejected by callee" in str(exc).lower()
                    if not transient or attempt >= max_attempts:
                        raise
                    if pythoncom is not None:
                        try:
                            pythoncom.PumpWaitingMessages()
                        except Exception:
                            pass
                    time.sleep(delay)
                    delay *= 1.8
        return wrapper
    return decorator


def addresses_to_str(header_val) -> str:
    """
    Convert one or more address headers into a semicolon-separated list for Outlook.
    Replaces semicolons inside display names with commas so Outlook does not split them.
    """
    if not header_val:
        return ""
    if isinstance(header_val, (list, tuple)):
        raw_list = [str(x) for x in header_val if x]
    else:
        raw_list = [str(header_val)]
    pairs = getaddresses(raw_list)
    out = []
    for name, addr in pairs:
        name = (name or "").replace(";", ",").strip()
        addr = (addr or "").strip()
        if name and addr:
            out.append(f"{name} <{addr}>")
        elif addr:
            out.append(addr)
        elif name:
            out.append(name)
    return "; ".join(out)


format_address_list = addresses_to_str


def set_prop(pa, prop_tag: str, val) -> bool:
    """Safely set a single MAPI property via PropertyAccessor; return True on success."""
    try:
        pa.SetProperty(prop_tag, val)
        return True
    except Exception:
        return False


def to_mapi_utc(pa, dt: datetime) -> datetime:
    """
    Convert a Python datetime to a naive UTC datetime for MAPI PT_SYSTIME.
    Wait: pywin32 converts timezone-aware datetimes to local system time when creating
    COM VT_DATE, whereas MAPI PropertyAccessor.SetProperty expects UTC in VT_DATE.
    Passing a naive datetime already in UTC avoids pywin32's local-timezone shift.
    """
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    else:
        dt = dt.astimezone(timezone.utc)
    if dt.year < 1970:
        dt = dt.replace(year=1970)
    elif dt.year > 2100:
        dt = dt.replace(year=2100)
    return dt.replace(tzinfo=None)


# ----------------------- MIME Body & Attachment Extraction -----------------------


def _decode_part_text(part) -> str:
    """Decode a text/plain or text/html MIME part safely into a Python str."""
    try:
        content = part.get_content()
        if isinstance(content, str):
            return content
    except Exception:
        pass
    try:
        payload = part.get_payload(decode=True) or b""
        charset = part.get_content_charset() or "utf-8"
        try:
            return payload.decode(charset, errors="replace")
        except LookupError:
            return payload.decode("utf-8", errors="replace")
    except Exception:
        return ""


def _walk_outer_parts(msg):
    """
    Walk MIME parts of `msg`, traversing `multipart/*` containers iteratively (preventing
    RecursionError on deeply nested MIME) but treating encapsulated `message/rfc822`
    parts (RFC 2046 §5.2.1) as atomic leaf parts so their inner bodies/attachments
    are not merged into the outer message.
    """
    stack = [msg]
    while stack:
        part = stack.pop()
        if part.get_content_maintype() == "multipart":
            payload = part.get_payload()
            if isinstance(payload, list):
                for sub in reversed(payload):
                    if hasattr(sub, "get_content_maintype"):
                        stack.append(sub)
        else:
            yield part


def pick_body(msg) -> Tuple[Optional[str], Optional[str], Set[int]]:
    """
    Extract (html, text, body_part_ids) from an email.message.Message.
    Tracks the exact MIME parts used as the body so they are not duplicated as attachments.
    """
    html = None
    text = None
    body_parts: Set[int] = set()

    if msg.get_content_maintype() == "multipart":
        for part in _walk_outer_parts(msg):
            ctype = part.get_content_type()
            disp = safe_get_header(part, "Content-Disposition", "").lower()
            try:
                fn = part.get_filename()
            except Exception:
                fn = None
            if "attachment" in disp or fn:
                continue
            if ctype == "text/html" and html is None:
                html = _decode_part_text(part)
                body_parts.add(id(part))
            elif ctype == "text/plain" and text is None:
                text = _decode_part_text(part)
                body_parts.add(id(part))
    else:
        ctype = msg.get_content_type()
        if ctype == "text/html":
            html = _decode_part_text(msg)
            body_parts.add(id(msg))
        elif ctype == "text/plain" or ctype.startswith("text/"):
            text = _decode_part_text(msg)
            body_parts.add(id(msg))

    return html, text, body_parts


def add_attachments(mail, msg, tmpdir: str, body_parts: Optional[Set[int]] = None) -> Tuple[int, bool]:
    """
    Save all non-body attachments and inline CID assets to temp files and attach to `mail`.
    Preserves original filenames and links Content-ID for inline HTML images.
    Returns (count, has_inline_cid).
    """
    if body_parts is None:
        body_parts = set()

    count = 0
    has_cid = False
    top_ctype = msg.get_content_type()
    is_opaque_smime = (
        msg.get_content_maintype() != "multipart"
        and ("pkcs7-mime" in top_ctype or "x-pkcs7-mime" in top_ctype)
    )

    for part in _walk_outer_parts(msg):
        if id(part) in body_parts:
            continue

        ctype = part.get_content_type()
        disp = safe_get_header(part, "Content-Disposition", "").lower()
        cid = safe_get_header(part, "Content-ID", "") or None
        try:
            fn = part.get_filename()
        except Exception:
            fn = None

        if ctype in ("application/pgp-signature", "application/pkcs7-signature", "application/x-pkcs7-signature"):
            if not fn:
                fn = "signature.asc" if "pgp" in ctype else "smime.p7s"

        is_rfc822 = (ctype == "message/rfc822")
        is_calendar = ctype in ("text/calendar", "application/ics")
        is_delivery_report = ctype in ("message/delivery-status", "message/disposition-notification")
        is_explicit_att = (
            ("attachment" in disp)
            or (fn is not None)
            or is_rfc822
            or is_opaque_smime
            or is_calendar
            or is_delivery_report
        )
        is_inline_asset = (cid is not None) or (
            "inline" in disp and ctype not in ("text/plain", "text/html")
        )

        if not (is_explicit_att or is_inline_asset):
            continue

        payload = None
        if is_rfc822 or is_delivery_report or part.get_content_maintype() == "message":
            try:
                sub = part.get_payload()
                if isinstance(sub, list) and sub and hasattr(sub[0], "as_bytes"):
                    payload = b"\r\n".join(s.as_bytes() for s in sub)
                elif hasattr(part, "as_bytes"):
                    payload = part.as_bytes()
                else:
                    payload = part.get_payload(decode=True)
            except Exception:
                payload = None
        else:
            try:
                payload = part.get_payload(decode=True)
            except Exception:
                payload = None

        if payload is None:
            continue

        if not fn:
            if is_opaque_smime and "pkcs7" in ctype:
                fn = "smime.p7m"
            elif is_rfc822:
                fn = f"forwarded_message_{count + 1}.eml"
            elif is_calendar:
                fn = f"invite_{count + 1}.ics"
            elif is_delivery_report:
                fn = f"delivery_status_{count + 1}.txt"
            else:
                ext = mimetypes.guess_extension(ctype or "") or ".bin"
                fn = f"attachment_{count + 1}{ext}"

        slot_dir = os.path.abspath(os.path.join(tmpdir, f"s_{count}"))
        os.makedirs(slot_dir, exist_ok=True)
        # Keep total path length safely under Windows MAX_PATH (260 chars) for Outlook COM
        avail_len = max(8, 240 - len(slot_dir) - 1)
        safe_fn = sanitize_name(fn, default=f"attachment_{count + 1}.bin", max_len=avail_len)
        p = os.path.realpath(os.path.join(slot_dir, safe_fn))
        if len(p) >= 250:
            p = os.path.realpath(os.path.join(slot_dir, f"a_{count}.bin"))

        try:
            with open(p, "wb") as f:
                f.write(payload)
            a = mail.Attachments.Add(p, 1, 1, safe_fn)
            if cid:
                clean_cid = str(cid).strip().strip("<>")
                if clean_cid:
                    has_cid = True
                    pa = a.PropertyAccessor
                    if not set_prop(pa, PR_ATTACH_CONTENT_ID_W, clean_cid):
                        set_prop(pa, PR_ATTACH_CONTENT_ID_A, clean_cid)
                    # 4 = ATT_MHTML_REF (renders inline in HTMLBody)
                    set_prop(pa, PR_ATTACH_FLAGS, 4)
            del a
            count += 1
        except Exception as e:
            print(f"\n[WARN] Attachment '{safe_fn}' failed: {e}")
        finally:
            shutil.rmtree(slot_dir, ignore_errors=True)

    return count, has_cid


def build_headers_text(raw_bytes: bytes) -> str:
    """Extract raw header block (up to the first blank line, CRLF or LF) for PR_TRANSPORT_MESSAGE_HEADERS."""
    try:
        idx_crlf = raw_bytes.find(b"\r\n\r\n")
        idx_lf = raw_bytes.find(b"\n\n")
        candidates = [i for i in (idx_crlf, idx_lf) if i != -1]
        if candidates:
            end = min(candidates)
        else:
            end = min(len(raw_bytes), 1024 * 32)
        hdr = raw_bytes[:end]
        return hdr.decode("utf-8", errors="replace")
    except Exception:
        return ""


def handle_sender_info(mail_item, parsed_msg) -> None:
    """Extract From header and inject Sender & SentRepresenting MAPI properties."""
    from_headers = safe_get_all_headers(parsed_msg, "From")
    if not from_headers:
        return

    pairs = getaddresses(from_headers)
    if not pairs:
        return

    sender_name, sender_email = (pairs[0][0] or "").strip(), (pairs[0][1] or "").strip()
    display = sender_name or sender_email
    if not display and not sender_email:
        return

    pa = mail_item.PropertyAccessor
    set_prop(pa, PR_SENDER_NAME_W, display)
    set_prop(pa, PR_SENDER_EMAIL_ADDRESS_W, sender_email or display)
    set_prop(pa, PR_SENDER_ADDRTYPE_W, "SMTP")
    set_prop(pa, PR_SENT_REPRESENTING_NAME_W, display)
    set_prop(pa, PR_SENT_REPRESENTING_EMAIL_W, sender_email or display)
    set_prop(pa, PR_SENT_REPRESENTING_ADDRTYPE_W, "SMTP")

    try:
        mail_item.SentOnBehalfOfName = display
    except Exception:
        pass


@com_retry()
def populate_mailitem(mail_item, parsed, msg_bytes: bytes, temp_dir: str, flags=None) -> None:
    """
    Populate an Outlook MailItem from a parsed email.message.Message and raw bytes,
    then call mail_item.Save().
    """
    pa = mail_item.PropertyAccessor

    # 1. Basic OOM fields (clamp Subject to Outlook COM's 250-char boundary)
    raw_subj = safe_get_header(parsed, "Subject", "")
    mail_item.Subject = raw_subj[:250] if len(raw_subj) > 250 else raw_subj
    to_str = format_address_list(safe_get_all_headers(parsed, "To"))
    if to_str:
        mail_item.To = to_str
    cc_str = format_address_list(safe_get_all_headers(parsed, "Cc"))
    if cc_str:
        mail_item.CC = cc_str
    bcc_str = format_address_list(safe_get_all_headers(parsed, "Bcc"))
    if bcc_str:
        mail_item.BCC = bcc_str

    # 2. Extract body & attachments before setting HTMLBody so inline CIDs link cleanly
    top_ctype = parsed.get_content_type()
    is_opaque_smime = (
        not parsed.is_multipart()
        and ("pkcs7-mime" in top_ctype or "x-pkcs7-mime" in top_ctype)
    )
    html_body, text_body, body_parts = pick_body(parsed)
    if is_opaque_smime and html_body is None and text_body is None:
        mail_item.Body = "[S/MIME Encrypted or Opaque-Signed Message - See attached smime.p7m]"

    add_attachments(mail_item, parsed, temp_dir, body_parts=body_parts)

    if html_body is not None:
        mail_item.HTMLBody = html_body
    elif text_body is not None:
        mail_item.Body = text_body
    elif not is_opaque_smime:
        mail_item.Body = ""

    # 3. Read / Flagged / Replied / Forwarded states
    if flags is None:
        flags = extract_message_flags(parsed)
    is_read, is_flagged, is_replied, is_forwarded, _ = flags

    try:
        mail_item.UnRead = not is_read
    except Exception:
        pass

    if is_flagged:
        try:
            mail_item.FlagStatus = 2  # olFlagMarked
            mail_item.FlagRequest = "Follow up"
        except Exception:
            pass

    # 4. Low-level MAPI properties set after Body/HTMLBody so OOM body sync does not reset them
    raw_headers_text = build_headers_text(msg_bytes)
    if raw_headers_text:
        if not set_prop(pa, PR_TRANSPORT_MESSAGE_HEADERS_W, raw_headers_text):
            set_prop(pa, PR_TRANSPORT_MESSAGE_HEADERS_A, raw_headers_text)

    mid = safe_get_header(parsed, "Message-ID", "").strip()
    if mid:
        if not set_prop(pa, PR_INTERNET_MESSAGE_ID_W, mid):
            set_prop(pa, PR_INTERNET_MESSAGE_ID_A, mid)

    handle_sender_info(mail_item, parsed)

    date_hdr = safe_get_header(parsed, "Date", "").strip()
    if date_hdr:
        try:
            dt = parsedate_to_datetime(date_hdr)
            if dt is not None:
                utc_dt = to_mapi_utc(pa, dt)
                set_prop(pa, PR_MESSAGE_DELIVERY_TIME, utc_dt)
                set_prop(pa, PR_CLIENT_SUBMIT_TIME, utc_dt)
        except Exception:
            pass

    if is_replied:
        set_prop(pa, PR_LAST_VERB_EXECUTED, 102)
        set_prop(pa, PR_ICON_INDEX, 261)
    elif is_forwarded:
        set_prop(pa, PR_LAST_VERB_EXECUTED, 104)
        set_prop(pa, PR_ICON_INDEX, 262)

    mail_item.Save()


@com_retry()
def create_mail_in_dest(dest_folder, msg, raw_headers_text: str, tmpdir: str):
    """
    Create MailItem directly in the target PST folder (avoids default Drafts),
    populate OOM fields, attachments, body, and MAPI properties.
    """
    mail = dest_folder.Items.Add(0)  # 0 = olMailItem
    raw_bytes = raw_headers_text.encode("utf-8", errors="replace") + b"\r\n\r\n" if raw_headers_text else b""
    populate_mailitem(mail, msg, raw_bytes, tmpdir)
    return mail


@com_retry()
def import_native_msg_in_dest(ns, dest_folder, msg_path: str):
    """
    Open a standalone Outlook .msg file via MAPI OpenSharedItem, copy it,
    and move the copy directly into dest_folder for 100% MAPI fidelity.
    """
    shared_item = ns.OpenSharedItem(os.path.abspath(msg_path))
    try:
        copied = shared_item.Copy()
        moved = copied.Move(dest_folder)
        moved.Save()
        return moved
    finally:
        try:
            shared_item.Close(1)  # 1 = olDiscard
        except Exception:
            pass
        del shared_item


def import_native_msg_to_outlook(namespace, outlook_folder, msg_path: str):
    """Open and import a native .msg into outlook_folder and release the COM reference."""
    moved = import_native_msg_in_dest(namespace, outlook_folder, msg_path)
    del moved


# ----------------------- Outlook Store & Folder Management -----------------------

def ensure_outlook_with_logon():
    """Start Outlook.Application and log on to the default MAPI profile."""
    pythoncom.CoInitialize()
    app = win32.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    ns.Logon("", "", False, True)
    return app, ns


def normcasepath(p: str) -> str:
    return os.path.normcase(os.path.normpath(p))


def find_store_by_path(ns, pst_path: str):
    """Return an Outlook Store matched by FilePath, else None."""
    target = normcasepath(pst_path)
    for store in ns.Stores:
        try:
            current = normcasepath(store.FilePath)
            if current == target:
                return store
        except Exception:
            pass
    return None


def create_or_attach_pst(ns, desired_path: str, overwrite: bool = False):
    """
    Create or attach a PST at desired_path; return (store, root_folder, actual_path).
    When overwrite=False (e.g. during periodic flush/reattach), preserves existing PST data.
    """
    desired_path = os.path.normpath(desired_path)
    out_dir = os.path.dirname(desired_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    exist = find_store_by_path(ns, desired_path)
    if exist:
        if not overwrite:
            return exist, exist.GetRootFolder(), desired_path
        try:
            ns.RemoveStore(exist.GetRootFolder())
        except Exception:
            pass

    if overwrite and os.path.exists(desired_path):
        try:
            os.remove(desired_path)
        except Exception:
            pass

    existing_ids = set()
    for s in ns.Stores:
        try:
            existing_ids.add(s.StoreID)
        except Exception:
            pass

    actual = desired_path
    try:
        ns.AddStoreEx(actual, 2)  # 2 = olStoreUnicode
    except Exception:
        try:
            ns.AddStoreEx(actual, 1)  # 1 = olStoreDefault
        except Exception:
            ns.AddStore(actual)

    for _ in range(20):
        for store in ns.Stores:
            try:
                if normcasepath(store.FilePath) == normcasepath(actual):
                    return store, store.GetRootFolder(), actual
            except Exception:
                pass
        for store in ns.Stores:
            try:
                if store.StoreID not in existing_ids:
                    return store, store.GetRootFolder(), store.FilePath or actual
            except Exception:
                pass
        time.sleep(0.25)

    last = None
    for s in ns.Stores:
        last = s
    if last:
        return last, last.GetRootFolder(), last.FilePath

    raise RuntimeError(f"Failed to create/attach PST: {desired_path}")


@com_retry()
def get_or_create_subfolder(parent_folder, folder_name: str):
    """Safely retrieve or create a subfolder inside an Outlook MAPI Folder."""
    clean_name = sanitize_folder_name(folder_name)
    try:
        return parent_folder.Folders(clean_name)
    except Exception:
        try:
            return parent_folder.Folders.Item(clean_name)
        except Exception:
            return parent_folder.Folders.Add(clean_name)


@com_retry()
def get_or_create_folder_path(root_folder, pst_root_name: str, rel_dir: str):
    """Navigate or create pst_root_name and any relative subfolder segments inside root_folder."""
    curr = root_folder
    if pst_root_name:
        safe_root = sanitize_name(pst_root_name, default="Imported")
        try:
            curr = curr.Folders.Item(safe_root)
        except Exception:
            curr = curr.Folders.Add(safe_root)

    if rel_dir and rel_dir != ".":
        for part in rel_dir.replace("/", os.sep).split(os.sep):
            if not part or part == ".":
                continue
            safe_part = sanitize_name(part, default="Folder")
            try:
                curr = curr.Folders.Item(safe_part)
            except Exception:
                curr = curr.Folders.Add(safe_part)
    return curr


class PstRouter:
    """
    Decide which PST (Store/Root) should receive the next item:
      - direct single-file path (auto-rotates to _part2.pst if > max_gb)
      - by year: one PST per year (auto-rotates if > max_gb)
      - by even split: target_bytes computed from total size / splits
    """
    def __init__(
        self,
        ns,
        out_dir: str,
        base_name: str,
        max_gb: float,
        splits: Optional[int],
        split_by_year: bool,
        overwrite: bool = False,
        single_pst_path: Optional[str] = None,
    ):
        self.ns = ns
        self.out_dir = os.path.normpath(out_dir)
        self.base = base_name
        self.max_bytes = int(max_gb * (1024**3)) if max_gb and max_gb > 0 else 0
        self.splits = splits
        self.by_year = split_by_year
        self.overwrite = overwrite
        self.single_pst_path = os.path.abspath(single_pst_path) if single_pst_path else None
        self.cur = {"store": None, "root": None, "path": None, "bytes": 0, "part": 0}
        self.year: Dict[int, Dict] = {}
        self.target_bytes: Optional[int] = None
        self.open_roots: Dict[str, object] = {}

    def set_total_bytes(self, total_bytes: int) -> None:
        if self.splits and self.splits > 0:
            self.target_bytes = max(1, total_bytes // self.splits)

    def _new_path(self, suffix: str) -> str:
        if self.single_pst_path and suffix == "part1":
            return self.single_pst_path
        return os.path.join(self.out_dir, f"{self.base}_{suffix}.pst")

    def _register_open(self, path: str, root) -> None:
        if path and root:
            self.open_roots[normcasepath(path)] = root

    def _detach_path(self, path: Optional[str]) -> None:
        if not path:
            return
        key = normcasepath(path)
        root = self.open_roots.pop(key, None)
        if root is not None:
            try:
                self.ns.RemoveStore(root)
            except Exception:
                pass

    def _initial_existing_bytes(self, desired_path: str) -> int:
        if not self.overwrite and os.path.isfile(desired_path):
            try:
                return os.path.getsize(desired_path)
            except OSError:
                return 0
        return 0

    def _rotate_new(self):
        if self.cur.get("path"):
            self._detach_path(self.cur["path"])
        self.cur["part"] += 1
        desired = self._new_path(f"part{self.cur['part']}")
        init_bytes = self._initial_existing_bytes(desired)
        store, root, actual = create_or_attach_pst(self.ns, desired, overwrite=self.overwrite)
        self.cur.update({"store": store, "root": root, "path": actual, "bytes": init_bytes})
        self._register_open(actual, root)
        print(f"\nUSING PST: {actual}")
        return store, root, actual

    def _need_rotate_seq(self, incoming: int) -> bool:
        if self.max_bytes and self.cur["store"] and self.cur["bytes"] > 0 and (self.cur["bytes"] + incoming > self.max_bytes):
            return True
        if self.target_bytes and self.cur["store"] and (self.cur["bytes"] >= self.target_bytes):
            return True
        return False

    def _ensure_year(self, year: int, incoming: int):
        s = self.year.get(year)
        if s is None:
            desired = self._new_path(f"{year}_part1")
            init_bytes = self._initial_existing_bytes(desired)
            store, root, actual = create_or_attach_pst(self.ns, desired, overwrite=self.overwrite)
            s = {"path": actual, "store": store, "root": root, "bytes": init_bytes, "part": 1}
            self.year[year] = s
            self._register_open(actual, root)
            print(f"\nUSING PST ({year}): {actual}")
        elif self.max_bytes and s["bytes"] > 0 and (s["bytes"] + incoming > self.max_bytes):
            self._detach_path(s.get("path"))
            s["part"] += 1
            desired = self._new_path(f"{year}_part{s['part']}")
            init_bytes = self._initial_existing_bytes(desired)
            store, root, actual = create_or_attach_pst(self.ns, desired, overwrite=self.overwrite)
            s.update({"path": actual, "store": store, "root": root, "bytes": init_bytes})
            self._register_open(actual, root)
            print(f"\nROTATED PST ({year}) -> {actual}")
        return s

    def route(self, approx_size: int, year_hint: Optional[int] = None):
        """Return (store, root, path) to drop the next item."""
        size = approx_size
        if self.by_year:
            y = year_hint or 1970
            s = self._ensure_year(y, size)
            s["bytes"] += size
            return s["store"], s["root"], s["path"]
        if self.cur["store"] is None or self._need_rotate_seq(size):
            self._rotate_new()
        self.cur["bytes"] += size
        return self.cur["store"], self.cur["root"], self.cur["path"]

    def flush_pst(self, pst_path: str):
        """Safely detach and reattach pst_path without deleting it, updating internal references."""
        self._detach_path(pst_path)
        store, root, actual = create_or_attach_pst(self.ns, pst_path, overwrite=False)
        self._register_open(actual, root)
        if self.cur.get("path") and normcasepath(self.cur["path"]) == normcasepath(pst_path):
            self.cur.update({"store": store, "root": root, "path": actual})
        for s in self.year.values():
            if s.get("path") and normcasepath(s["path"]) == normcasepath(pst_path):
                s.update({"store": store, "root": root, "path": actual})
        return store, root, actual

    def close_all(self) -> None:
        """Detach all open PST stores tracked by this router."""
        for key, root in list(self.open_roots.items()):
            try:
                self.ns.RemoveStore(root)
            except Exception:
                pass
            self.open_roots.pop(key, None)


def estimate_source_bytes(src_path: str) -> int:
    """Estimate total source bytes across a file, archive, or directory tree."""
    lp = win_long_path(src_path)
    if os.path.isfile(lp):
        try:
            return os.path.getsize(lp)
        except OSError:
            return 0
    total = 0
    for base, _, fnames in os.walk(lp):
        for fn in fnames:
            if fn.lower().endswith((".msf", ".dat", ".sqlite", ".json")):
                continue
            try:
                total += os.path.getsize(os.path.join(base, fn))
            except OSError:
                pass
    return total


# ----------------------- Unified mail2pst CLI -----------------------

def build_arg_parser(prog: Optional[str] = None) -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog=prog,
        description=(
            "mail2pst — Convert and import any mail format (EML, MBOX, Thunderbird ImapMail/Mail, "
            "Maildir/Maildir++, Apple Mail .emlx/.mbox, Outlook .msg/.olm, MH, MMDF, Babyl, "
            "or .zip/.tgz/.gz archives) into Microsoft Outlook .pst file(s)."
        ),
    )
    ap.add_argument("src_pos", nargs="?", help="Source mail directory, mailbox file, or archive")
    ap.add_argument("out_pos", nargs="?", help="Output .pst file path or destination directory")
    ap.add_argument("--src", "--mbox", dest="src_opt", help="Source mail directory, mailbox file, or archive")
    ap.add_argument("--out", "--pst", dest="out_pst", help="Output .pst file path (for single-PST mode)")
    ap.add_argument("--out-dir", dest="out_dir", help="Directory to write .pst file(s) (for multi-PST split/rotated mode)")
    ap.add_argument("--base-name", default="emails", help="Base name for PST files when using --out-dir (default: emails)")
    ap.add_argument("--split-by", choices=["year"], default=None, help="Split by mail year (one PST per year)")
    ap.add_argument("--splits", type=int, default=None, help="Even-split into N PSTs by total bytes")
    ap.add_argument("--max-pst-gb", type=float, default=15.0, help="Max ~GB per PST before rotating part2/part3 (default: 15.0, 0=unlimited)")
    ap.add_argument("--pst-root", default=None, help="Top folder name inside each PST (default: root for single-PST, 'Imported (EML)' for --out-dir)")
    ap.add_argument("--flatten", action="store_true", help="Do not replicate source subfolder structure inside PST")
    ap.add_argument("--start-year", type=int, default=None, help="Import messages from this year (inclusive)")
    ap.add_argument("--end-year", type=int, default=None, help="Import messages up to this year (inclusive)")
    ap.add_argument("--overwrite-pst", action="store_true", help="Overwrite existing target PST files if present")
    ap.add_argument("--include-deleted", action="store_true", help="Import messages marked expunged/deleted")
    ap.add_argument("--flush-every", type=int, default=0, help="Detach/reattach PST every N items (0=off)")
    ap.add_argument("--count-every", type=int, default=200, help="Print folder Items.Count every N items (0=off)")
    return ap


def main(argv: Optional[List[str]] = None, prog: Optional[str] = None) -> None:
    for _stream in (sys.stdout, sys.stderr):
        if hasattr(_stream, "reconfigure"):
            try:
                _stream.reconfigure(errors="replace")
            except Exception:
                pass

    ap = build_arg_parser(prog=prog)
    args = ap.parse_args(argv)

    if not HAS_PYWIN32:
        print("Error: pywin32 is not installed. Run:  pip install pywin32", file=sys.stderr)
        sys.exit(1)

    src = args.src_opt or args.src_pos
    target = args.out_pst or args.out_dir or args.out_pos
    if not src or not target:
        ap.error("Both a source path (positional or --src) and an output target (positional, --out, or --out-dir) are required.")

    src = os.path.normpath(src)
    if not os.path.exists(win_long_path(src)):
        print(f"Source path not found: {src}", file=sys.stderr)
        sys.exit(1)

    # Determine whether target is a specific .pst file or an output directory
    single_pst_path: Optional[str] = None
    if args.out_pst or (args.out_pos and args.out_pos.lower().endswith(".pst") and not args.out_dir):
        single_pst_path = os.path.abspath(args.out_pst or args.out_pos)
        out_dir = os.path.dirname(single_pst_path) or "."
        base_name = os.path.splitext(os.path.basename(single_pst_path))[0]
        pst_root_name = args.pst_root if args.pst_root is not None else ""
    else:
        out_dir = os.path.normpath(args.out_dir or args.out_pos)
        base_name = args.base_name
        pst_root_name = args.pst_root if args.pst_root is not None else "Imported (EML)"

    os.makedirs(out_dir, exist_ok=True)
    total_bytes = estimate_source_bytes(src)
    print(f"Source: {src} | Estimated size: {total_bytes / 1e9:,.2f} GB")

    app, ns = ensure_outlook_with_logon()

    router = PstRouter(
        ns=ns,
        out_dir=out_dir,
        base_name=base_name,
        max_gb=args.max_pst_gb,
        splits=args.splits,
        split_by_year=(args.split_by == "year"),
        overwrite=args.overwrite_pst,
        single_pst_path=single_pst_path,
    )
    router.set_total_bytes(total_bytes)

    parser = BytesParser(policy=policy.default)
    start = time.perf_counter()
    processed = 0
    imported = 0
    skipped = 0
    done_bytes = 0
    tmpdir = tempfile.mkdtemp(prefix="mail2pst_")
    current_pst = None

    try:
        for rec in iter_mail_records(src, include_deleted=True):
            raw = rec.raw_bytes
            sz = len(raw)
            processed += 1
            done_bytes += sz

            try:
                msg = parser.parsebytes(raw)
            except Exception:
                msg = parser.parsebytes(b"Subject: (no subject)\r\n\r\n")

            flags = extract_message_flags(msg)
            if (rec.is_deleted or flags[4]) and not args.include_deleted:
                skipped += 1
                continue

            year, _ = pick_year_month(msg)

            if args.start_year and year < args.start_year:
                skipped += 1
                continue
            if args.end_year and year > args.end_year:
                skipped += 1
                continue

            store, root, pst_path = router.route(sz, year)
            if pst_path != current_pst:
                current_pst = pst_path
                print(f"\nCURRENT PST: {current_pst}")

            folder_parts = rec.folder_parts
            if single_pst_path and not pst_root_name and not folder_parts:
                folder_parts = ["Inbox"]
            rel_dir = "." if (args.flatten or not folder_parts) else os.sep.join(folder_parts)
            dest = get_or_create_folder_path(root, pst_root_name, rel_dir)

            try:
                if pythoncom is not None:
                    try:
                        pythoncom.PumpWaitingMessages()
                    except Exception:
                        pass
                item = None
                if rec.native_msg_path and os.path.isfile(win_long_path(rec.native_msg_path)):
                    try:
                        item = import_native_msg_in_dest(ns, dest, rec.native_msg_path)
                    except Exception:
                        item = None

                if item is None:
                    item = dest.Items.Add(0)
                    populate_mailitem(item, msg, raw, tmpdir, flags=flags)
                    try:
                        if item.Parent and item.Parent.EntryID != dest.EntryID:
                            moved = item.Move(dest)
                            moved.Save()
                            del moved
                    except Exception:
                        pass

                del item
                imported += 1
            except Exception as e:
                print(f"\n[WARN] Failed to save item from {rec.source_name}: {e}")
                skipped += 1

            if args.count_every and processed % args.count_every == 0:
                try:
                    print(f"\n  Items in '{dest.Name}': {dest.Items.Count}")
                except Exception:
                    pass
                gc.collect()

            if args.flush_every and processed % args.flush_every == 0:
                try:
                    del dest
                    gc.collect()
                    store, root, pst_path = router.flush_pst(pst_path)
                    print(f"\n[FLUSH] Re-attached PST: {pst_path}")
                except Exception as e:
                    print(f"\n[WARN] Periodic flush failed: {e}")

            if processed % 100 == 0:
                pct = min(100.0, (done_bytes / max(total_bytes, 1)) * 100.0)
                elapsed = time.perf_counter() - start
                rate = done_bytes / max(elapsed, 1e-9)
                eta = max(0.0, (total_bytes - done_bytes)) / max(rate, 1e-9)
                h, m, s = int(eta // 3600), int(eta % 3600 // 60), int(eta % 60)
                print(
                    f"\rProcessed: {processed:,} ({pct:6.2f}%) | "
                    f"ETA {h:02d}:{m:02d}:{s:02d} | PST: {current_pst}",
                    end="",
                )

        if processed == 0:
            print("No supported email messages found in source.", file=sys.stderr)
            sys.exit(1)

        print("\nDone.")
        print(f"Processed: {processed:,} | Imported: {imported:,} | Skipped: {skipped:,}")
        print("PST(s) written to:", single_pst_path or out_dir)

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
        try:
            router.close_all()
        except Exception as e:
            print("\n[WARN] Failed to close stores:", e)
        try:
            gc.collect()
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    main()
