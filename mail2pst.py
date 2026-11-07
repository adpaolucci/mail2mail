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
    extract_attachment_filename,
    extract_best_datetime,
    extract_message_flags,
    iter_mail_records,
    pick_year_month,
    safe_get_all_headers,
    safe_get_header,
    sanitize_name,
    win_long_path,
)
from mail_audit import AuditLogger, AuditRecord, compute_sha256, perform_dry_run
from mail_filter import DateFilter, FolderFilter, MigrationState, parse_size_bytes
from mail_shutdown import ShutdownController

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
PR_CREATION_TIME               = MAPI_PROPTAG_BASE + "0x30070040"
PR_LAST_MODIFICATION_TIME      = MAPI_PROPTAG_BASE + "0x30080040"

# Message Flags (PT_LONG = 0003)
PR_MESSAGE_FLAGS               = MAPI_PROPTAG_BASE + "0x0E070003"
MSGFLAG_READ                   = 0x00000001
MSGFLAG_UNMODIFIED             = 0x00000002
MSGFLAG_SUBMIT                 = 0x00000004
MSGFLAG_UNSENT                 = 0x00000008
MSGFLAG_HASATTACH              = 0x00000010

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
        day = min(dt.day, 28) if dt.month == 2 else dt.day
        dt = dt.replace(year=1970, day=day)
    elif dt.year > 2100:
        day = min(dt.day, 28) if dt.month == 2 else dt.day
        dt = dt.replace(year=2100, day=day)
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
            fn = extract_attachment_filename(part)
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


def check_or_lift_outlook_attachment_limit() -> None:
    """
    Check if Outlook's MaximumAttachmentSize registry key is set to 0 (unlimited).
    Attempts to set it to 0 in HKCU if missing or non-zero, preventing
    'The file you're attaching is bigger than the server allows' COM errors.
    """
    if sys.platform != "win32":
        return
    try:
        import winreg
        for ver in ("16.0", "15.0", "14.0"):
            key_path = f"Software\\Microsoft\\Office\\{ver}\\Outlook\\Preferences"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ | winreg.KEY_SET_VALUE) as k:
                    try:
                        val, _ = winreg.QueryValueEx(k, "MaximumAttachmentSize")
                        if val != 0:
                            winreg.SetValueEx(k, "MaximumAttachmentSize", 0, winreg.REG_DWORD, 0)
                    except FileNotFoundError:
                        winreg.SetValueEx(k, "MaximumAttachmentSize", 0, winreg.REG_DWORD, 0)
            except Exception:
                pass
    except Exception:
        pass


def add_attachments(
    mail,
    msg,
    tmpdir: str,
    body_parts: Optional[Set[int]] = None,
    skipped_oversized: Optional[List[Tuple[str, int]]] = None,
    out_dir: Optional[str] = None,
) -> Tuple[int, bool]:
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
        fn = extract_attachment_filename(part)

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
            err_msg = str(e)
            sz_mb = len(payload) / (1024 * 1024)
            saved_disk_path = ""
            target_save_dir = os.path.join(out_dir or tmpdir, "_unattached_files")
            try:
                os.makedirs(target_save_dir, exist_ok=True)
                disk_fn = safe_fn
                candidate_path = os.path.join(target_save_dir, disk_fn)
                if os.path.exists(win_long_path(candidate_path)):
                    bname, bext = os.path.splitext(disk_fn)
                    c_idx = 1
                    while os.path.exists(win_long_path(os.path.join(target_save_dir, f"{bname}_{c_idx}{bext}"))):
                        c_idx += 1
                    candidate_path = os.path.join(target_save_dir, f"{bname}_{c_idx}{bext}")
                with open(win_long_path(candidate_path), "wb") as f_u:
                    f_u.write(payload)
                saved_disk_path = candidate_path
            except Exception:
                pass

            if "bigger than the server allows" in err_msg or "4096" in err_msg:
                if skipped_oversized is not None:
                    skipped_oversized.append((safe_fn, len(payload)))
                print(
                    f"\n[WARN] Attachment '{safe_fn}' ({sz_mb:.1f} MB) exceeds Outlook's server attachment limit.\n"
                    f"       To lift this limit permanently in Outlook, run in PowerShell:\n"
                    f"       New-ItemProperty -Path 'HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook\\Preferences' -Name 'MaximumAttachmentSize' -Value 0 -PropertyType DWORD -Force"
                )
                if saved_disk_path:
                    print(f"       [Preserved] Saved unattached file safely to disk: {saved_disk_path}")
            else:
                print(f"\n[WARN] Attachment '{safe_fn}' failed: {e}")
                if saved_disk_path:
                    print(f"       [Preserved] Saved unattached file safely to disk: {saved_disk_path}")
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
def populate_mailitem(mail_item, parsed, msg_bytes: bytes, temp_dir: str, flags=None, out_dir: Optional[str] = None) -> None:
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

    oversized: List[Tuple[str, int]] = []
    att_count, _ = add_attachments(mail_item, parsed, temp_dir, body_parts=body_parts, skipped_oversized=oversized, out_dir=out_dir)

    oversized_note = ""
    if oversized:
        items_str = "\n".join(f" - {fn} ({sz / (1024*1024):.1f} MB)" for fn, sz in oversized)
        oversized_note = (
            f"\n\n[mail2pst Note: The following attachment(s) could not be embedded because Outlook's "
            f"MaximumAttachmentSize server limit was exceeded:\n{items_str}\n"
            f"To allow large attachments, set MaximumAttachmentSize=0 in HKCU\\Software\\Microsoft\\Office\\16.0\\Outlook\\Preferences]"
        )

    if html_body is not None:
        if oversized_note:
            html_body += f"<br><br><pre style='color:#b00;font-size:12px;'>{oversized_note.strip()}</pre>"
        mail_item.HTMLBody = html_body
    elif text_body is not None:
        mail_item.Body = text_body + oversized_note
    elif not is_opaque_smime:
        mail_item.Body = oversized_note.strip()

    # 3. Read / Flagged / Replied / Forwarded / Draft states
    if flags is None:
        flags = extract_message_flags(parsed)
    is_read, is_flagged, is_replied, is_forwarded, is_draft = flags

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

    dt = extract_best_datetime(parsed)
    if dt is not None:
        try:
            utc_dt = to_mapi_utc(pa, dt)
            set_prop(pa, PR_MESSAGE_DELIVERY_TIME, utc_dt)
            set_prop(pa, PR_CLIENT_SUBMIT_TIME, utc_dt)
            set_prop(pa, PR_CREATION_TIME, utc_dt)
            set_prop(pa, PR_LAST_MODIFICATION_TIME, utc_dt)
            try:
                mail_item.SentOn = utc_dt
            except Exception:
                pass
        except Exception:
            pass

    # 5. MAPI Message Flags (clear MSGFLAG_UNSENT for received/sent mail so Outlook
    # does not treat them as drafts or override the received date with the import date)
    try:
        raw_msg_flags = pa.GetProperty(PR_MESSAGE_FLAGS)
        cur_flags = int(raw_msg_flags)
    except Exception:
        cur_flags = 0

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

    set_prop(pa, PR_MESSAGE_FLAGS, cur_flags)

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
        for _attempt in range(5):
            try:
                os.remove(desired_path)
                break
            except PermissionError:
                time.sleep(0.25)
            except Exception:
                break

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
        if self.single_pst_path and (not suffix or suffix == "part1"):
            return self.single_pst_path
        if suffix:
            return os.path.join(self.out_dir, f"{self.base}_{suffix}.pst")
        return os.path.join(self.out_dir, f"{self.base}.pst")

    def _register_open(self, path: str, root) -> None:
        if path and root:
            self.open_roots[normcasepath(path)] = root

    def _detach_path(self, path: Optional[str]) -> None:
        if not path:
            return
        key = normcasepath(path)
        root = self.open_roots.pop(key, None)
        if root is not None:
            for _attempt in range(3):
                try:
                    self.ns.RemoveStore(root)
                    break
                except Exception:
                    time.sleep(0.2)

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
        suffix = "" if self.cur["part"] == 1 else f"part{self.cur['part']}"
        desired = self._new_path(suffix)
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
            desired = self._new_path(f"{year}")
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
    ap.add_argument("--split-pst", "--max-pst-size", dest="split_pst", help="Max size per PST before rotating (e.g. 20GB, 15G, 500MB)")
    ap.add_argument("--max-pst-gb", type=float, default=15.0, help="Max ~GB per PST before rotating part2/part3 (default: 15.0, 0=unlimited)")
    ap.add_argument("--pst-root", default=None, help="Top folder name inside each PST (default: root for single-PST, 'Imported (EML)' for --out-dir)")
    ap.add_argument("--flatten", action="store_true", help="Do not replicate source subfolder structure inside PST")
    ap.add_argument("--since", "--after", dest="since", help="Filter messages on or after date/time (e.g. 2023-01-01 or 2023-01-01T00:00:00)")
    ap.add_argument("--until", "--before", dest="until", help="Filter messages on or before date/time (e.g. 2023-12-31 or 2023-12-31T23:59:59)")
    ap.add_argument("--include-folder", dest="include_folders", action="append", help="Include only folders matching name/glob (repeatable)")
    ap.add_argument("--exclude-folder", dest="exclude_folders", action="append", help="Exclude folders matching name/glob (repeatable)")
    ap.add_argument("--start-year", type=int, default=None, help="Import messages from this year (inclusive)")
    ap.add_argument("--end-year", type=int, default=None, help="Import messages up to this year (inclusive)")
    ap.add_argument("--overwrite-pst", action="store_true", help="Overwrite existing target PST files if present")
    ap.add_argument("--include-deleted", action="store_true", help="Import messages marked expunged/deleted")
    ap.add_argument("--flush-every", type=int, default=5000, help="Detach/reattach PST every N items to release COM memory (default: 5000, 0=off)")
    ap.add_argument("--count-every", type=int, default=200, help="Print folder Items.Count every N items (0=off)")
    ap.add_argument(
        "--dedup",
        choices=["auto", "message-id", "sha256", "none"],
        nargs="?",
        const="auto",
        default="none",
        help="Deduplicate messages by Message-ID, SHA-256 content hash, or auto (default: none)",
    )
    ap.add_argument(
        "--resume",
        nargs="?",
        const=True,
        default=None,
        help="Resume an interrupted migration by skipping messages already marked SUCCESS in the audit log or specified manifest file",
    )
    ap.add_argument(
        "--state-file",
        dest="state_file",
        help="Path to persistent SQLite or JSONL state database for incremental / delta cutover sync across runs",
    )
    ap.add_argument(
        "--no-gmail-labels",
        action="store_true",
        help="Disable automatic Google Takeout X-Gmail-Labels folder hierarchy reconstruction",
    )
    ap.add_argument("--audit-log", help="Path to write structured audit manifest (.csv or .jsonl)")
    ap.add_argument("--error-log", help="Path to write structured failure report (.jsonl)")
    ap.add_argument("--log-file", help="Path to write persistent operational log file")
    ap.add_argument("--dry-run", action="store_true", help="Perform pre-flight inspection and report without modifying disk/PST")
    ap.add_argument("-v", "--verbose", action="store_true", help="Enable verbose logging")
    ap.add_argument("-q", "--quiet", action="store_true", help="Suppress progress output, show only summary and errors")
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

    src = args.src_opt or args.src_pos
    if not src:
        ap.error("A source path (positional or --src) is required.")

    src = os.path.normpath(src)
    if not os.path.exists(win_long_path(src)):
        print(f"Source path not found: {src}", file=sys.stderr)
        sys.exit(1)

    if args.dry_run:
        perform_dry_run(src)
        return

    if not HAS_PYWIN32:
        print("Error: pywin32 is not installed. Run:  pip install pywin32", file=sys.stderr)
        sys.exit(1)

    check_or_lift_outlook_attachment_limit()

    target = args.out_pst or args.out_dir or args.out_pos
    if not target:
        ap.error("Both a source path (positional or --src) and an output target (positional, --out, or --out-dir) are required.")

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

    max_pst_gb = args.max_pst_gb
    if args.split_pst:
        max_pst_gb = parse_size_bytes(args.split_pst) / (1024**3)

    date_filter = DateFilter(since=args.since, until=args.until, start_year=args.start_year, end_year=args.end_year)
    folder_filter = FolderFilter(include_patterns=args.include_folders, exclude_patterns=args.exclude_folders)
    state = MigrationState(args.state_file) if args.state_file else None

    logger = AuditLogger(
        audit_log_path=args.audit_log,
        error_log_path=args.error_log,
        log_file_path=args.log_file,
        quiet=args.quiet,
        verbose=args.verbose,
    )

    app, ns = ensure_outlook_with_logon()

    router = PstRouter(
        ns=ns,
        out_dir=out_dir,
        base_name=base_name,
        max_gb=max_pst_gb,
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
    skipped_date = 0
    skipped_deleted = 0
    skipped_folder = 0
    skipped_state = 0
    done_bytes = 0
    tmpdir = tempfile.mkdtemp(prefix="mail2pst_")
    current_pst = None

    seen_keys: Set[str] = set()
    seen_mid_signatures: Dict[str, Tuple[str, str, str]] = {}
    skipped_duplicate = 0

    if args.resume:
        resume_source = args.resume if isinstance(args.resume, str) and os.path.isfile(args.resume) else args.audit_log
        if resume_source and os.path.isfile(win_long_path(resume_source)):
            existing = AuditLogger.load_completed_keys(resume_source)
            seen_keys.update(existing)
            logger.info(f"[*] Resuming: loaded {len(existing):,} previously completed message signatures from {resume_source}")
        elif args.resume is not True:
            logger.warn(f"Resume source manifest not found: {args.resume}")

    if (single_pst_path or args.max_pst_gb == 0) and total_bytes > 45 * (1024**3):
        logger.warn(
            f"Source mailbox ({total_bytes / (1024**3):.1f} GB) approaches or exceeds Outlook's default 50 GB single-PST limit.\n"
            f"       To prevent Outlook COM error 0x8004011D (PST file size limit), consider using --out-dir with --max-pst-gb 15.0 or --split-by year."
        )

    shutdown = ShutdownController(tool_name="mail2pst", logger=logger)
    shutdown.install()

    try:
        logger.info(f"Source: {src} | Estimated size: {total_bytes / 1e9:,.2f} GB")
        for rec in iter_mail_records(src, include_deleted=True, parse_gmail=not args.no_gmail_labels):
            if shutdown.is_interrupted:
                if not logger.quiet:
                    sys.stdout.write(f"\n[!] Stopping PST import gracefully: {shutdown.interruption_reason}\n")
                    sys.stdout.flush()
                break
            shutdown.check_pause()

            raw = rec.raw_bytes
            sz = len(raw)
            processed += 1
            done_bytes += sz
            folder_display = os.sep.join(rec.folder_parts) if rec.folder_parts else "Root"

            # 1. Folder Inclusion / Exclusion Rule Filtering
            if not folder_filter.matches(rec.folder_parts):
                skipped += 1
                skipped_folder += 1
                if logger and logger.audit_log_path:
                    logger.record_message(AuditRecord(
                        timestamp=datetime.now(timezone.utc).isoformat(),
                        status="SKIPPED_FOLDER",
                        source_file=rec.source_name,
                        source_folder=folder_display,
                        target_location="",
                        message_id="",
                        subject="",
                        date="",
                        sender="",
                        recipients="",
                        size_bytes=sz,
                        attachment_count=0,
                        notes=f"Filtered by folder rule ({folder_filter.describe()})",
                    ))
                continue

            try:
                msg = parser.parsebytes(raw)
            except Exception:
                msg = parser.parsebytes(b"Subject: (no subject)\r\n\r\n")

            mid = safe_get_header(msg, "Message-ID", "").strip()
            subj_val = safe_get_header(msg, "Subject", "(no subject)").strip()
            date_val = safe_get_header(msg, "Date", "").strip()

            # 2. Incremental Migration State Cache Check
            if state and state.is_active:
                sha_check = compute_sha256(raw) if not mid else ""
                if state.is_completed(mid, sha_check):
                    skipped += 1
                    skipped_state += 1
                    if logger and logger.audit_log_path:
                        logger.record_message(AuditRecord(
                            timestamp=datetime.now(timezone.utc).isoformat(),
                            status="SKIPPED_STATE_CACHE",
                            source_file=rec.source_name,
                            source_folder=folder_display,
                            target_location="",
                            message_id=mid,
                            subject=subj_val,
                            date=date_val,
                            sender=safe_get_header(msg, "From", ""),
                            recipients="; ".join(safe_get_all_headers(msg, "To")),
                            size_bytes=sz,
                            attachment_count=0,
                            notes="Previously migrated in delta state cache",
                        ))
                    continue

            is_dup = False
            dup_key = None
            sha = ""

            if args.dedup != "none" or args.resume:
                sha = compute_sha256(raw) if (args.dedup in ("sha256", "auto") or not mid or args.resume) else ""
                if args.dedup == "message-id":
                    dup_key = f"mid:{mid}" if mid else None
                    if dup_key and dup_key in seen_keys:
                        is_dup = True
                elif args.dedup == "sha256":
                    dup_key = f"sha256:{sha}" if sha else None
                    if dup_key and dup_key in seen_keys:
                        is_dup = True
                else:  # auto or resume
                    if args.resume and mid and f"mid:{mid}" in seen_keys:
                        is_dup = True
                        dup_key = f"mid:{mid}"
                    elif args.resume and sha and f"sha256:{sha}" in seen_keys:
                        is_dup = True
                        dup_key = f"sha256:{sha}"
                    elif args.dedup == "auto":
                        if mid and mid in seen_mid_signatures:
                            prev_sha, prev_subj, prev_date = seen_mid_signatures[mid]
                            if sha == prev_sha or (prev_date and prev_date == date_val) or (prev_subj and prev_subj == subj_val):
                                is_dup = True
                                dup_key = f"mid:{mid}"
                            # else: distinct message colliding on Message-ID, preserve it!
                        elif sha and f"sha256:{sha}" in seen_keys:
                            is_dup = True
                            dup_key = f"sha256:{sha}"
                        else:
                            dup_key = f"mid:{mid}" if mid else (f"sha256:{sha}" if sha else None)

            if is_dup:
                skipped += 1
                skipped_duplicate += 1
                if logger and logger.audit_log_path:
                    logger.record_message(AuditRecord(
                        timestamp=datetime.now(timezone.utc).isoformat(),
                        status="SKIPPED_DUPLICATE",
                        source_file=rec.source_name,
                        source_folder=folder_display,
                        target_location="",
                        message_id=mid,
                        subject=safe_get_header(msg, "Subject", "(no subject)"),
                        date=safe_get_header(msg, "Date", ""),
                        sender=safe_get_header(msg, "From", ""),
                        recipients="; ".join(safe_get_all_headers(msg, "To")),
                        size_bytes=sz,
                        attachment_count=0,
                        notes="Skipped duplicate message (dedup/resume)",
                    ))
                continue

            if dup_key:
                seen_keys.add(dup_key)
                if mid:
                    seen_mid_signatures[mid] = (sha, subj_val, date_val)
                    if args.dedup in ("sha256", "auto") or args.resume:
                        sha = compute_sha256(raw) if not sha else sha
                        seen_keys.add(f"mid:{mid}")
                        seen_keys.add(f"sha256:{sha}")

            flags = extract_message_flags(msg)
            if (rec.is_deleted or flags[4]) and not args.include_deleted:
                skipped += 1
                skipped_deleted += 1
                if logger and logger.audit_log_path:
                    logger.record_message(AuditRecord(
                        timestamp=datetime.now(timezone.utc).isoformat(),
                        status="SKIPPED_DELETED",
                        source_file=rec.source_name,
                        source_folder=folder_display,
                        target_location="",
                        message_id=safe_get_header(msg, "Message-ID", ""),
                        subject=safe_get_header(msg, "Subject", "(no subject)"),
                        date=safe_get_header(msg, "Date", ""),
                        sender=safe_get_header(msg, "From", ""),
                        recipients="; ".join(safe_get_all_headers(msg, "To")),
                        size_bytes=sz,
                        attachment_count=0,
                        notes="Message marked deleted/expunged",
                    ))
                continue

            # 3. Granular Date Range Filtering
            msg_dt = extract_best_datetime(msg, fallback_timestamp=getattr(rec, "file_mtime", None))
            year, _ = (msg_dt.year, msg_dt.month) if msg_dt else pick_year_month(msg, fallback_timestamp=getattr(rec, "file_mtime", None))

            if not date_filter.matches(msg_dt, fallback_year=year):
                skipped += 1
                skipped_date += 1
                if logger and logger.audit_log_path:
                    logger.record_message(AuditRecord(
                        timestamp=datetime.now(timezone.utc).isoformat(),
                        status="SKIPPED_DATE",
                        source_file=rec.source_name,
                        source_folder=folder_display,
                        target_location="",
                        message_id=safe_get_header(msg, "Message-ID", ""),
                        subject=safe_get_header(msg, "Subject", "(no subject)"),
                        date=safe_get_header(msg, "Date", ""),
                        sender=safe_get_header(msg, "From", ""),
                        recipients="; ".join(safe_get_all_headers(msg, "To")),
                        size_bytes=sz,
                        attachment_count=0,
                        notes=f"Filtered by date range ({date_filter.describe()})",
                    ))
                continue

            store, root, pst_path = router.route(sz, year)
            if pst_path != current_pst:
                current_pst = pst_path
                logger.info(f"\nCURRENT PST: {current_pst}")

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
                    try:
                        populate_mailitem(item, msg, raw, tmpdir, flags=flags, out_dir=args.out_dir)
                    except Exception as pop_err:
                        # Zero-Data-Loss Fallback Preservation:
                        # If complex MIME body/header formatting fails in Outlook COM, preserve the
                        # original message intact as an attached source file with original timestamps.
                        logger.warn(
                            f"Standard MailItem conversion failed for {rec.source_name} ({pop_err}); "
                            f"preserving raw message as fallback attachment."
                        )
                        try:
                            item.Delete()
                        except Exception:
                            pass
                        item = dest.Items.Add(0)
                        subj_hdr = safe_get_header(msg, "Subject", "")
                        clean_subj = subj_hdr or os.path.basename(rec.source_name)
                        item.Subject = f"[Preserved Mail] {clean_subj}"[:250]
                        item.Body = (
                            f"Notice: This message encountered an issue during standard MIME rendering in Outlook COM.\n"
                            f"Source: {rec.source_name}\n"
                            f"Error: {pop_err}\n\n"
                            f"The full, unmodified raw message has been preserved intact as an attachment.\n"
                        )
                        fallback_file = os.path.join(tmpdir, f"raw_msg_{processed}.eml")
                        with open(fallback_file, "wb") as f_fb:
                            f_fb.write(raw)
                        item.Attachments.Add(fallback_file, 1, 1, f"original_{processed}.eml")
                        pa_fb = item.PropertyAccessor
                        dt = extract_best_datetime(msg, fallback_timestamp=getattr(rec, "file_mtime", None))
                        if dt is not None:
                            try:
                                utc_dt = to_mapi_utc(pa_fb, dt)
                                set_prop(pa_fb, PR_MESSAGE_DELIVERY_TIME, utc_dt)
                                set_prop(pa_fb, PR_CLIENT_SUBMIT_TIME, utc_dt)
                                set_prop(pa_fb, PR_CREATION_TIME, utc_dt)
                                set_prop(pa_fb, PR_LAST_MODIFICATION_TIME, utc_dt)
                            except Exception:
                                pass
                        set_prop(pa_fb, PR_MESSAGE_FLAGS, 0)
                        item.Save()
                        try:
                            os.remove(fallback_file)
                        except OSError:
                            pass

                    try:
                        if item.Parent and item.Parent.EntryID != dest.EntryID:
                            moved = item.Move(dest)
                            moved.Save()
                            del moved
                    except Exception:
                        pass

                del item
                imported += 1
                if logger and logger.audit_log_path:
                    logger.record_message(AuditRecord(
                        timestamp=datetime.now(timezone.utc).isoformat(),
                        status="SUCCESS",
                        source_file=rec.source_name,
                        source_folder=folder_display,
                        target_location=f"{pst_path}::{rel_dir}",
                        message_id=safe_get_header(msg, "Message-ID", ""),
                        subject=safe_get_header(msg, "Subject", "(no subject)"),
                        date=safe_get_header(msg, "Date", ""),
                        sender=safe_get_header(msg, "From", ""),
                        recipients="; ".join(safe_get_all_headers(msg, "To")),
                        size_bytes=sz,
                        attachment_count=0,
                        sha256=compute_sha256(raw) if logger.audit_log_path else "",
                    ))

                if state and state.is_active:
                    state.record_success(
                        mid=mid,
                        sha256=sha if sha else compute_sha256(raw),
                        source_file=rec.source_name,
                        folder=folder_display,
                    )
            except Exception as e:
                # Zero-Data-Loss Emergency Fallback: preserve unimported message to disk
                unimported_dir = os.path.join(args.out_dir, "_unimported_messages")
                try:
                    os.makedirs(unimported_dir, exist_ok=True)
                    subj = sanitize_name(safe_get_header(msg, "Subject", "msg") or "msg", max_len=60)
                    fb_eml = os.path.join(unimported_dir, f"msg_{processed:06d}_{subj}.eml")
                    with open(win_long_path(fb_eml), "wb") as f_err:
                        f_err.write(raw)
                    logger.warn(f"Emergency preservation: saved unimported raw message to {fb_eml}")
                except Exception:
                    pass

                logger.error(f"Failed to save item from {rec.source_name}", exc=e)
                logger.record_failure(
                    source_file=rec.source_name,
                    source_folder=folder_display,
                    message_id=safe_get_header(msg, "Message-ID", ""),
                    subject=safe_get_header(msg, "Subject", "(no subject)"),
                    error_type=type(e).__name__,
                    error_details=str(e),
                )
                skipped += 1

            if processed % 500 == 0:
                if pythoncom is not None:
                    try:
                        pythoncom.PumpWaitingMessages()
                        pythoncom.CoFreeUnusedLibraries()
                    except Exception:
                        pass
                gc.collect()

            if args.count_every and processed % args.count_every == 0:
                try:
                    logger.info(f"\n  Items in '{dest.Name}': {dest.Items.Count}")
                except Exception:
                    pass
                gc.collect()

            if args.flush_every and processed % args.flush_every == 0:
                try:
                    del dest
                    gc.collect()
                    store, root, pst_path = router.flush_pst(pst_path)
                    logger.info(f"\n[FLUSH] Re-attached PST: {pst_path}")
                except Exception as e:
                    logger.warn(f"Periodic flush failed: {e}")

            if processed % 100 == 0:
                pct = min(100.0, (done_bytes / max(total_bytes, 1)) * 100.0)
                elapsed = time.perf_counter() - start
                rate = done_bytes / max(elapsed, 1e-9)
                eta = max(0.0, (total_bytes - done_bytes)) / max(rate, 1e-9)
                h, m, s = int(eta // 3600), int(eta % 3600 // 60), int(eta % 60)
                if not logger.quiet:
                    sys.stdout.write(
                        f"\rProcessed: {processed:,} ({pct:6.2f}%) | "
                        f"ETA {h:02d}:{m:02d}:{s:02d} | PST: {current_pst}"
                    )
                    sys.stdout.flush()

        if processed == 0 and not shutdown.is_interrupted:
            print("No supported email messages found in source.", file=sys.stderr)
            sys.exit(1)

        if processed > 0 or not shutdown.is_interrupted:
            logger.print_reconciliation_summary(
                source=src,
                target=single_pst_path or out_dir,
                start_time=start,
                total_scanned=processed,
                exported=imported,
                skipped_date=skipped_date,
                skipped_deleted=skipped_deleted,
                failed=skipped - skipped_date - skipped_deleted - skipped_duplicate - skipped_folder - skipped_state,
                total_bytes=done_bytes,
                extra_details={"Target PST Location": single_pst_path or out_dir},
                skipped_duplicate=skipped_duplicate,
                skipped_folder=skipped_folder,
                skipped_state=skipped_state,
                interrupted=shutdown.is_interrupted,
                interruption_reason=shutdown.interruption_reason,
            )

        if shutdown.is_interrupted:
            shutdown.print_resumption_banner(
                resume_file=logger.audit_log_path if logger else None,
                state_file=args.state_file,
                tool="pst",
            )
            sys.exit(130)

    except KeyboardInterrupt:
        shutdown.request_stop("KeyboardInterrupt (Ctrl+C)")
        if processed > 0:
            logger.print_reconciliation_summary(
                source=src,
                target=single_pst_path or out_dir,
                start_time=start,
                total_scanned=processed,
                exported=imported,
                skipped_date=skipped_date,
                skipped_deleted=skipped_deleted,
                failed=skipped - skipped_date - skipped_deleted - skipped_duplicate - skipped_folder - skipped_state,
                total_bytes=done_bytes,
                extra_details={"Target PST Location": single_pst_path or out_dir},
                skipped_duplicate=skipped_duplicate,
                skipped_folder=skipped_folder,
                skipped_state=skipped_state,
                interrupted=True,
                interruption_reason=shutdown.interruption_reason,
            )
        shutdown.print_resumption_banner(
            resume_file=logger.audit_log_path if logger else None,
            state_file=args.state_file,
            tool="pst",
        )
        sys.exit(130)

    finally:
        shutdown.restore()
        logger.close()
        shutil.rmtree(tmpdir, ignore_errors=True)
        if state:
            state.close()
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
