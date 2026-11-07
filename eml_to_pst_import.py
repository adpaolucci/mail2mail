# eml_to_pst_import.py
# -----------------------------------------------------------------------------
# Import .EML files into Outlook PST files via MAPI/COM (pywin32), with:
#   - split by YEAR (one PST per year) OR even-split into N PSTs by bytes
#   - max PST size limit (e.g., 15–20 GB) -> auto-rotate to part2, part3...
#   - periodic flush (detach/reattach PST) so Explorer shows file growth
#   - live folder item counts for progress validation
#   - subfolder hierarchy preservation (mirrors relative folder tree inside PST)
#   - preservation of Read/Unread, Flagged, and Replied/Forwarded states
#
# WHY this version works (Drafts fix):
#   We create each MailItem **directly in the PST target folder** using
#   `folder.Items.Add(0)` (0 = olMailItem). This avoids Outlook saving new
#   items into the default profile Drafts folder (which happens with CreateItem).
#
# Tested on: Windows + Outlook (M365 / 2016+). Requires Outlook installed.
#
# Dependencies (install on CMD/PowerShell):
#   pip install pywin32
#
# Outlook prerequisites:
#   1) Have Outlook installed and a default mail profile configured.
#   2) Put Outlook in "Work Offline" while importing (recommended).
#   3) Do NOT keep the target PST open in other apps.
#   4) If you have multiple profiles, ensure the default one opens without
#      prompts; this script calls `ns.Logon("", "", False, True)`.
#
# Usage examples (PowerShell / CMD):
#   # Year-based, limit ~15 GB per PST, flush every 5000 items, progress every 200
#   python eml_to_pst_import.py ^
#     --src "D:\Export_EML" ^
#     --out-dir "D:\PSTs" ^
#     --base-name emails ^
#     --split-by year ^
#     --max-pst-gb 15 ^
#     --pst-root "Imported (EML)" ^
#     --flush-every 5000 ^
#     --count-every 200
#
#   # Even-split into 6 PSTs by total bytes (no per-year folders):
#   python eml_to_pst_import.py ^
#     --src "D:\Export_EML" ^
#     --out-dir "D:\PSTs" ^
#     --base-name emails ^
#     --splits 6 ^
#     --max-pst-gb 18 ^
#     --pst-root "Imported (EML)"
# -----------------------------------------------------------------------------

import argparse
import binascii
import gc
import mimetypes
import os
import re
import shutil
import sys
import time
from datetime import timezone
from email import policy
from email.header import decode_header, make_header
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime

try:
    import pythoncom
    import pywintypes
    import win32com.client as win32
except Exception:
    print("Error: pywin32 is not installed. Run:  pip install pywin32", file=sys.stderr)
    sys.exit(1)

# --- MAPI property tags (Unicode 0x001F preferred, ANSI 0x001E fallback) -----
PR_TRANSPORT_MESSAGE_HEADERS_W  = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
PR_TRANSPORT_MESSAGE_HEADERS_A  = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
PR_MESSAGE_DELIVERY_TIME        = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
PR_CLIENT_SUBMIT_TIME           = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
PR_INTERNET_MESSAGE_ID_W        = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_INTERNET_MESSAGE_ID_A        = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
PR_SENDER_NAME_W                = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
PR_SENDER_EMAIL_ADDRESS_W       = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"
PR_SENDER_ADDRTYPE_W            = "http://schemas.microsoft.com/mapi/proptag/0x0C1E001F"
PR_SENT_REPRESENTING_NAME_W     = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
PR_SENT_REPRESENTING_EMAIL_W    = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"
PR_SENT_REPRESENTING_ADDRTYPE_W = "http://schemas.microsoft.com/mapi/proptag/0x0064001F"
PR_ATTACH_CONTENT_ID_W          = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_CONTENT_ID_A          = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
PR_ATTACH_FLAGS                 = "http://schemas.microsoft.com/mapi/proptag/0x37140003"
PR_ICON_INDEX                   = "http://schemas.microsoft.com/mapi/proptag/0x10800003"
PR_LAST_VERB_EXECUTED           = "http://schemas.microsoft.com/mapi/proptag/0x10810003"

# Thunderbird X-Mozilla-Status bitmask flags (nsMsgMessageFlags.h)
MSG_FLAG_READ      = 0x0001
MSG_FLAG_REPLIED   = 0x0002
MSG_FLAG_MARKED    = 0x0004  # Flagged / Starred
MSG_FLAG_EXPUNGED  = 0x0008  # Deleted / Expunged
MSG_FLAG_FORWARDED = 0x1000  # Forwarded

FORBIDDEN_CHARS = r'\/:*?"<>|'
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

# Transient COM RPC errors when Outlook is busy
TRANSIENT_COM_HRESULTS = {
    0x80010001,  # RPC_E_CALL_REJECTED
    0x800706BA,  # RPC_S_SERVER_UNAVAILABLE
    0x8001010A,  # RPC_E_SERVERCALL_RETRYLATER
}


def win_long_path(path: str) -> str:
    """Add \\\\?\\ long-path prefix on Windows when path approaches MAX_PATH (260 chars)."""
    abspath = os.path.abspath(path)
    if os.name != "nt" or abspath.startswith("\\\\?\\") or len(abspath) < 240:
        return path
    if abspath.startswith("\\\\"):
        return "\\\\?\\UNC\\" + abspath[2:]
    return "\\\\?\\" + abspath


def com_retry(max_attempts=4, base_delay=0.35):
    """Retry transient Outlook COM RPC_E_CALL_REJECTED errors with message pumping."""
    def decorator(func):
        def wrapper(*args, **kwargs):
            last_err = None
            for attempt in range(1, max_attempts + 1):
                try:
                    pythoncom.PumpWaitingMessages()
                    return func(*args, **kwargs)
                except pywintypes.com_error as e:
                    last_err = e
                    hr = getattr(e, "hresult", 0) & 0xFFFFFFFF
                    if hr in TRANSIENT_COM_HRESULTS and attempt < max_attempts:
                        time.sleep(base_delay * (2 ** (attempt - 1)))
                        pythoncom.PumpWaitingMessages()
                        continue
                    raise
            raise last_err
        return wrapper
    return decorator


def set_prop(pa, tag, value):
    """Safely set a MAPI property; swallow COM-specific errors."""
    try:
        pa.SetProperty(tag, value)
        return True
    except Exception:
        return False


def to_mapi_utc(pa, dt):
    """
    Convert a parsed datetime into a naive UTC datetime for MAPI PT_SYSTIME.
    Note: pywin32 converts timezone-aware datetimes into local system time when
    constructing VT_DATE, so stripping tzinfo after converting to UTC ensures
    PropertyAccessor.SetProperty receives exact UTC wall-clock values.
    """
    try:
        if dt.tzinfo is not None:
            return dt.astimezone(timezone.utc).replace(tzinfo=None)
        return pa.LocalTimeToUTC(dt)
    except Exception:
        try:
            return dt.replace(tzinfo=None)
        except Exception:
            return dt


def safe_get_header(msg, name: str, default: str = "") -> str:
    """Safely read a single header from msg without failing on malformed RFC 5322 syntax."""
    try:
        val = msg.get(name)
        if val is not None:
            return str(val)
    except Exception:
        pass
    try:
        target = name.lower()
        for k, raw_v in msg.raw_items():
            if k.lower() == target:
                try:
                    return str(make_header(decode_header(str(raw_v))))
                except Exception:
                    return str(raw_v)
    except Exception:
        pass
    return default


def safe_get_all_headers(msg, name: str):
    """Safely read all occurrences of a header from msg as strings."""
    try:
        vals = msg.get_all(name)
        if vals:
            return [str(v) for v in vals if v is not None]
    except Exception:
        pass
    results = []
    try:
        target = name.lower()
        for k, raw_v in msg.raw_items():
            if k.lower() == target and raw_v is not None:
                try:
                    results.append(str(make_header(decode_header(str(raw_v)))))
                except Exception:
                    results.append(str(raw_v))
    except Exception:
        pass
    return results


def decode_modified_utf7(name):
    """Decode IMAP Modified UTF-7 folder names if present."""
    if '&' not in name:
        return name
    res = []
    parts = name.split('&')
    res.append(parts[0])
    for part in parts[1:]:
        if '-' in part:
            b64, rest = part.split('-', 1)
            if b64 == '':
                res.append('&')
            else:
                b64_std = b64.replace(',', '/')
                pad = len(b64_std) % 4
                if pad:
                    b64_std += '=' * (4 - pad)
                try:
                    res.append(binascii.a2b_base64(b64_std).decode('utf-16-be'))
                except Exception:
                    res.append('&' + b64 + '-')
            res.append(rest)
        else:
            res.append('&' + part)
    return "".join(res)


def sanitize_name(name, default="Unnamed", max_len=120):
    """Sanitize a folder or attachment name for Windows and Outlook MAPI."""
    decoded = decode_modified_utf7(str(name or ""))
    clean = "".join(
        '_' if (c in FORBIDDEN_CHARS or ord(c) < 32 or ord(c) == 127) else c
        for c in decoded
    )
    clean = re.sub(r'\s+', ' ', clean).strip(' .')
    if not clean:
        clean = default
    if clean.upper().split('.')[0] in WINDOWS_RESERVED_NAMES:
        clean = f"_{clean}"
    if len(clean) > max_len:
        base, ext = os.path.splitext(clean)
        ext = ext[:20]
        clean = base[: max(1, max_len - len(ext))].strip(' .') + ext
    return clean or default


def addresses_to_str(values):
    """Normalize RFC822 addresses into Outlook-friendly 'Name <addr>; ...' string."""
    if not values:
        return ""
    if isinstance(values, str):
        values = [values]
    parts = []
    for name, addr in getaddresses(values):
        name = (name or "").strip().replace(';', ',')
        addr = (addr or "").strip()
        if addr and name:
            parts.append(f"{name} <{addr}>")
        elif addr:
            parts.append(addr)
        elif name:
            parts.append(name)
    return "; ".join(parts)


def extract_message_flags(msg):
    """
    Extract (is_read, is_flagged, is_replied, is_forwarded, is_expunged)
    from Status, X-Status, X-Message-Flag, and Thunderbird X-Mozilla-Status headers.
    """
    is_read = None
    is_flagged = False
    is_replied = False
    is_forwarded = False
    is_expunged = False

    # 1. Standard Status header (e.g., 'RO' = read, 'O' = unread)
    status_hdr = safe_get_header(msg, "Status", "").upper()
    if status_hdr:
        is_read = ("R" in status_hdr)

    # 2. Standard X-Status header ('F' = flagged, 'A' = answered/replied, 'W' = forwarded)
    x_status = safe_get_header(msg, "X-Status", "").upper()
    if "F" in x_status:
        is_flagged = True
    if "A" in x_status:
        is_replied = True
    if "W" in x_status:
        is_forwarded = True

    # 3. X-Message-Flag header
    if safe_get_header(msg, "X-Message-Flag", "").strip():
        is_flagged = True

    # 4. Thunderbird X-Mozilla-Status bitmask
    tb_status = safe_get_header(msg, "X-Mozilla-Status", "").strip()
    if tb_status:
        try:
            val = int(tb_status, 16)
            if is_read is None:
                is_read = bool(val & MSG_FLAG_READ)
            if val & MSG_FLAG_MARKED:
                is_flagged = True
            if val & MSG_FLAG_REPLIED:
                is_replied = True
            if val & MSG_FLAG_FORWARDED:
                is_forwarded = True
            if val & MSG_FLAG_EXPUNGED:
                is_expunged = True
        except ValueError:
            pass

    # Default to read if no status headers exist, or unread if explicitly Status: O
    if is_read is None:
        is_read = True

    return is_read, is_flagged, is_replied, is_forwarded, is_expunged


def pick_body(msg):
    """
    Pick HTML if available; otherwise text/plain; ignore attachment bodies.
    Returns (html, text, body_parts_set) so add_attachments does not duplicate the body.
    """
    html = None
    text = None
    body_parts = set()

    # Prefer EmailMessage.get_body when available
    if hasattr(msg, "get_body"):
        try:
            html_part = msg.get_body(preferencelist=("html",))
            if html_part is not None:
                body_parts.add(id(html_part))
                html = html_part.get_content()
        except Exception:
            pass
        try:
            text_part = msg.get_body(preferencelist=("plain",))
            if text_part is not None:
                body_parts.add(id(text_part))
                text = text_part.get_content()
        except Exception:
            pass
        if html is not None or text is not None:
            return html, text, body_parts

    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_maintype() == "message":
                # Do not descend into attached message/rfc822 for the parent body
                continue
            ctype = part.get_content_type()
            disp = safe_get_header(part, "Content-Disposition", "").lower()
            try:
                has_fn = bool(part.get_filename())
            except Exception:
                has_fn = False
            if "attachment" in disp or has_fn:
                continue
            if ctype == "text/html" and html is None:
                body_parts.add(id(part))
                try:
                    html = part.get_content()
                except Exception:
                    try:
                        raw_p = part.get_payload(decode=True)
                        if raw_p is not None:
                            html = raw_p.decode(part.get_content_charset() or "utf-8", errors="replace")
                    except Exception:
                        pass
            elif ctype == "text/plain" and text is None:
                body_parts.add(id(part))
                try:
                    text = part.get_content()
                except Exception:
                    try:
                        raw_p = part.get_payload(decode=True)
                        if raw_p is not None:
                            text = raw_p.decode(part.get_content_charset() or "utf-8", errors="replace")
                    except Exception:
                        pass
    else:
        ctype = msg.get_content_type()
        if ctype == "text/html":
            body_parts.add(id(msg))
            try:
                html = msg.get_content()
            except Exception:
                try:
                    raw_p = msg.get_payload(decode=True)
                    if raw_p is not None:
                        html = raw_p.decode(msg.get_content_charset() or "utf-8", errors="replace")
                except Exception:
                    pass
        elif ctype.startswith("text/"):
            body_parts.add(id(msg))
            try:
                text = msg.get_content()
            except Exception:
                try:
                    raw_p = msg.get_payload(decode=True)
                    if raw_p is not None:
                        text = raw_p.decode(msg.get_content_charset() or "utf-8", errors="replace")
                except Exception:
                    pass
    return html, text, body_parts


def _iter_parts_skipping_rfc822_children(msg):
    """Walk MIME tree but treat message/rfc822 parts as leaf attachments rather than descending."""
    if not msg.is_multipart():
        yield msg
        return
    for part in msg.get_payload() or []:
        ctype = part.get_content_type()
        if ctype == "message/rfc822":
            yield part
        elif part.is_multipart():
            yield from _iter_parts_skipping_rfc822_children(part)
        else:
            yield part


def add_attachments(mail, msg, tmpdir, body_parts=None):
    """Save each attachment or inline CID asset to a clean temp path and add via Outlook Attachments.Add."""
    if body_parts is None:
        body_parts = set()
    count = 0
    has_cid = False

    for part in _iter_parts_skipping_rfc822_children(msg):
        if id(part) in body_parts:
            continue

        ctype = part.get_content_type()
        disp = safe_get_header(part, "Content-Disposition", "").lower()
        cid = safe_get_header(part, "Content-ID", "") or None
        try:
            fn = part.get_filename()
        except Exception:
            fn = None

        is_rfc822 = (ctype == "message/rfc822")
        is_explicit_att = ("attachment" in disp) or (fn is not None) or is_rfc822
        is_inline_asset = (cid is not None) or (
            "inline" in disp and ctype not in ("text/plain", "text/html")
        )

        if not (is_explicit_att or is_inline_asset):
            continue

        payload = None
        if is_rfc822:
            try:
                sub = part.get_payload()
                if isinstance(sub, list) and sub:
                    payload = sub[0].as_bytes()
                else:
                    payload = part.as_bytes()
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
            if is_rfc822:
                fn = f"forwarded_message_{count + 1}.eml"
            else:
                ext = mimetypes.guess_extension(ctype or "") or ".bin"
                fn = f"attachment_{count + 1}{ext}"

        slot_dir = os.path.abspath(os.path.join(tmpdir, f"s_{count}"))
        os.makedirs(slot_dir, exist_ok=True)
        # Keep total path length safely under Windows MAX_PATH (260 chars) for Outlook COM
        max_fn_len = max(24, min(120, 230 - len(slot_dir)))
        safe_fn = sanitize_name(fn, default=f"attachment_{count + 1}.bin", max_len=max_fn_len)
        p = os.path.join(slot_dir, safe_fn)

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


def build_headers_text(raw_bytes):
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


@com_retry()
def create_mail_in_dest(dest_folder, msg, raw_headers_text, tmpdir):
    """
    Create MailItem directly in the target PST folder (avoids default Drafts),
    populate OOM fields, attachments, body, and MAPI properties.
    """
    mail = dest_folder.Items.Add(0)  # 0 = olMailItem
    pa   = mail.PropertyAccessor

    # 1. Basic OOM fields
    mail.Subject = safe_get_header(msg, "Subject", "")
    mail.To  = addresses_to_str(safe_get_all_headers(msg, "To"))
    mail.CC  = addresses_to_str(safe_get_all_headers(msg, "Cc"))
    mail.BCC = addresses_to_str(safe_get_all_headers(msg, "Bcc"))

    # 2. Extract body & add attachments before setting HTMLBody so inline CIDs link cleanly
    html, text, body_parts = pick_body(msg)
    add_attachments(mail, msg, tmpdir, body_parts=body_parts)

    if html is not None:
        mail.HTMLBody = html
    elif text is not None:
        mail.Body = text
    else:
        mail.Body = ""

    # 3. Read / Flagged / Replied / Forwarded states
    is_read, is_flagged, is_replied, is_forwarded, _ = extract_message_flags(msg)
    try:
        mail.UnRead = not is_read
    except Exception:
        pass

    if is_flagged:
        try:
            mail.FlagStatus = 2  # olFlagMarked
            mail.FlagRequest = "Follow up"
        except Exception:
            pass

    # 4. Low-level MAPI properties set after Body/HTMLBody so OOM body sync does not reset them
    if raw_headers_text:
        if not set_prop(pa, PR_TRANSPORT_MESSAGE_HEADERS_W, raw_headers_text):
            set_prop(pa, PR_TRANSPORT_MESSAGE_HEADERS_A, raw_headers_text)

    from_list = getaddresses(safe_get_all_headers(msg, "From"))
    if from_list:
        name, addr = (from_list[0][0] or "").strip(), (from_list[0][1] or "").strip()
        display = name or addr or ""
        set_prop(pa, PR_SENDER_NAME_W, display)
        set_prop(pa, PR_SENDER_EMAIL_ADDRESS_W, addr or display)
        set_prop(pa, PR_SENDER_ADDRTYPE_W, "SMTP")
        set_prop(pa, PR_SENT_REPRESENTING_NAME_W, display)
        set_prop(pa, PR_SENT_REPRESENTING_EMAIL_W, addr or display)
        set_prop(pa, PR_SENT_REPRESENTING_ADDRTYPE_W, "SMTP")
        try:
            mail.SentOnBehalfOfName = display
        except Exception:
            pass

    mid = safe_get_header(msg, "Message-ID", "").strip()
    if mid:
        if not set_prop(pa, PR_INTERNET_MESSAGE_ID_W, mid):
            set_prop(pa, PR_INTERNET_MESSAGE_ID_A, mid)

    sent_dt = None
    date_str = safe_get_header(msg, "Date", "").strip()
    if date_str:
        try:
            sent_dt = parsedate_to_datetime(date_str)
        except Exception:
            pass
    if sent_dt:
        utc_dt = to_mapi_utc(pa, sent_dt)
        set_prop(pa, PR_MESSAGE_DELIVERY_TIME, utc_dt)
        set_prop(pa, PR_CLIENT_SUBMIT_TIME, utc_dt)

    if is_replied:
        set_prop(pa, PR_LAST_VERB_EXECUTED, 102)  # EXCHIVERB_REPLYTOSENDER
        set_prop(pa, PR_ICON_INDEX, 261)          # Mail replied icon
    elif is_forwarded:
        set_prop(pa, PR_LAST_VERB_EXECUTED, 104)  # EXCHIVERB_FORWARD
        set_prop(pa, PR_ICON_INDEX, 262)          # Mail forwarded icon

    return mail


def ensure_outlook_with_logon():
    """Start Outlook.Application and log on to the default MAPI profile."""
    pythoncom.CoInitialize()
    app = win32.Dispatch("Outlook.Application")
    ns  = app.GetNamespace("MAPI")
    ns.Logon("", "", False, True)  # default profile (no UI)
    return app, ns


def normcasepath(p):
    return os.path.normcase(os.path.normpath(p))


def find_store_by_path(ns, pst_path):
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


def create_or_attach_pst(ns, desired_path, overwrite=False):
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

    # Find the Store that matches the actual path we requested
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

    # Fallback: if Outlook created the PST in a default location, return the last added
    last = None
    for s in ns.Stores:
        last = s
    if last:
        return last, last.GetRootFolder(), last.FilePath

    raise RuntimeError(f"Failed to create/attach PST: {desired_path}")


@com_retry()
def get_or_create_folder_path(root_folder, pst_root_name, rel_dir):
    """Navigate or create pst_root_name and any relative subfolder segments inside root_folder."""
    curr = root_folder
    if pst_root_name:
        safe_root = sanitize_name(pst_root_name, default="Imported (EML)")
        try:
            curr = curr.Folders.Item(safe_root)
        except Exception:
            curr = curr.Folders.Add(safe_root)

    if rel_dir and rel_dir != ".":
        for part in rel_dir.split(os.sep):
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
      - by year: one PST per year (auto-rotate if > max size)
      - by even split: target_bytes computed from total size / splits
      - always tracks approximate bytes to trigger rotation
    """
    def __init__(self, ns, out_dir, base_name, max_gb, splits, split_by_year, overwrite=False):
        self.ns = ns
        self.out_dir = os.path.normpath(out_dir)
        self.base = base_name
        self.max_bytes = int(max_gb * (1024**3))
        self.splits = splits
        self.by_year = split_by_year
        self.overwrite = overwrite
        self.cur = {"store": None, "root": None, "path": None, "bytes": 0, "part": 0}
        self.year = {}
        self.target_bytes = None
        self.open_roots = {}  # normcasepath -> root_folder

    def set_total_bytes(self, total_bytes):
        if self.splits and self.splits > 0:
            self.target_bytes = max(1, total_bytes // self.splits)

    def _new_path(self, suffix):
        return os.path.join(self.out_dir, f"{self.base}_{suffix}.pst")

    def _register_open(self, path, root):
        if path and root:
            self.open_roots[normcasepath(path)] = root

    def _detach_path(self, path):
        if not path:
            return
        key = normcasepath(path)
        root = self.open_roots.pop(key, None)
        if root is not None:
            try:
                self.ns.RemoveStore(root)
            except Exception:
                pass

    def _initial_existing_bytes(self, desired_path):
        if not self.overwrite and os.path.isfile(desired_path):
            try:
                return os.path.getsize(desired_path)
            except OSError:
                return 0
        return 0

    def _rotate_new(self):
        # Detach previous sequential part so Outlook flushes it immediately
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

    def _need_rotate_seq(self, incoming):
        if self.cur["store"] and self.cur["bytes"] > 0 and (self.cur["bytes"] + incoming > self.max_bytes):
            return True
        if self.target_bytes and self.cur["store"] and (self.cur["bytes"] >= self.target_bytes):
            return True
        return False

    def _ensure_year(self, year, incoming):
        s = self.year.get(year)
        if s is None:
            desired = self._new_path(f"{year}_part1")
            init_bytes = self._initial_existing_bytes(desired)
            store, root, actual = create_or_attach_pst(self.ns, desired, overwrite=self.overwrite)
            s = {"path": actual, "store": store, "root": root, "bytes": init_bytes, "part": 1}
            self.year[year] = s
            self._register_open(actual, root)
            print(f"\nUSING PST ({year}): {actual}")
        elif s["bytes"] > 0 and (s["bytes"] + incoming > self.max_bytes):
            self._detach_path(s.get("path"))
            s["part"] += 1
            desired = self._new_path(f"{year}_part{s['part']}")
            init_bytes = self._initial_existing_bytes(desired)
            store, root, actual = create_or_attach_pst(self.ns, desired, overwrite=self.overwrite)
            s.update({"path": actual, "store": store, "root": root, "bytes": init_bytes})
            self._register_open(actual, root)
            print(f"\nROTATED PST ({year}) -> {actual}")
        return s

    def route(self, approx_size, year_hint=None):
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

    def flush_pst(self, pst_path):
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

    def close_all(self):
        """Detach all open PST stores tracked by this router."""
        for key, root in list(self.open_roots.items()):
            try:
                self.ns.RemoveStore(root)
            except Exception:
                pass
            self.open_roots.pop(key, None)


def list_eml_files(root_dir):
    """Return [(path, size)], total_bytes – recursively collects *.eml."""
    files = []
    total_bytes = 0
    for base, _, fnames in os.walk(win_long_path(root_dir)):
        clean_base = base[4:] if base.startswith("\\\\?\\") and not base.startswith("\\\\?\\UNC\\") else base
        for fn in fnames:
            if fn.lower().endswith(".eml"):
                p = os.path.join(clean_base, fn)
                try:
                    sz = os.path.getsize(win_long_path(p))
                except OSError:
                    sz = 0
                files.append((p, sz))
                total_bytes += sz
    files.sort()
    return files, total_bytes


def main():
    for _stream in (sys.stdout, sys.stderr):
        if hasattr(_stream, "reconfigure"):
            try:
                _stream.reconfigure(errors="replace")
            except Exception:
                pass
    ap = argparse.ArgumentParser(description="Import EML -> PST (Outlook/MAPI) with split & progress")
    ap.add_argument("--src", required=True, help="Root folder containing .eml files (recursively)")
    ap.add_argument("--out-dir", required=True, help="Directory to write .pst files")
    ap.add_argument("--base-name", default="emails", help="Base name for PST files")
    ap.add_argument("--split-by", choices=["year"], default=None, help="Split by mail year (one PST per year)")
    ap.add_argument("--splits", type=int, default=None, help="Even-split into N PSTs by total bytes")
    ap.add_argument("--max-pst-gb", type=float, default=15.0, help="Max ~GB per PST before rotating part2/part3...")
    ap.add_argument("--pst-root", default="Imported (EML)", help="Top folder name inside each PST (empty string for root)")
    ap.add_argument("--flatten", action="store_true", help="Do not replicate source subfolder structure inside PST")
    ap.add_argument("--overwrite-pst", action="store_true", help="Overwrite existing target PST files if present")
    ap.add_argument("--include-deleted", action="store_true", help="Import messages marked expunged (0x0008) in X-Mozilla-Status")
    ap.add_argument("--flush-every", type=int, default=0, help="Detach/reattach PST every N items (0=off)")
    ap.add_argument("--count-every", type=int, default=200, help="Print folder Items.Count every N items (0=off)")
    args = ap.parse_args()

    src = os.path.normpath(args.src)
    out_dir = os.path.normpath(args.out_dir)
    os.makedirs(out_dir, exist_ok=True)

    files, total_bytes = list_eml_files(src)
    if not files:
        print("No .eml files found.", file=sys.stderr)
        sys.exit(1)
    print(f"EML files: {len(files):,} | Total size: {total_bytes/1e9:,.2f} GB")

    app, ns = ensure_outlook_with_logon()

    router = PstRouter(
        ns=ns,
        out_dir=out_dir,
        base_name=args.base_name,
        max_gb=args.max_pst_gb,
        splits=args.splits,
        split_by_year=(args.split_by == "year"),
        overwrite=args.overwrite_pst,
    )
    router.set_total_bytes(total_bytes)

    parser = BytesParser(policy=policy.default)
    start = time.perf_counter()
    processed = 0
    imported = 0
    skipped = 0
    done_bytes = 0
    tmpdir = os.path.join(out_dir, f"_tmp_eml_to_pst_{os.getpid()}")
    os.makedirs(tmpdir, exist_ok=True)
    current_pst = None

    try:
        for (path, sz) in files:
            processed += 1
            done_bytes += sz

            # Read raw bytes & parse message
            try:
                with open(win_long_path(path), "rb") as f:
                    raw = f.read()
            except Exception as e:
                print(f"\n[WARN] Could not read {path}: {e}")
                skipped += 1
                continue

            try:
                msg = parser.parsebytes(raw)
            except Exception:
                msg = parser.parsebytes(b"Subject: (no subject)\r\n\r\n")

            _, _, _, _, is_expunged = extract_message_flags(msg)
            if is_expunged and not args.include_deleted:
                skipped += 1
                continue

            # Year hint (used when --split-by year)
            year = 1970
            try:
                date_str = safe_get_header(msg, "Date", "").strip()
                if date_str:
                    dt = parsedate_to_datetime(date_str)
                    if dt is not None and 1900 <= dt.year <= 2100:
                        year = dt.year
            except Exception:
                pass

            # Choose PST destination
            store, root, pst_path = router.route(sz or len(raw), year)
            if pst_path != current_pst:
                current_pst = pst_path
                print(f"\nCURRENT PST: {current_pst}")

            rel_dir = "." if args.flatten else os.path.relpath(os.path.dirname(path), src)
            dest = get_or_create_folder_path(root, args.pst_root, rel_dir)

            # Create item directly in destination folder, add attachments, save
            try:
                item = create_mail_in_dest(dest, msg, build_headers_text(raw), tmpdir)
                item.Save()

                # Safety: ensure the item actually resides in the target folder
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
                print(f"\n[WARN] Failed to save item from {path}: {e}")
                skipped += 1

            # Periodic live count
            if args.count_every and processed % args.count_every == 0:
                try:
                    print(f"\n  Items in '{dest.Name}': {dest.Items.Count}")
                except Exception:
                    pass
                gc.collect()

            # Periodic flush: close & reopen the current PST without deleting it
            if args.flush_every and processed % args.flush_every == 0:
                try:
                    del dest
                    gc.collect()
                    store, root, pst_path = router.flush_pst(pst_path)
                    print(f"\n[FLUSH] Re-attached PST: {pst_path}")
                except Exception as e:
                    print(f"\n[WARN] Periodic flush failed: {e}")

            # Console progress & ETA
            if processed % 100 == 0 or processed == len(files):
                pct = (done_bytes / max(total_bytes, 1)) * 100.0
                elapsed = time.perf_counter() - start
                rate = done_bytes / max(elapsed, 1e-9)
                eta = (total_bytes - done_bytes) / max(rate, 1e-9)
                h, m, s = int(eta // 3600), int(eta % 3600 // 60), int(eta % 60)
                print(
                    f"\r{processed:,}/{len(files):,} ({pct:7.3f}%) | "
                    f"ETA {h:02d}:{m:02d}:{s:02d} | PST: {current_pst}",
                    end="",
                )

        print("\nDone.")
        print(f"Processed: {processed:,} | Imported: {imported:,} | Skipped: {skipped:,}")
        print("PST(s) written to:", out_dir)

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
        try:
            router.close_all()
        except Exception as e:
            print("\n[WARN] Failed to close stores:", e)
        try:
            gc.collect()
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    main()
