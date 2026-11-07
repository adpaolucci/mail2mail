import argparse
import binascii
import email
import gc
import mimetypes
import os
import re
import shutil
import sys
import tempfile
from datetime import timezone
from email import policy
from email.utils import getaddresses, parsedate_to_datetime

try:
    import pythoncom
    import win32com.client
except Exception:
    print("Error: pywin32 is not installed. Run:  pip install pywin32", file=sys.stderr)
    sys.exit(1)

import mailbox

# Outlook MAPI invalid characters: \ / : * ? " < > |
FORBIDDEN_CHARS = r'\/:*?"<>|'
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

# MAPI Property Tags (Unicode 0x001F preferred, ANSI 0x001E fallback)
PR_TRANSPORT_MESSAGE_HEADERS_W = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
PR_TRANSPORT_MESSAGE_HEADERS_A = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
PR_MESSAGE_DELIVERY_TIME       = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
PR_CLIENT_SUBMIT_TIME          = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
PR_INTERNET_MESSAGE_ID_W       = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_INTERNET_MESSAGE_ID_A       = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
PR_ATTACH_CONTENT_ID_W         = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_CONTENT_ID_A         = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
PR_ATTACH_FLAGS                = "http://schemas.microsoft.com/mapi/proptag/0x37140003"
PR_SENDER_NAME_W               = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
PR_SENDER_EMAIL_W              = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"
PR_SENDER_ADDRTYPE_W           = "http://schemas.microsoft.com/mapi/proptag/0x0C1E001F"
PR_SENT_REPRESENTING_NAME_W    = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
PR_SENT_REPRESENTING_EMAIL_W   = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"
PR_SENT_REPRESENTING_ADDRTYPE_W = "http://schemas.microsoft.com/mapi/proptag/0x0064001F"
PR_ICON_INDEX                  = "http://schemas.microsoft.com/mapi/proptag/0x10800003"
PR_LAST_VERB_EXECUTED          = "http://schemas.microsoft.com/mapi/proptag/0x10810003"

# Thunderbird X-Mozilla-Status bitmask flags (nsMsgMessageFlags.h)
MSG_FLAG_READ      = 0x0001
MSG_FLAG_REPLIED   = 0x0002
MSG_FLAG_MARKED    = 0x0004  # Flagged / Starred
MSG_FLAG_EXPUNGED  = 0x0008  # Deleted (awaiting folder compaction)
MSG_FLAG_FORWARDED = 0x1000  # Forwarded


def set_prop(pa, tag, value):
    """Safely set a MAPI property via PropertyAccessor."""
    try:
        pa.SetProperty(tag, value)
        return True
    except Exception:
        return False


def to_mapi_utc(pa, dt):
    """Convert a parsed datetime into a UTC datetime suitable for MAPI PT_SYSTIME."""
    try:
        if dt.tzinfo is not None:
            return dt.astimezone(timezone.utc)
        return pa.LocalTimeToUTC(dt)
    except Exception:
        return dt


def decode_modified_utf7(name):
    """
    Decodes IMAP Modified UTF-7 (RFC 3501) folder names used by Thunderbird.
    Example: 'Notes &- Invoices' -> 'Notes & Invoices', 'Caf&AOk-' -> 'Café'
    """
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


def sanitize_folder_name(name, default="Unnamed_Folder", max_len=120):
    """Decodes IMAP Modified UTF-7 and sanitizes names for Outlook MAPI and Windows."""
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


def get_or_create_subfolder(parent_folder, folder_name):
    """Retrieves an existing MAPI subfolder or creates it if not found."""
    safe_name = sanitize_folder_name(folder_name)
    try:
        return parent_folder.Folders.Item(safe_name)
    except Exception:
        try:
            return parent_folder.Folders.Add(safe_name)
        except Exception as e:
            print(f"    [!] Could not create subfolder '{safe_name}': {e}")
            return parent_folder


def is_probable_mbox(file_path):
    """Checks if a file is an mbox file based on extension and magic bytes."""
    basename = os.path.basename(file_path)
    ignored_exts = (
        '.msf', '.dat', '.json', '.html', '.txt', '.ini',
        '.bak', '.tmp', '.sqlite', '.wdseml', '.mozmsgs'
    )
    if basename.startswith('.') or basename.lower().endswith(ignored_exts):
        return False

    try:
        if os.path.getsize(file_path) == 0:
            return True
        with open(file_path, 'rb') as f:
            head = f.read(1024).lstrip()
            return head.startswith(b"From ")
    except Exception:
        return False


def parse_message_flags(parsed_msg):
    """
    Parse Thunderbird X-Mozilla-Status and standard Status/X-Status headers.
    Returns (is_read, is_flagged, is_replied, is_forwarded, is_expunged).
    """
    is_read = None
    is_flagged = False
    is_replied = False
    is_forwarded = False
    is_expunged = False

    status_hdr = str(parsed_msg.get("Status", "") or "").upper()
    if status_hdr:
        is_read = ("R" in status_hdr)

    x_status = str(parsed_msg.get("X-Status", "") or "").upper()
    if "F" in x_status:
        is_flagged = True
    if "A" in x_status:
        is_replied = True
    if "W" in x_status:
        is_forwarded = True

    if str(parsed_msg.get("X-Message-Flag", "") or "").strip():
        is_flagged = True

    tb_status = str(parsed_msg.get("X-Mozilla-Status", "") or "").strip()
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

    if is_read is None:
        is_read = True

    return is_read, is_flagged, is_replied, is_forwarded, is_expunged


def build_headers_text(raw_bytes):
    """Extract raw RFC 5322 header block up to the first blank line (CRLF or LF)."""
    try:
        idx_crlf = raw_bytes.find(b"\r\n\r\n")
        idx_lf = raw_bytes.find(b"\n\n")
        candidates = [i for i in (idx_crlf, idx_lf) if i != -1]
        end = min(candidates) if candidates else min(len(raw_bytes), 1024 * 32)
        return raw_bytes[:end].decode("utf-8", errors="replace")
    except Exception:
        return ""


def add_recipients(mail_item, parsed_msg):
    """Adds To, CC, and BCC recipients to the Outlook MailItem."""
    def _fmt(hdr_name):
        headers = [str(h) for h in parsed_msg.get_all(hdr_name, [])]
        if not headers:
            return ""
        parts = []
        for realname, addr in getaddresses(headers):
            realname = (realname or "").strip()
            addr = (addr or "").strip()
            if realname and addr:
                parts.append(f"{realname} <{addr}>")
            elif addr:
                parts.append(addr)
            elif realname:
                parts.append(realname)
        return "; ".join(parts)

    mail_item.To = _fmt("to")
    mail_item.CC = _fmt("cc")
    mail_item.BCC = _fmt("bcc")


def handle_sender_info(mail_item, parsed_msg):
    """Maps From header using Unicode MAPI PropertyAccessor properties."""
    from_headers = [str(h) for h in parsed_msg.get_all('from', [])]
    if from_headers:
        senders = getaddresses(from_headers)
        if senders:
            realname, addr = (senders[0][0] or "").strip(), (senders[0][1] or "").strip()
            display = realname or addr or ""
            pa = mail_item.PropertyAccessor
            if display:
                set_prop(pa, PR_SENDER_NAME_W, display)
                set_prop(pa, PR_SENT_REPRESENTING_NAME_W, display)
            if addr or display:
                set_prop(pa, PR_SENDER_EMAIL_W, addr or display)
                set_prop(pa, PR_SENDER_ADDRTYPE_W, "SMTP")
                set_prop(pa, PR_SENT_REPRESENTING_EMAIL_W, addr or display)
                set_prop(pa, PR_SENT_REPRESENTING_ADDRTYPE_W, "SMTP")
            try:
                mail_item.SentOnBehalfOfName = display
            except Exception:
                pass


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


def populate_mailitem(mail_item, parsed, msg_bytes, temp_dir, flags=None):
    """
    Populates a native Outlook MailItem from parsed MIME message and raw bytes.
    Handles transport headers, UTC MAPI dates, sender/recipients, Thunderbird flags,
    HTML/Text body, inline CID images, regular attachments, and S/MIME.
    """
    pa = mail_item.PropertyAccessor

    # 1. Raw Internet Headers & Message-ID
    raw_headers_text = build_headers_text(msg_bytes)
    if raw_headers_text:
        if not set_prop(pa, PR_TRANSPORT_MESSAGE_HEADERS_W, raw_headers_text):
            set_prop(pa, PR_TRANSPORT_MESSAGE_HEADERS_A, raw_headers_text)

    mid = str(parsed.get("Message-ID", "") or "").strip()
    if mid:
        if not set_prop(pa, PR_INTERNET_MESSAGE_ID_W, mid):
            set_prop(pa, PR_INTERNET_MESSAGE_ID_A, mid)

    # 2. Subject
    mail_item.Subject = str(parsed.get('subject', '') or '')

    # 3. Sent / Delivery Date (MAPI PT_SYSTIME in UTC; MailItem.SentOn is read-only)
    try:
        date_hdr = parsed.get('date')
        if date_hdr:
            dt = parsedate_to_datetime(str(date_hdr))
            if dt is not None:
                utc_dt = to_mapi_utc(pa, dt)
                set_prop(pa, PR_MESSAGE_DELIVERY_TIME, utc_dt)
                set_prop(pa, PR_CLIENT_SUBMIT_TIME, utc_dt)
    except Exception:
        pass

    # 4. Sender and Recipients
    handle_sender_info(mail_item, parsed)
    add_recipients(mail_item, parsed)

    # 5. Read / Flagged / Replied / Forwarded Status
    if flags is None:
        flags = parse_message_flags(parsed)
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

    if is_replied:
        set_prop(pa, PR_LAST_VERB_EXECUTED, 102)
        set_prop(pa, PR_ICON_INDEX, 261)
    elif is_forwarded:
        set_prop(pa, PR_LAST_VERB_EXECUTED, 104)
        set_prop(pa, PR_ICON_INDEX, 262)

    # 6. S/MIME Detection (only opaque pkcs7-mime lacks a readable MIME body)
    content_type = parsed.get_content_type().lower()
    is_opaque_smime = ("pkcs7-mime" in content_type or "signed-data" in content_type)
    if is_opaque_smime:
        mail_item.MessageClass = "IPM.Note.SMIME"

    # 7. Body Extraction (track body_parts so inline bodies are not re-added as attachments)
    body_parts = set()
    html_body = None
    text_body = None

    if not is_opaque_smime:
        try:
            html_part = parsed.get_body(preferencelist=('html',))
            if html_part is not None:
                body_parts.add(id(html_part))
                html_body = html_part.get_content()
        except Exception:
            pass
        try:
            plain_part = parsed.get_body(preferencelist=('plain',))
            if plain_part is not None:
                body_parts.add(id(plain_part))
                text_body = plain_part.get_content()
        except Exception:
            pass

        if html_body is None and text_body is None and not parsed.is_multipart():
            if content_type.startswith("text/"):
                body_parts.add(id(parsed))
                try:
                    raw_p = parsed.get_payload(decode=True)
                    if raw_p is not None:
                        decoded_body = raw_p.decode(parsed.get_content_charset() or "utf-8", errors="replace")
                        if content_type == "text/html":
                            html_body = decoded_body
                        else:
                            text_body = decoded_body
                except Exception:
                    pass

    # 8. Extract Attachments and Inline CID Assets
    for att_idx, part in enumerate(_iter_parts_skipping_rfc822_children(parsed), 1):
        if id(part) in body_parts:
            continue

        part_ctype = part.get_content_type().lower()
        cdisp = str(part.get("Content-Disposition", "") or "").lower()
        cid = part.get("Content-ID", None)
        filename = part.get_filename()

        is_rfc822 = (part_ctype == "message/rfc822")
        is_attachment = ("attachment" in cdisp) or (filename is not None) or is_rfc822
        is_inline_cid = (cid is not None) or (
            "inline" in cdisp and part_ctype not in ("text/plain", "text/html")
        )

        if not (is_attachment or is_inline_cid or is_opaque_smime):
            continue

        att_bytes = None
        if is_rfc822:
            try:
                sub = part.get_payload()
                if isinstance(sub, list) and sub:
                    att_bytes = sub[0].as_bytes()
                else:
                    att_bytes = part.as_bytes()
            except Exception:
                att_bytes = None
        else:
            try:
                att_bytes = part.get_payload(decode=True)
            except Exception:
                att_bytes = None

        if att_bytes is None:
            continue

        if not filename:
            if is_opaque_smime and "pkcs7" in part_ctype:
                filename = "smime.p7m"
            elif is_rfc822:
                filename = f"forwarded_message_{att_idx}.eml"
            else:
                ext = mimetypes.guess_extension(part_ctype) or ".bin"
                filename = f"attachment_{att_idx}{ext}"

        safe_filename = sanitize_folder_name(filename, default=f"file_{att_idx}.bin", max_len=120)
        slot_dir = os.path.join(temp_dir, f"att_slot_{att_idx}")
        os.makedirs(slot_dir, exist_ok=True)
        temp_att_path = os.path.realpath(os.path.join(slot_dir, safe_filename))

        try:
            with open(temp_att_path, "wb") as f:
                f.write(att_bytes)

            # Type 1 = olByValue (embeds copy directly in PST using the file's real basename)
            att = mail_item.Attachments.Add(temp_att_path, 1, 1, safe_filename)

            # Map Content-ID for inline images
            if cid:
                clean_cid = str(cid).strip().strip("<>")
                if clean_cid:
                    att_pa = att.PropertyAccessor
                    if not set_prop(att_pa, PR_ATTACH_CONTENT_ID_W, clean_cid):
                        set_prop(att_pa, PR_ATTACH_CONTENT_ID_A, clean_cid)
                    # 4 = ATT_MHTML_REF (renders inline in HTMLBody)
                    set_prop(att_pa, PR_ATTACH_FLAGS, 4)

            del att
        except Exception as att_err:
            print(f"        [!] Warning: Failed to process attachment {safe_filename}: {att_err}")
        finally:
            shutil.rmtree(slot_dir, ignore_errors=True)

    # 9. Set Body after attachments so inline CID references resolve properly
    if html_body is not None:
        mail_item.HTMLBody = html_body
    elif text_body is not None:
        mail_item.Body = text_body
    elif not is_opaque_smime:
        mail_item.Body = ""

    # 10. Commit changes to PST store
    mail_item.Save()


def import_mbox_to_outlook(mbox_path, outlook_folder, temp_dir, include_deleted=False):
    """Processes an mbox file and adds messages directly into the destination Outlook folder."""
    if os.path.getsize(mbox_path) == 0:
        return

    print(f"[*] Processing: {os.path.basename(mbox_path)} -> [{outlook_folder.Name}]")

    try:
        mbox = mailbox.mbox(mbox_path, create=False)
    except Exception as e:
        print(f"    [!] Error opening mbox {mbox_path}: {e}")
        return

    count = 0
    skipped_deleted = 0

    try:
        for idx, msg in enumerate(mbox):
            try:
                raw_bytes = msg.as_bytes()
                if not raw_bytes:
                    continue

                parsed = email.message_from_bytes(raw_bytes, policy=policy.default)
                flags = parse_message_flags(parsed)
                if flags[4] and not include_deleted:
                    skipped_deleted += 1
                    continue

                pythoncom.PumpWaitingMessages()
                new_item = outlook_folder.Items.Add(0)  # 0 = olMailItem
                populate_mailitem(new_item, parsed, raw_bytes, temp_dir, flags=flags)
                del new_item

                count += 1
                if count % 100 == 0:
                    print(f"    -> {count} items imported...")
                    gc.collect()

            except Exception as err:
                print(f"    [!] Skipped message #{idx}: {err}")
                continue
    finally:
        mbox.close()

    summary = f"    [+] Finished: {count} messages imported into '{outlook_folder.Name}'"
    if skipped_deleted:
        summary += f" ({skipped_deleted} deleted/expunged skipped)"
    print(summary + ".")


def traverse_and_convert(current_dir, current_outlook_folder, temp_dir, include_deleted=False):
    """Recursively walks through Thunderbird mbox files and .sbd directories."""
    try:
        entries = sorted(os.listdir(current_dir))
    except Exception as e:
        print(f"[!] Access error reading {current_dir}: {e}")
        return

    # Process MBOX files
    for entry in entries:
        full_path = os.path.join(current_dir, entry)
        if os.path.isfile(full_path) and is_probable_mbox(full_path):
            target_folder = get_or_create_subfolder(current_outlook_folder, entry)
            import_mbox_to_outlook(full_path, target_folder, temp_dir, include_deleted=include_deleted)

    # Process .sbd container folders
    for entry in entries:
        full_path = os.path.join(current_dir, entry)
        if os.path.isdir(full_path) and entry.lower().endswith(".sbd"):
            parent_folder_name = entry[:-4]
            nested_outlook_folder = get_or_create_subfolder(current_outlook_folder, parent_folder_name)
            traverse_and_convert(full_path, nested_outlook_folder, temp_dir, include_deleted=include_deleted)


def main():
    for _stream in (sys.stdout, sys.stderr):
        if hasattr(_stream, "reconfigure"):
            try:
                _stream.reconfigure(errors="replace")
            except Exception:
                pass
    ap = argparse.ArgumentParser(
        description="Convert a Thunderbird ImapMail/Mail account directory directly into an Outlook PST file."
    )
    ap.add_argument("imap_src", help="Path to Thunderbird ImapMail account directory")
    ap.add_argument("pst_dest", help="Output .pst file path")
    ap.add_argument(
        "--include-deleted",
        action="store_true",
        help="Include messages marked as expunged/deleted (0x0008) in X-Mozilla-Status",
    )
    args = ap.parse_args()

    imap_src = os.path.abspath(args.imap_src)
    pst_dest = os.path.abspath(args.pst_dest)

    if not os.path.isdir(imap_src):
        print(f"[!] Source path does not exist or is not a directory: {imap_src}")
        sys.exit(1)

    os.makedirs(os.path.dirname(pst_dest), exist_ok=True)

    print("[*] Initializing Outlook MAPI Interface...")
    pythoncom.CoInitialize()
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook_app.GetNamespace("MAPI")
    namespace.Logon("", "", False, False)

    # Identify existing stores to isolate the new target
    existing_store_ids = set()
    for s in namespace.Stores:
        try:
            existing_store_ids.add(s.StoreID)
        except Exception:
            pass

    print(f"[*] Attaching target PST: {pst_dest}")
    try:
        namespace.AddStoreEx(pst_dest, 2)  # 2 = olStoreUnicode
    except Exception:
        try:
            namespace.AddStoreEx(pst_dest, 1)  # 1 = olStoreDefault
        except Exception:
            namespace.AddStore(pst_dest)

    target_store = None
    for s in namespace.Stores:
        try:
            if s.StoreID not in existing_store_ids:
                target_store = s
                break
        except Exception:
            continue

    if not target_store:
        for s in namespace.Stores:
            try:
                if s.FilePath and os.path.abspath(s.FilePath).lower() == pst_dest.lower():
                    target_store = s
                    break
            except Exception:
                continue

    if not target_store:
        print("[!] Error: Could not locate the newly attached PST store.")
        sys.exit(1)

    pst_root = target_store.GetRootFolder()
    print(f"[+] PST Root opened: {pst_root.Name}")

    temp_staging = tempfile.mkdtemp(prefix="tb_pst_conv_")

    try:
        traverse_and_convert(imap_src, pst_root, temp_staging, include_deleted=args.include_deleted)
    finally:
        print("[*] Cleaning temporary staging directory...")
        shutil.rmtree(temp_staging, ignore_errors=True)

        print("[*] Detaching PST store to flush buffers...")
        try:
            namespace.RemoveStore(pst_root)
        except Exception as e:
            print(f"[!] Notice during PST unmount: {e}")

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        print("[+] Conversion completed successfully.")


if __name__ == "__main__":
    main()
