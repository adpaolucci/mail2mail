import argparse
import binascii
import os
import re
import sys
from typing import Dict

from mail_formats import (
    MailRecord,
    WINDOWS_RESERVED_NAMES,
    iter_mail_records,
    win_long_path,
)

# Thunderbird X-Mozilla-Status bitmask flags (from mailnews/base/public/nsMsgMessageFlags.h)
MSG_FLAG_READ      = 0x0001
MSG_FLAG_REPLIED   = 0x0002
MSG_FLAG_MARKED    = 0x0004  # Flagged / Starred in Thunderbird
MSG_FLAG_EXPUNGED  = 0x0008  # Deleted (awaiting folder compaction)
MSG_FLAG_FORWARDED = 0x1000  # Forwarded

FORBIDDEN_CHARS = r'\/:*?"<>|'
IGNORED_DIR_SUFFIXES = ('.mozmsgs', '.wdseml')


def decode_modified_utf7(name):
    """
    Decodes IMAP Modified UTF-7 (RFC 3501) folder names used by Thunderbird and Maildir++.
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


def sanitize_filename(name):
    """Decodes IMAP UTF-7 and strips illegal Windows/POSIX filesystem characters."""
    decoded = decode_modified_utf7(name)
    clean = "".join(
        '_' if (c in FORBIDDEN_CHARS or ord(c) < 32 or ord(c) == 127) else c
        for c in decoded
    )
    clean = re.sub(r'\s+', ' ', clean).strip(' .')
    if not clean:
        return "Unnamed_Folder"
    if clean.upper().split('.')[0] in WINDOWS_RESERVED_NAMES:
        clean = f"_{clean}"
    return clean


def is_probable_mbox(file_path):
    """Verifies if an extension-less file is an actual mbox mailbox."""
    basename = os.path.basename(file_path)
    ignored_exts = (
        '.msf', '.dat', '.json', '.html', '.txt', '.ini',
        '.bak', '.tmp', '.sqlite', '.wdseml', '.mozmsgs', '.db', '.log'
    )
    if basename.startswith('.') or basename.lower().endswith(ignored_exts):
        return False

    lp = win_long_path(file_path)
    try:
        if os.path.getsize(lp) == 0:
            return True
        with open(lp, 'rb') as f:
            head = f.read(1024).lstrip()
            return head.startswith(b"From ")
    except Exception:
        return False


def parse_thunderbird_flags(header_bytes):
    """
    Parses X-Mozilla-Status hex bitmasks directly from raw header bytes.
    Returns (is_read, is_flagged, is_replied, is_forwarded, is_expunged).
    """
    is_read = False
    is_flagged = False
    is_replied = False
    is_forwarded = False
    is_expunged = False

    for line in header_bytes.splitlines():
        if line.lower().startswith(b'x-mozilla-status:'):
            parts = line.split(b':', 1)
            if len(parts) > 1:
                try:
                    val = int(parts[1].strip(), 16)
                    if val & MSG_FLAG_READ:
                        is_read = True
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
            break
    return is_read, is_flagged, is_replied, is_forwarded, is_expunged


def update_status_headers(raw_header_bytes, is_read, is_flagged, is_replied, is_forwarded=False):
    """
    Directly filters and updates status headers at the byte level.
    Preserves 100% of MIME structure, DKIM signatures, and line folding.
    """
    lines = raw_header_bytes.splitlines(keepends=True)
    filtered_lines = []
    skip_continuation = False
    target_prefixes = (b'status:', b'x-status:', b'x-message-flag:')

    for line in lines:
        lower = line.lower()
        if any(lower.startswith(prefix) for prefix in target_prefixes):
            skip_continuation = True
            continue
        if skip_continuation and (line.startswith(b' ') or line.startswith(b'\t')):
            continue
        skip_continuation = False
        filtered_lines.append(line)

    # Ensure last header line ends with CRLF before appending our headers
    if filtered_lines and not filtered_lines[-1].endswith((b'\r\n', b'\n', b'\r')):
        filtered_lines[-1] = filtered_lines[-1] + b'\r\n'

    # Append standard client-compatible status headers
    status_hdr = b'Status: RO\r\n' if is_read else b'Status: O\r\n'
    filtered_lines.append(status_hdr)

    x_status_val = b''
    if is_flagged:
        x_status_val += b'F'
    if is_replied:
        x_status_val += b'A'
    if is_forwarded:
        x_status_val += b'W'
    if x_status_val:
        filtered_lines.append(b'X-Status: ' + x_status_val + b'\r\n')

    if is_flagged:
        filtered_lines.append(b'X-Message-Flag: Follow up\r\n')

    return b"".join(filtered_lines)


def _normalize_eml_bytes_with_flags(raw_bytes: bytes, flags: Dict[str, bool]) -> bytes:
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

    tb_read, tb_flagged, tb_replied, tb_forwarded, _ = parse_thunderbird_flags(hdr)
    is_read = flags.get("is_read", tb_read)
    is_flagged = flags.get("is_flagged", False) or tb_flagged
    is_replied = flags.get("is_replied", False) or tb_replied
    is_forwarded = flags.get("is_forwarded", False) or tb_forwarded

    norm_hdr = update_status_headers(hdr, is_read, is_flagged, is_replied, is_forwarded)
    return norm_hdr.rstrip(b"\r\n") + b"\r\n\r\n" + body


def _next_eml_index(dest_folder):
    """Return the highest existing numeric .eml index in dest_folder so collisions append cleanly."""
    max_idx = 0
    try:
        for fn in os.listdir(win_long_path(dest_folder)):
            if fn.lower().endswith('.eml'):
                stem = os.path.splitext(fn)[0]
                if stem.isdigit():
                    max_idx = max(max_idx, int(stem))
    except OSError:
        pass
    return max_idx


def export_mbox_stream(mbox_path, dest_folder, include_deleted=False):
    """
    Single-pass streaming parser:
    - Streams byte chunks directly without memory overhead.
    - Accurately identifies mboxrd envelope lines (including after empty-body messages).
    - Skips Thunderbird expunged (deleted) messages unless include_deleted=True.
    - Reverses mboxrd '>From ' quoting artifacts.
    Returns (written_in_run, skipped_deleted).
    """
    os.makedirs(win_long_path(dest_folder), exist_ok=True)

    lp_mbox = win_long_path(mbox_path)
    if os.path.getsize(lp_mbox) == 0:
        return 0, 0

    print(f"[*] Processing Mbox: {os.path.basename(mbox_path)} -> [{dest_folder}]")

    count = _next_eml_index(dest_folder)
    written_in_run = 0
    skipped_deleted = 0

    with open(lp_mbox, 'rb') as f:
        line = f.readline()

        while line:
            # Advance to the next 'From ' envelope marker
            if not line.startswith(b'From '):
                line = f.readline()
                continue

            # Read raw header block until blank line (or unexpected next 'From ' marker)
            header_lines = []
            hit_next_from_in_headers = False
            while True:
                line = f.readline()
                if not line:
                    break
                if line in (b'\r\n', b'\n', b'\r'):
                    break
                if line.startswith(b'From '):
                    hit_next_from_in_headers = True
                    break
                header_lines.append(line)

            if hit_next_from_in_headers and not header_lines:
                continue

            raw_headers = b"".join(header_lines)
            is_read, is_flagged, is_replied, is_forwarded, is_expunged = parse_thunderbird_flags(raw_headers)

            if is_expunged and not include_deleted:
                skipped_deleted += 1
                if not hit_next_from_in_headers:
                    while True:
                        line = f.readline()
                        if not line or line.startswith(b'From '):
                            break
                continue

            normalized_headers = update_status_headers(
                raw_headers, is_read, is_flagged, is_replied, is_forwarded
            )

            count += 1
            written_in_run += 1
            out_eml_path = os.path.join(dest_folder, f"{count:06d}.eml")

            with open(win_long_path(out_eml_path), 'wb') as out_f:
                out_f.write(normalized_headers.rstrip(b'\r\n'))
                out_f.write(b'\r\n\r\n')

                if not hit_next_from_in_headers:
                    pending_blank = b''
                    while True:
                        line = f.readline()
                        if not line:
                            if pending_blank:
                                out_f.write(pending_blank)
                            break

                        if line.startswith(b'From '):
                            break

                        if pending_blank:
                            out_f.write(pending_blank)
                            pending_blank = b''

                        if line in (b'\r\n', b'\n', b'\r'):
                            pending_blank = line
                            continue

                        if line.startswith(b'>') and line.lstrip(b'>').startswith(b'From '):
                            out_f.write(line[1:])
                        else:
                            out_f.write(line)

            if written_in_run % 500 == 0:
                print(f"    -> {written_in_run} messages written...")

    msg_summary = f"    [+] Finished: {written_in_run} messages saved"
    if skipped_deleted:
        msg_summary += f" ({skipped_deleted} deleted/expunged skipped)"
    print(msg_summary + ".")
    return written_in_run, skipped_deleted


def export_multi_format_source(src_path, dest_root, include_deleted=False):
    """
    Universal exporter using mail_formats.iter_mail_records.
    Supports Thunderbird mbox & Maildir, Maildir/Maildir++, Apple Mail .emlx/.mbox bundles,
    Outlook .msg, Outlook for Mac .olm, MH, MMDF, Babyl, and compressed .gz/.bz2/.xz/.zip/.tar* archives.
    Returns (total_written, total_skipped_deleted).
    """
    folder_counters: Dict[str, int] = {}
    total_written = 0
    total_skipped = 0

    for rec in iter_mail_records(src_path, include_deleted=True):
        if rec.is_deleted and not include_deleted:
            total_skipped += 1
            continue

        safe_parts = [sanitize_filename(p) for p in (rec.folder_parts or ["Inbox"])]
        target_dir = os.path.join(dest_root, *safe_parts)
        if target_dir not in folder_counters:
            os.makedirs(win_long_path(target_dir), exist_ok=True)
            folder_counters[target_dir] = _next_eml_index(target_dir)

        folder_counters[target_dir] += 1
        idx = folder_counters[target_dir]
        out_eml_path = os.path.join(target_dir, f"{idx:06d}.eml")

        norm_bytes = _normalize_eml_bytes_with_flags(rec.raw_bytes, rec.flags)
        with open(win_long_path(out_eml_path), "wb") as out_f:
            out_f.write(norm_bytes)

        total_written += 1
        if total_written % 500 == 0:
            print(f"    -> {total_written} messages written...")

    return total_written, total_skipped


def traverse_and_export(current_dir, current_out_dir, include_deleted=False):
    """
    Walks a directory or archive of any supported mailbox format and exports a 1:1
    hierarchy of .eml files into current_out_dir.
    Returns (total_written, total_skipped_deleted).
    """
    return export_multi_format_source(current_dir, current_out_dir, include_deleted=include_deleted)


def main():
    for _stream in (sys.stdout, sys.stderr):
        if hasattr(_stream, "reconfigure"):
            try:
                _stream.reconfigure(errors="replace")
            except Exception:
                pass
    ap = argparse.ArgumentParser(
        description=(
            "Export Thunderbird ImapMail/Mail (mbox or Maildir), Maildir/Maildir++, "
            "Apple Mail (.mbox/.emlx), Outlook (.msg/.olm), MH, MMDF, Babyl, or "
            "compressed archives (.zip/.tgz/.gz) into an .eml folder hierarchy."
        )
    )
    ap.add_argument("imap_src", help="Path to source mail directory, mailbox file, or archive")
    ap.add_argument("dest_dir", help="Destination directory for exported .eml files")
    ap.add_argument(
        "--include-deleted",
        action="store_true",
        help="Include messages marked as expunged/deleted",
    )
    args = ap.parse_args()

    imap_src = os.path.abspath(args.imap_src)
    dest_dir = os.path.abspath(args.dest_dir)

    if not os.path.exists(win_long_path(imap_src)):
        print(f"[!] Source path does not exist: {imap_src}", file=sys.stderr)
        sys.exit(1)

    os.makedirs(win_long_path(dest_dir), exist_ok=True)
    print(f"[*] Starting EML export from: {imap_src}")
    print(f"[*] Destination directory: {dest_dir}\n")

    if os.path.isfile(win_long_path(imap_src)) and is_probable_mbox(imap_src):
        folder_name = sanitize_filename(os.path.splitext(os.path.basename(imap_src))[0])
        w, s = export_mbox_stream(
            imap_src, os.path.join(dest_dir, folder_name), include_deleted=args.include_deleted
        )
    else:
        w, s = export_multi_format_source(imap_src, dest_dir, include_deleted=args.include_deleted)

    print(f"\n[SUCCESS] Export complete! Total saved: {w:,} | Skipped deleted: {s:,}")


if __name__ == "__main__":
    main()
