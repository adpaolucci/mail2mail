import argparse
import binascii
import os
import re
import sys

# Thunderbird X-Mozilla-Status bitmask flags (from mailnews/base/public/nsMsgMessageFlags.h)
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
IGNORED_DIR_SUFFIXES = ('.mozmsgs', '.wdseml')


def win_long_path(path: str) -> str:
    """Add \\\\?\\ long-path prefix on Windows when path approaches MAX_PATH (260 chars)."""
    abspath = os.path.abspath(path)
    if os.name != "nt" or abspath.startswith("\\\\?\\") or len(abspath) < 240:
        return path
    if abspath.startswith("\\\\"):
        return "\\\\?\\UNC\\" + abspath[2:]
    return "\\\\?\\" + abspath


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
                    # Next message started before a blank line (either consecutive From_ or header-only msg)
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
                    # Advance stream to next 'From ' boundary without writing a file
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
                    # Stream body until next mbox 'From ' envelope line or EOF.
                    # Hold a single trailing blank line in pending_blank so the mbox
                    # inter-message blank line separator is not appended to the body.
                    pending_blank = b''
                    while True:
                        line = f.readline()
                        if not line:
                            if pending_blank:
                                out_f.write(pending_blank)
                            break

                        # In mboxrd (Thunderbird), body lines starting with 'From ' are always
                        # escaped as '>From ', so an unescaped 'From ' is an envelope boundary.
                        if line.startswith(b'From '):
                            break

                        if pending_blank:
                            out_f.write(pending_blank)
                            pending_blank = b''

                        if line in (b'\r\n', b'\n', b'\r'):
                            pending_blank = line
                            continue

                        # Reverse mboxrd quote escaping (e.g. '>From ' -> 'From ', '>>From ' -> '>From ')
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


def traverse_and_export(current_dir, current_out_dir, include_deleted=False):
    """
    Recursively walks ImapMail/Mail directories to replicate folder trees
    from mbox files, .sbd containers, and account subdirectories.
    Returns (total_written, total_skipped_deleted).
    """
    try:
        entries = sorted(os.listdir(win_long_path(current_dir)))
    except Exception as e:
        print(f"[!] Access error reading {current_dir}: {e}")
        return 0, 0

    total_written = 0
    total_skipped = 0

    # 1. Process MBOX files
    for entry in entries:
        full_path = os.path.join(current_dir, entry)
        if os.path.isfile(win_long_path(full_path)) and is_probable_mbox(full_path):
            folder_name = sanitize_filename(entry)
            target_eml_dir = os.path.join(current_out_dir, folder_name)
            w, s = export_mbox_stream(full_path, target_eml_dir, include_deleted=include_deleted)
            total_written += w
            total_skipped += s

    # 2. Process .sbd subdirectories and account folders (if user pointed at ImapMail/ or Mail/ root)
    for entry in entries:
        full_path = os.path.join(current_dir, entry)
        if not os.path.isdir(win_long_path(full_path)):
            continue
        lower = entry.lower()
        if entry.startswith('.') or lower.endswith(IGNORED_DIR_SUFFIXES):
            continue
        if lower.endswith(".sbd"):
            parent_folder_name = sanitize_filename(entry[:-4])
        else:
            parent_folder_name = sanitize_filename(entry)
        nested_out_dir = os.path.join(current_out_dir, parent_folder_name)
        w, s = traverse_and_export(full_path, nested_out_dir, include_deleted=include_deleted)
        total_written += w
        total_skipped += s

    return total_written, total_skipped


def main():
    for _stream in (sys.stdout, sys.stderr):
        if hasattr(_stream, "reconfigure"):
            try:
                _stream.reconfigure(errors="replace")
            except Exception:
                pass
    ap = argparse.ArgumentParser(
        description="Export a Thunderbird ImapMail/Mail account directory (or single mbox) to an .eml folder hierarchy."
    )
    ap.add_argument("imap_src", help="Path to Thunderbird ImapMail/Mail directory or a single mbox file")
    ap.add_argument("dest_dir", help="Destination directory for exported .eml files")
    ap.add_argument(
        "--include-deleted",
        action="store_true",
        help="Include messages marked as expunged/deleted (0x0008) in X-Mozilla-Status",
    )
    args = ap.parse_args()

    imap_src = os.path.abspath(args.imap_src)
    dest_dir = os.path.abspath(args.dest_dir)

    if not os.path.exists(win_long_path(imap_src)):
        print(f"[!] Source path does not exist: {imap_src}", file=sys.stderr)
        sys.exit(1)

    os.makedirs(win_long_path(dest_dir), exist_ok=True)
    print(f"[*] Starting direct EML export from: {imap_src}")
    print(f"[*] Destination directory: {dest_dir}\n")

    if os.path.isfile(win_long_path(imap_src)):
        folder_name = sanitize_filename(os.path.splitext(os.path.basename(imap_src))[0])
        target_eml_dir = os.path.join(dest_dir, folder_name)
        total_written, total_skipped = export_mbox_stream(
            imap_src, target_eml_dir, include_deleted=args.include_deleted
        )
    else:
        total_written, total_skipped = traverse_and_export(
            imap_src, dest_dir, include_deleted=args.include_deleted
        )

    print(
        f"\n[+] Full mailbox export finished: {total_written:,} messages exported"
        + (f" ({total_skipped:,} deleted/expunged skipped)." if total_skipped else ".")
    )


if __name__ == "__main__":
    main()
