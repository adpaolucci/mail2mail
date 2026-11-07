import os
import sys
import binascii

# Thunderbird X-Mozilla-Status bitmask flags
MSG_FLAG_READ = 0x0001
MSG_FLAG_REPLIED = 0x0002
MSG_FLAG_FORWARDED = 0x0004
MSG_FLAG_FLAGGED = 0x0008

FORBIDDEN_CHARS = r'\/:*?"<>|'

def decode_modified_utf7(name):
    """
    Decodes IMAP Modified UTF-7 (RFC 3501) folder names used by Thunderbird.
    Example: 'Notes &- Invoices' -> 'Notes & Invoices', 'Caf&AOk-' -> 'Café'
    """
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
    """Decodes IMAP UTF-7 and strips illegal filesystem characters."""
    decoded = decode_modified_utf7(name)
    clean = "".join('_' if c in FORBIDDEN_CHARS else c for c in decoded).strip()
    return clean or "Unnamed_Folder"

def is_probable_mbox(file_path):
    """Verifies if an extension-less file is an actual mbox mailbox."""
    basename = os.path.basename(file_path)
    ignored_exts = ('.msf', '.dat', '.json', '.html', '.txt', '.ini', '.bak', '.tmp', '.sqlite')
    if basename.startswith('.') or basename.lower().endswith(ignored_exts):
        return False

    if os.path.getsize(file_path) == 0:
        return True

    try:
        with open(file_path, 'rb') as f:
            head = f.read(1024).lstrip()
            return head.startswith(b"From ")
    except Exception:
        return False

def parse_thunderbird_flags(header_bytes):
    """Parses X-Mozilla-Status hex bitmasks directly from raw header bytes."""
    is_read = False
    is_flagged = False
    is_replied = False

    for line in header_bytes.splitlines():
        if line.lower().startswith(b'x-mozilla-status:'):
            parts = line.split(b':', 1)
            if len(parts) > 1:
                try:
                    val = int(parts[1].strip(), 16)
                    if val & MSG_FLAG_READ:
                        is_read = True
                    if val & MSG_FLAG_FLAGGED:
                        is_flagged = True
                    if val & MSG_FLAG_REPLIED:
                        is_replied = True
                except ValueError:
                    pass
            break
    return is_read, is_flagged, is_replied

def update_status_headers(raw_header_bytes, is_read, is_flagged, is_replied):
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

    # Append standard client-compatible status headers
    status_hdr = b'Status: RO\r\n' if is_read else b'Status: O\r\n'
    filtered_lines.append(status_hdr)

    x_status_val = b''
    if is_flagged:
        x_status_val += b'F'
    if is_replied:
        x_status_val += b'A'
    if x_status_val:
        filtered_lines.append(b'X-Status: ' + x_status_val + b'\r\n')

    if is_flagged:
        filtered_lines.append(b'X-Message-Flag: Follow up\r\n')

    return b"".join(filtered_lines)

def export_mbox_stream(mbox_path, dest_folder):
    """
    Single-pass streaming parser:
    - Streams byte chunks directly without memory overhead.
    - Accurately identifies mbox envelope lines via preceding-newline checks.
    - Reverses mboxrd '>From ' quoting artifacts.
    """
    os.makedirs(dest_folder, exist_ok=True)
    
    if os.path.getsize(mbox_path) == 0:
        return

    print(f"[*] Processing Mbox: {os.path.basename(mbox_path)} -> [{dest_folder}]")

    count = 0
    with open(mbox_path, 'rb') as f:
        line = f.readline()

        while line:
            # Advance to the first 'From ' envelope marker
            if not line.startswith(b'From '):
                line = f.readline()
                continue

            # Read raw header block until double newline
            header_lines = []
            while True:
                line = f.readline()
                if not line:
                    break
                if line in (b'\r\n', b'\n', b'\r'):
                    break
                header_lines.append(line)

            raw_headers = b"".join(header_lines)
            is_read, is_flagged, is_replied = parse_thunderbird_flags(raw_headers)
            normalized_headers = update_status_headers(raw_headers, is_read, is_flagged, is_replied)

            count += 1
            out_eml_path = os.path.join(dest_folder, f"{count:06d}.eml")

            with open(out_eml_path, 'wb') as out_f:
                out_f.write(normalized_headers.rstrip(b'\r\n'))
                out_f.write(b'\r\n\r\n')

                # Stream body until next valid mbox envelope line or EOF
                prev_line_empty = False
                while True:
                    line = f.readline()
                    if not line:
                        break

                    # True mbox boundary: 'From ' line preceded by an empty line
                    if prev_line_empty and line.startswith(b'From '):
                        break

                    # Track blank line state for the next cycle
                    prev_line_empty = line in (b'\r\n', b'\n', b'\r')

                    # Reverse mboxrd quote escaping (e.g. '>From ' -> 'From ', '>>From ' -> '>From ')
                    if line.startswith(b'>') and line.lstrip(b'>').startswith(b'From '):
                        out_f.write(line[1:])
                    else:
                        out_f.write(line)

            if count % 500 == 0:
                print(f"    -> {count} messages written...")

    print(f"    [+] Finished: {count} messages saved.")

def traverse_and_export(current_dir, current_out_dir):
    """Recursively walks ImapMail to replicate folder trees from mbox and .sbd containers."""
    try:
        entries = os.listdir(current_dir)
    except Exception as e:
        print(f"[!] Access error reading {current_dir}: {e}")
        return

    # 1. Process MBOX files
    for entry in entries:
        full_path = os.path.join(current_dir, entry)
        if os.path.isfile(full_path) and is_probable_mbox(full_path):
            folder_name = sanitize_filename(entry)
            target_eml_dir = os.path.join(current_out_dir, folder_name)
            export_mbox_stream(full_path, target_eml_dir)

    # 2. Process .sbd subdirectories
    for entry in entries:
        full_path = os.path.join(current_dir, entry)
        if os.path.isdir(full_path) and entry.lower().endswith(".sbd"):
            parent_folder_name = sanitize_filename(entry[:-4])
            nested_out_dir = os.path.join(current_out_dir, parent_folder_name)
            traverse_and_export(full_path, nested_out_dir)

def main():
    if len(sys.argv) < 3:
        print("Usage: python imapmail_to_eml.py <path_to_ImapMail_account_dir> <output_directory>")
        print(r'Example: python imapmail_to_eml.py "C:\path\to\ImapMail\imap.example.com" "C:\path\to\Exported_EMLs"')
        sys.exit(1)

    imap_src = os.path.abspath(sys.argv[1])
    dest_dir = os.path.abspath(sys.argv[2])

    if not os.path.isdir(imap_src):
        print(f"[!] Source path does not exist or is not a directory: {imap_src}")
        sys.exit(1)

    os.makedirs(dest_dir, exist_ok=True)
    print(f"[*] Starting direct EML export from: {imap_src}")
    print(f"[*] Destination directory: {dest_dir}\n")

    traverse_and_export(imap_src, dest_dir)
    print("\n[+] Full mailbox export finished successfully.")

if __name__ == "__main__":
    main()