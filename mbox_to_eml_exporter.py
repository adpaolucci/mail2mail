# mbox_to_eml_exporter.py
# Export messages from a .mbox file to .eml files with flexible folder layouts and filters.
# Features:
#   - --layout {year,month,flat}: choose directory structure (YYYY, YYYY/MM, or flat)
#   - --start-year / --end-year: export only messages within a year range (based on Date header)
#   - --max-per-dir N: cap number of files per directory (0 = no limit)
#   - --max-dir-bytes GB: cap total size per directory in GB (0 = no limit)
#   - --sanitize-filenames: make file names safe for Windows/macOS
#   - periodic progress logs and final stats
#
# Example:
#   python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout month --start-year 2005 --end-year 2016 --max-per-dir 50000 --sanitize-filenames
#   python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout month --sanitize-filenames
#   python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout year --start-year 2007 --end-year 2016
#   python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout month --max-per-dir 50000 --max-dir-bytes 9
#
import argparse
import hashlib
import mailbox
import os
import re
import sys
import time
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime

# Thunderbird X-Mozilla-Status bitmask flags (nsMsgMessageFlags.h)
MSG_FLAG_EXPUNGED = 0x0008

WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def get_header_str(msg, name: str, default: str = "") -> str:
    """Safely extract and decode an RFC 2047 header into a plain string."""
    raw_val = msg.get(name)
    if raw_val is None:
        return default
    try:
        return str(make_header(decode_header(str(raw_val))))
    except Exception:
        try:
            return str(raw_val)
        except Exception:
            return default


def is_thunderbird_expunged(msg) -> bool:
    """Return True if X-Mozilla-Status indicates the message was deleted/expunged (0x0008)."""
    tb_status = get_header_str(msg, 'X-Mozilla-Status', '').strip()
    if not tb_status:
        return False
    try:
        val = int(tb_status, 16)
        return bool(val & MSG_FLAG_EXPUNGED)
    except ValueError:
        return False


def safe_name(s: str, max_len: int = 100) -> str:
    """Return a filesystem-friendly string (safe for Windows/macOS)."""
    if not isinstance(s, str):
        s = str(s or "")
    s = re.sub(r'[\x00-\x1f\x7f\\/:*?"<>|]+', '_', s)
    s = re.sub(r'\s+', ' ', s).strip(' .')
    if not s:
        s = 'msg'
    if s.upper().split('.')[0] in WINDOWS_RESERVED_NAMES:
        s = f"_{s}"
    return s[:max_len].strip(' .') or 'msg'


def ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path


def pick_year_month(msg):
    """Extract (year, month) from the Date header; default to (1970, 1) if missing/invalid."""
    year = 1970
    month = 1
    try:
        d = get_header_str(msg, 'Date')
        if d:
            dt = parsedate_to_datetime(d)
            if dt is not None and 1900 <= dt.year <= 2100:
                year = dt.year
                month = dt.month
    except Exception:
        pass
    return year, month


def eml_bytes(msg):
    """Return raw EML bytes for a mailbox.mboxMessage."""
    try:
        return msg.as_bytes()
    except Exception:
        # Fallback: encode string representation if needed
        return msg.as_string().encode('utf-8', errors='ignore')


def unique_eml_name(idx, msg):
    """Create a unique file name using Message-ID + index + time-based salt."""
    mid = get_header_str(msg, 'Message-ID', '').encode('utf-8', errors='ignore')
    h = hashlib.sha1(mid + str(idx).encode() + str(time.time_ns()).encode()).hexdigest()[:12]
    subj = safe_name(get_header_str(msg, 'Subject', 'no_subject') or 'no_subject', max_len=100)
    return f"{subj}__{h}.eml"


def fits_limits(count_in_dir, bytes_in_dir, incoming_size, max_per_dir, max_dir_bytes):
    if max_per_dir and count_in_dir >= max_per_dir:
        return False
    if max_dir_bytes and bytes_in_dir > 0 and (bytes_in_dir + incoming_size > max_dir_bytes):
        return False
    return True


def main():
    ap = argparse.ArgumentParser(description="Export .mbox → .eml with layout and filters")
    ap.add_argument('--mbox', required=True, help='Path to the .mbox file')
    ap.add_argument('--out-dir', required=True, help='Destination directory for .eml files')
    ap.add_argument('--layout', choices=['year','month','flat'], default='year', help='Directory structure (year/month/flat)')
    ap.add_argument('--start-year', type=int, default=None, help='Export from this year (inclusive)')
    ap.add_argument('--end-year', type=int, default=None, help='Export up to this year (inclusive)')
    ap.add_argument('--max-per-dir', type=int, default=0, help='Max files per directory (0 = no limit)')
    ap.add_argument('--max-dir-bytes', type=float, default=0.0, help='Max total size per directory in GB (0 = no limit)')
    ap.add_argument('--sanitize-filenames', action='store_true', help='Sanitize filenames (enabled by default for safety)')
    ap.add_argument('--include-deleted', action='store_true', help='Include messages marked expunged (0x0008) in X-Mozilla-Status')
    ap.add_argument('--progress-every', type=int, default=1000, help='Log progress every N messages (0 = disabled)')
    args = ap.parse_args()

    mbox_path = os.path.normpath(args.mbox)
    out_root  = os.path.normpath(args.out_dir)
    if not os.path.isfile(mbox_path):
        print(f"Error: .mbox file not found: {mbox_path}", file=sys.stderr)
        sys.exit(1)
    os.makedirs(out_root, exist_ok=True)

    max_dir_bytes = int(args.max_dir_bytes * (1024**3)) if args.max_dir_bytes and args.max_dir_bytes > 0 else 0

    print(f"Reading MBOX: {mbox_path}")
    mb = mailbox.mbox(mbox_path, factory=mailbox.mboxMessage, create=False)
    try:
        total = len(mb)
        print(f"Messages in MBOX: {total:,}")

        # Directory trackers
        dir_counts = {}   # path → int
        dir_bytes  = {}   # path → int
        active_part = {}  # base_dir → current part int

        exported = 0
        skipped = 0
        started = time.time()

        for i, msg in enumerate(mb, 1):
            if not args.include_deleted and is_thunderbird_expunged(msg):
                skipped += 1
                continue

            year, month = pick_year_month(msg)

            # Year filters
            if args.start_year and year < args.start_year:
                skipped += 1
                continue
            if args.end_year and year > args.end_year:
                skipped += 1
                continue

            # Destination directory
            if args.layout == 'flat':
                base_dir = out_root
            elif args.layout == 'year':
                base_dir = os.path.join(out_root, f"{year:04d}")
            else:  # 'month'
                base_dir = os.path.join(out_root, f"{year:04d}", f"{month:02d}")

            raw = eml_bytes(msg)
            size = len(raw)

            # Per-dir limits: track active part per base_dir to avoid O(parts) re-scanning
            part = active_part.get(base_dir, 1)
            dest_dir = base_dir if part == 1 else f"{base_dir}__part{part}"
            c = dir_counts.get(dest_dir, 0)
            b = dir_bytes.get(dest_dir, 0)

            while not fits_limits(c, b, size, args.max_per_dir, max_dir_bytes):
                part += 1
                active_part[base_dir] = part
                dest_dir = f"{base_dir}__part{part}"
                c = dir_counts.get(dest_dir, 0)
                b = dir_bytes.get(dest_dir, 0)

            ensure_dir(dest_dir)

            # File name: unique_eml_name sanitizes the subject stem and appends __<hash>.eml
            fname = unique_eml_name(i, msg)
            dest_path = os.path.join(dest_dir, fname)

            try:
                with open(dest_path, 'wb') as f:
                    f.write(raw)
                dir_counts[dest_dir] = c + 1
                dir_bytes[dest_dir]  = b + size
                exported += 1
            except Exception as e:
                print(f"[WARN] Failed to save {dest_path}: {e}")
                skipped += 1

            if args.progress_every and args.progress_every > 0 and i % args.progress_every == 0:
                elapsed = time.time() - started
                rate = exported / elapsed if elapsed > 0 else 0
                print(f"{i:,}/{total:,} | exported: {exported:,} | skipped: {skipped:,} | {rate:.1f} msg/s")
    finally:
        mb.close()

    elapsed = time.time() - started
    print("\nDone.")
    print(f"Exported: {exported:,} | Skipped: {skipped:,} | Distinct dirs: {len(dir_counts):,}")
    if exported:
        print(f"Average speed: {exported/max(elapsed, 1e-9):.1f} msg/s")
        # Top 10 directories by total size
        top = sorted(dir_bytes.items(), key=lambda kv: kv[1], reverse=True)[:10]
        print("\nTop directories by size:")
        for p, bytes_ in top:
            print(f" - {p} → {bytes_/1e9:.2f} GB, {dir_counts.get(p,0)} files")

if __name__ == "__main__":
    main()
