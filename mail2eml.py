#!/usr/bin/env python3
"""
mail2eml.py — Universal Mail/Mailbox -> EML Exporter (part of mail2mail)

Converts any supported email or mailbox format (MBOX, Thunderbird ImapMail/Mail,
Maildir/Maildir++, Apple Mail .emlx/.mbox, Outlook .msg, Outlook for Mac .olm,
MH, MMDF, Babyl, EML, and compressed .gz/.bz2/.xz/.zip/.tar* archives) into
standard .eml files with:
  - Flexible directory layouts: --layout {folder,year,month,flat}
  - Filename styles: --naming {subject,index}
  - Date range filtering: --start-year / --end-year
  - Per-directory file count & byte limits: --max-per-dir / --max-dir-bytes
  - Normalized status headers (Status, X-Status, X-Message-Flag, X-Mozilla-Status)
  - Deleted/expunged message filtering (override with --include-deleted)

Cross-platform (Windows, macOS, Linux); uses only the Python standard library.
"""

from __future__ import annotations

import argparse
import binascii
import email.parser
import hashlib
import os
import re
import sys
import time
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
from typing import Dict, List, Optional, Tuple

from mail_formats import (
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
IGNORED_DIR_SUFFIXES = (".mozmsgs", ".wdseml")
_HEADER_PARSER = email.parser.BytesHeaderParser()


# ----------------------- String, Header & Path Helpers -----------------------

def decode_modified_utf7(name: str) -> str:
    """
    Decode IMAP Modified UTF-7 (RFC 3501) folder names used by Thunderbird and Maildir++.
    Example: 'Notes &- Invoices' -> 'Notes & Invoices', 'Caf&AOk-' -> 'Café'
    """
    if "&" not in name:
        return name
    res = []
    parts = name.split("&")
    res.append(parts[0])
    for part in parts[1:]:
        if "-" in part:
            b64, rest = part.split("-", 1)
            if b64 == "":
                res.append("&")
            else:
                b64_std = b64.replace(",", "/")
                pad = len(b64_std) % 4
                if pad:
                    b64_std += "=" * (4 - pad)
                try:
                    res.append(binascii.a2b_base64(b64_std).decode("utf-16-be"))
                except Exception:
                    res.append("&" + b64 + "-")
            res.append(rest)
        else:
            res.append("&" + part)
    return "".join(res)


def sanitize_filename(name: str) -> str:
    """Decode IMAP UTF-7 and strip illegal Windows/POSIX filesystem characters."""
    decoded = decode_modified_utf7(name)
    clean = "".join(
        "_" if (c in FORBIDDEN_CHARS or ord(c) < 32 or ord(c) == 127) else c
        for c in decoded
    )
    clean = re.sub(r"\s+", " ", clean).strip(" .")
    if not clean:
        return "Unnamed_Folder"
    if clean.upper().split(".")[0] in WINDOWS_RESERVED_NAMES:
        clean = f"_{clean}"
    return clean


def safe_name(s: str, max_len: int = 100) -> str:
    """Return a filesystem-friendly string stem (safe for Windows/macOS/Linux)."""
    if not isinstance(s, str):
        s = str(s or "")
    s = re.sub(r'[\x00-\x1f\x7f\\/:*?"<>|]+', "_", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    if not s:
        s = "msg"
    if s.upper().split(".")[0] in WINDOWS_RESERVED_NAMES:
        s = f"_{s}"
    return s[:max_len].strip(" .") or "msg"


def ensure_dir(path: str) -> str:
    os.makedirs(win_long_path(path), exist_ok=True)
    return path


def get_header_str(msg, name: str, default: str = "") -> str:
    """Safely extract and decode an RFC 2047 header into a plain string."""
    try:
        raw_val = msg.get(name)
    except Exception:
        raw_val = None
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
    tb_status = get_header_str(msg, "X-Mozilla-Status", "").strip()
    if not tb_status:
        return False
    try:
        val = int(tb_status, 16)
        return bool(val & MSG_FLAG_EXPUNGED)
    except ValueError:
        return False


def pick_year_month(msg) -> Tuple[int, int]:
    """Extract (year, month) from the Date header; default to (1970, 1) if missing/invalid."""
    year = 1970
    month = 1
    try:
        d = get_header_str(msg, "Date")
        if d:
            dt = parsedate_to_datetime(d)
            if dt is not None and 1900 <= dt.year <= 2100:
                year = dt.year
                month = dt.month
    except Exception:
        pass
    return year, month


def eml_bytes(msg) -> bytes:
    """Return raw EML bytes for an email.message.Message or mailbox.mboxMessage."""
    try:
        return msg.as_bytes()
    except Exception:
        return msg.as_string().encode("utf-8", errors="ignore")


def unique_eml_name(idx: int, msg) -> str:
    """Create a unique file name using Message-ID + index + time-based salt."""
    mid = get_header_str(msg, "Message-ID", "").encode("utf-8", errors="ignore")
    h = hashlib.sha1(mid + str(idx).encode() + str(time.time_ns()).encode()).hexdigest()[:12]
    subj = safe_name(get_header_str(msg, "Subject", "no_subject") or "no_subject", max_len=100)
    return f"{subj}__{h}.eml"


def fits_limits(count_in_dir: int, bytes_in_dir: int, incoming_size: int, max_per_dir: int, max_dir_bytes: int) -> bool:
    if max_per_dir and count_in_dir >= max_per_dir:
        return False
    if max_dir_bytes and bytes_in_dir > 0 and (bytes_in_dir + incoming_size > max_dir_bytes):
        return False
    return True


def is_probable_mbox(file_path: str) -> bool:
    """Verify if an extensionless file is an actual mbox mailbox."""
    basename = os.path.basename(file_path)
    ignored_exts = (
        ".msf", ".dat", ".json", ".html", ".txt", ".ini",
        ".bak", ".tmp", ".sqlite", ".wdseml", ".mozmsgs", ".db", ".log",
    )
    if basename.startswith(".") or basename.lower().endswith(ignored_exts):
        return False

    lp = win_long_path(file_path)
    try:
        if os.path.getsize(lp) == 0:
            return True
        with open(lp, "rb") as f:
            head = f.read(1024).lstrip()
            return head.startswith(b"From ")
    except Exception:
        return False


def parse_thunderbird_flags(header_bytes: bytes) -> Tuple[bool, bool, bool, bool, bool]:
    """
    Parse X-Mozilla-Status hex bitmasks directly from raw header bytes.
    Returns (is_read, is_flagged, is_replied, is_forwarded, is_expunged).
    """
    is_read = False
    is_flagged = False
    is_replied = False
    is_forwarded = False
    is_expunged = False

    for line in header_bytes.splitlines():
        if line.lower().startswith(b"x-mozilla-status:"):
            parts = line.split(b":", 1)
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


def update_status_headers(
    raw_header_bytes: bytes,
    is_read: bool,
    is_flagged: bool,
    is_replied: bool,
    is_forwarded: bool = False,
) -> bytes:
    """
    Directly filter and update status headers at the byte level.
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

    tb_read, tb_flagged, tb_replied, tb_forwarded, _ = parse_thunderbird_flags(hdr)
    is_read = flags.get("is_read", tb_read)
    is_flagged = flags.get("is_flagged", False) or tb_flagged
    is_replied = flags.get("is_replied", False) or tb_replied
    is_forwarded = flags.get("is_forwarded", False) or tb_forwarded

    norm_hdr = update_status_headers(hdr, is_read, is_flagged, is_replied, is_forwarded)
    return norm_hdr.rstrip(b"\r\n") + b"\r\n\r\n" + body


def _next_eml_index(dest_folder: str) -> int:
    """Return the highest existing numeric .eml index in dest_folder so collisions append cleanly."""
    max_idx = 0
    try:
        for fn in os.listdir(win_long_path(dest_folder)):
            if fn.lower().endswith(".eml"):
                stem = os.path.splitext(fn)[0]
                if stem.isdigit():
                    max_idx = max(max_idx, int(stem))
    except OSError:
        pass
    return max_idx


# ----------------------- Stream & Multi-Format Exporters -----------------------

def export_mbox_stream(mbox_path: str, dest_folder: str, include_deleted: bool = False) -> Tuple[int, int]:
    """
    Single-pass streaming mbox parser:
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

    with open(lp_mbox, "rb") as f:
        line = f.readline()

        while line:
            if not line.startswith(b"From "):
                line = f.readline()
                continue

            header_lines = []
            hit_next_from_in_headers = False
            while True:
                line = f.readline()
                if not line:
                    break
                if line in (b"\r\n", b"\n", b"\r"):
                    break
                if line.startswith(b"From "):
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
                        if not line or line.startswith(b"From "):
                            break
                continue

            normalized_headers = update_status_headers(
                raw_headers, is_read, is_flagged, is_replied, is_forwarded
            )

            count += 1
            written_in_run += 1
            out_eml_path = os.path.join(dest_folder, f"{count:06d}.eml")

            with open(win_long_path(out_eml_path), "wb") as out_f:
                out_f.write(normalized_headers.rstrip(b"\r\n"))
                out_f.write(b"\r\n\r\n")

                if not hit_next_from_in_headers:
                    pending_blank = b""
                    while True:
                        line = f.readline()
                        if not line:
                            if pending_blank:
                                out_f.write(pending_blank)
                            break

                        if line.startswith(b"From "):
                            break

                        if pending_blank:
                            out_f.write(pending_blank)
                            pending_blank = b""

                        if line in (b"\r\n", b"\n", b"\r"):
                            pending_blank = line
                            continue

                        if line.startswith(b">") and line.lstrip(b">").startswith(b"From "):
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


def export_to_eml(
    src_path: str,
    out_root: str,
    layout: str = "folder",
    naming: str = "subject",
    start_year: Optional[int] = None,
    end_year: Optional[int] = None,
    max_per_dir: int = 0,
    max_dir_gb: float = 0.0,
    include_deleted: bool = False,
    normalize_headers: bool = True,
    progress_every: int = 1000,
) -> Tuple[int, int, Dict[str, int], Dict[str, int]]:
    """
    Core unified exporter from any supported format into .eml files.
    Returns (exported_count, skipped_count, dir_counts, dir_bytes).
    """
    ensure_dir(out_root)
    max_dir_bytes = int(max_dir_gb * (1024**3)) if max_dir_gb and max_dir_gb > 0 else 0

    dir_counts: Dict[str, int] = {}
    dir_bytes: Dict[str, int] = {}
    active_part: Dict[str, int] = {}
    folder_indices: Dict[str, int] = {}

    exported = 0
    skipped = 0
    started = time.time()

    for i, rec in enumerate(iter_mail_records(src_path, include_deleted=True), 1):
        raw = rec.raw_bytes
        msg = _HEADER_PARSER.parsebytes(raw)

        if not include_deleted and (rec.is_deleted or is_thunderbird_expunged(msg)):
            skipped += 1
            continue

        year, month = pick_year_month(msg)
        if start_year and year < start_year:
            skipped += 1
            continue
        if end_year and year > end_year:
            skipped += 1
            continue

        if layout == "flat":
            base_dir = out_root
        elif layout == "year":
            base_dir = os.path.join(out_root, f"{year:04d}")
        elif layout == "month":
            base_dir = os.path.join(out_root, f"{year:04d}", f"{month:02d}")
        else:  # 'folder'
            safe_parts = [sanitize_filename(p) for p in (rec.folder_parts or ["Inbox"])]
            base_dir = os.path.join(out_root, *safe_parts)

        out_bytes = normalize_eml_bytes_with_flags(raw, rec.flags) if normalize_headers else raw
        size = len(out_bytes)

        part = active_part.get(base_dir, 1)
        dest_dir = base_dir if part == 1 else f"{base_dir}__part{part}"
        c = dir_counts.get(dest_dir, 0)
        b = dir_bytes.get(dest_dir, 0)

        while not fits_limits(c, b, size, max_per_dir, max_dir_bytes):
            part += 1
            active_part[base_dir] = part
            dest_dir = f"{base_dir}__part{part}"
            c = dir_counts.get(dest_dir, 0)
            b = dir_bytes.get(dest_dir, 0)

        ensure_dir(dest_dir)

        if naming == "index":
            if dest_dir not in folder_indices:
                folder_indices[dest_dir] = _next_eml_index(dest_dir)
            folder_indices[dest_dir] += 1
            fname = f"{folder_indices[dest_dir]:06d}.eml"
        else:
            fname = unique_eml_name(i, msg)

        dest_path = os.path.join(dest_dir, fname)
        try:
            with open(win_long_path(dest_path), "wb") as f:
                f.write(out_bytes)
            dir_counts[dest_dir] = c + 1
            dir_bytes[dest_dir] = b + size
            exported += 1
        except Exception as e:
            print(f"[WARN] Failed to save {dest_path}: {e}")
            skipped += 1

        if progress_every and progress_every > 0 and i % progress_every == 0:
            elapsed = time.time() - started
            rate = exported / elapsed if elapsed > 0 else 0
            print(f"Processed {i:,} | exported: {exported:,} | skipped: {skipped:,} | {rate:.1f} msg/s")

    return exported, skipped, dir_counts, dir_bytes


def export_multi_format_source(src_path: str, dest_root: str, include_deleted: bool = False) -> Tuple[int, int]:
    """Export any supported source into a 1:1 folder hierarchy with sequential numeric .eml names."""
    exported, skipped, _, _ = export_to_eml(
        src_path=src_path,
        out_root=dest_root,
        layout="folder",
        naming="index",
        include_deleted=include_deleted,
        normalize_headers=True,
        progress_every=500,
    )
    return exported, skipped


def traverse_and_export(current_dir: str, current_out_dir: str, include_deleted: bool = False) -> Tuple[int, int]:
    """Walk any supported directory or archive and export a 1:1 hierarchy of .eml files."""
    return export_multi_format_source(current_dir, current_out_dir, include_deleted=include_deleted)


def build_arg_parser(prog: Optional[str] = None) -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog=prog,
        description=(
            "mail2eml — Convert any mail format (MBOX, Thunderbird ImapMail/Mail, "
            "Maildir/Maildir++, Apple Mail .emlx/.mbox, Outlook .msg/.olm, MH, MMDF, "
            "Babyl, or .zip/.tgz/.gz archives) into .eml files."
        ),
    )
    ap.add_argument("src_pos", nargs="?", help="Source mail directory, mailbox file, or archive")
    ap.add_argument("out_pos", nargs="?", help="Destination directory for exported .eml files")
    ap.add_argument("--src", "--mbox", dest="src_opt", help="Source mail directory, mailbox file, or archive")
    ap.add_argument("--out-dir", "--out", dest="out_opt", help="Destination directory for .eml files")
    ap.add_argument(
        "--layout",
        choices=["folder", "year", "month", "flat"],
        default=None,
        help="Output directory layout (default: 'folder' for positional/hierarchy, 'year' when --mbox is used)",
    )
    ap.add_argument(
        "--naming",
        choices=["subject", "index"],
        default="subject",
        help="EML filename format: 'subject' (<subject>__<hash>.eml) or 'index' (000001.eml)",
    )
    ap.add_argument("--start-year", type=int, default=None, help="Export from this year (inclusive)")
    ap.add_argument("--end-year", type=int, default=None, help="Export up to this year (inclusive)")
    ap.add_argument("--max-per-dir", type=int, default=0, help="Max files per directory before rolling to __part2 (0 = unlimited)")
    ap.add_argument("--max-dir-bytes", type=float, default=0.0, help="Max size per directory in GB before rolling to __part2 (0 = unlimited)")
    ap.add_argument("--sanitize-filenames", action="store_true", help="Sanitize filenames (enabled by default)")
    ap.add_argument("--include-deleted", action="store_true", help="Include messages marked expunged/deleted")
    ap.add_argument("--progress-every", type=int, default=1000, help="Log progress every N messages (0 = disabled)")
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
    out_dir = args.out_opt or args.out_pos
    if not src or not out_dir:
        ap.error("Both a source path (positional or --src/--mbox) and destination (--out-dir) are required.")

    src_path = os.path.normpath(src)
    out_root = os.path.normpath(out_dir)
    if not os.path.exists(win_long_path(src_path)):
        print(f"Error: source path not found: {src_path}", file=sys.stderr)
        sys.exit(1)

    layout = args.layout
    if layout is None:
        # Default to 'folder' for directory hierarchies, or 'year' if invoked via --mbox without --layout
        layout = "year" if (args.src_opt and not args.src_pos) else "folder"

    print(f"[*] Source: {src_path}")
    print(f"[*] Output directory: {out_root} (layout={layout}, naming={args.naming})")

    started = time.time()
    exported, skipped, dir_counts, dir_bytes = export_to_eml(
        src_path=src_path,
        out_root=out_root,
        layout=layout,
        naming=args.naming,
        start_year=args.start_year,
        end_year=args.end_year,
        max_per_dir=args.max_per_dir,
        max_dir_gb=args.max_dir_bytes,
        include_deleted=args.include_deleted,
        normalize_headers=True,
        progress_every=args.progress_every,
    )

    elapsed = time.time() - started
    print("\nDone.")
    print(f"Exported: {exported:,} | Skipped: {skipped:,} | Distinct dirs: {len(dir_counts):,}")
    if exported:
        print(f"Average speed: {exported / max(elapsed, 1e-9):.1f} msg/s")
        top = sorted(dir_bytes.items(), key=lambda kv: kv[1], reverse=True)[:10]
        print("\nTop directories by size:")
        for p, bytes_ in top:
            print(f" - {p} -> {bytes_ / 1e9:.2f} GB, {dir_counts.get(p, 0)} files")


if __name__ == "__main__":
    main()
