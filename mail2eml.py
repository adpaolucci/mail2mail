#!/usr/bin/env python3
"""
mail2eml.py — Universal Mail/Mailbox -> EML Exporter (part of mail2mail)

Converts any supported email or mailbox format (MBOX, Thunderbird ImapMail/Mail,
Maildir/Maildir++, Apple Mail .emlx/.mbox, Outlook .msg, Outlook for Mac .olm,
Outlook .pst, MH, MMDF, Babyl, EML, and compressed .gz/.bz2/.xz/.zip/.tar* archives)
into standard .eml files with:
  - Flexible directory layouts: --layout {folder,year,month,flat}
  - Filename styles: --naming {subject,index}
  - Date range filtering: --start-year / --end-year
  - Per-directory file count & byte limits: --max-per-dir / --max-dir-bytes
  - Normalized status headers (Status, X-Status, X-Message-Flag, X-Mozilla-Status)
  - Deleted/expunged message filtering (override with --include-deleted)

Cross-platform (Windows, macOS, Linux); uses only the Python standard library
(except when reading an Outlook .pst source on Windows, which uses pywin32).
"""

from __future__ import annotations

import argparse
import email.parser
import hashlib
import os
import sys
import time
from typing import Dict, List, Optional, Tuple

from mail_formats import (
    MSG_FLAG_EXPUNGED,
    MSG_FLAG_FORWARDED,
    MSG_FLAG_MARKED,
    MSG_FLAG_READ,
    MSG_FLAG_REPLIED,
    WINDOWS_RESERVED_NAMES,
    decode_imap_utf7,
    decode_modified_utf7,
    extract_message_flags,
    iter_mail_records,
    normalize_eml_bytes_with_flags,
    pick_year_month,
    safe_get_header,
    sanitize_name,
    update_status_headers,
    win_long_path,
)

_HEADER_PARSER = email.parser.BytesHeaderParser()


# ----------------------- EML Filename & Limit Helpers -----------------------

def ensure_dir(path: str) -> str:
    os.makedirs(win_long_path(path), exist_ok=True)
    return path


def unique_eml_name(idx: int, msg) -> str:
    """Create a unique .eml filename using Subject + SHA-1(Message-ID + index + time)."""
    mid = safe_get_header(msg, "Message-ID", "").encode("utf-8", errors="ignore")
    h = hashlib.sha1(mid + str(idx).encode() + str(time.time_ns()).encode()).hexdigest()[:12]
    subj = sanitize_name(safe_get_header(msg, "Subject", "no_subject") or "no_subject", default="no_subject", max_len=100)
    return f"{subj}__{h}.eml"


def fits_limits(count_in_dir: int, bytes_in_dir: int, incoming_size: int, max_per_dir: int, max_dir_bytes: int) -> bool:
    if max_per_dir and count_in_dir >= max_per_dir:
        return False
    if max_dir_bytes and bytes_in_dir > 0 and (bytes_in_dir + incoming_size > max_dir_bytes):
        return False
    return True


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


# ----------------------- Core Multi-Format -> EML Exporter -----------------------

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
    Export messages from any supported format into .eml files.
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
        _, _, _, _, is_expunged = extract_message_flags(msg)

        if not include_deleted and (rec.is_deleted or is_expunged):
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
            safe_parts = [sanitize_name(p, default="Unnamed_Folder", max_len=100) for p in (rec.folder_parts or ["Inbox"])]
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


# ----------------------- CLI Entrypoint -----------------------

def build_arg_parser(prog: Optional[str] = None) -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog=prog,
        description=(
            "mail2eml — Convert any mail format (MBOX, Thunderbird ImapMail/Mail, "
            "Maildir/Maildir++, Apple Mail .emlx/.mbox, Outlook .msg/.olm/.pst, MH, MMDF, "
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
