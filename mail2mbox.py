#!/usr/bin/env python3
"""
mail2mbox.py — Universal Mail/Mailbox -> MBOX Converter & Exporter (part of mail2mail)

Converts any supported email or mailbox format:
  - Standard .mbox & Thunderbird ImapMail/Mail (mbox and Maildir modes)
  - Maildir & Maildir++ (cur/, new/, .Folder.Subfolder)
  - Apple Mail (.emlx, .partial.emlx, *.mbox / *.imapmbox bundles)
  - Outlook .msg (standalone Compound File Binary)
  - Outlook for Mac .olm archives
  - Outlook .pst Personal Folders (via Outlook MAPI on Windows when Outlook is installed)
  - MH / Claws Mail / Sylpheed directories
  - MMDF & Emacs Babyl files
  - Standard .eml files & directory trees
  - Compressed archives (.gz, .bz2, .xz, .zip, .tar, .tar.gz, .tgz, .tar.bz2, .tar.xz)
into standard RFC 4155 / mboxrd `.mbox` file(s) (including compressed `.mbox.gz`, `.mbox.bz2`, `.mbox.xz`).

Supports both:
  1) Single-file MBOX output (optionally compressed):
       python mail2mbox.py "D:\\Export_EML" "D:\\Backups\\AllMail.mbox"
       python mail2mbox.py "D:\\Mail\\Archive.pst" "D:\\Backups\\Archive.mbox.gz"
  2) Multi-file / Folder-hierarchy MBOX directory output:
       python mail2mbox.py "D:\\Mail\\ImapMail" "D:\\Backups\\MboxTree" --layout folder
       python mail2mbox.py "D:\\Mail\\ImapMail" "D:\\Backups\\ByYear" --layout year
"""

from __future__ import annotations

import argparse
import bz2
import email
import gzip
import lzma
import os
import re
import sys
import time
from datetime import datetime, timezone
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from typing import Dict, IO, List, Optional, Tuple

from mail_formats import (
    extract_message_flags,
    iter_mail_records,
    normalize_eml_bytes_with_flags,
    pick_year_month,
    safe_get_all_headers,
    safe_get_header,
    sanitize_name,
    win_long_path,
)
from mail_audit import AuditLogger, AuditRecord, compute_sha256, perform_dry_run

MBOXRD_BODY_FROM_RE = re.compile(rb"^(>*From )", re.MULTILINE)
MBOX_FILE_EXTS = (".mbox", ".mbx", ".mbox.gz", ".mbox.bz2", ".mbox.xz")


def is_mbox_output_filename(path: str) -> bool:
    """Return True if path looks like a target .mbox (or compressed .mbox.*) file."""
    lower = path.lower()
    return any(lower.endswith(ext) for ext in MBOX_FILE_EXTS)


def build_mbox_from_line(msg, fallback_dt: Optional[datetime] = None) -> bytes:
    """
    Construct a standard RFC 4155 `From <sender> <24-char-asctime>\\n` envelope line
    from a parsed `email.message.Message`.
    """
    sender = "MAILER-DAEMON"
    from_hdrs = safe_get_all_headers(msg, "Return-Path") or safe_get_all_headers(msg, "From")
    if from_hdrs:
        pairs = getaddresses(from_hdrs)
        for _, addr in pairs:
            clean_addr = (addr or "").strip()
            if clean_addr and not any(c.isspace() for c in clean_addr):
                sender = clean_addr
                break

    dt = fallback_dt
    if isinstance(dt, str):
        try:
            dt = parsedate_to_datetime(dt)
        except Exception:
            dt = None
    if dt is None:
        date_str = safe_get_header(msg, "Date", "").strip()
        if date_str:
            try:
                dt = parsedate_to_datetime(date_str)
            except Exception:
                dt = None

    if dt is None:
        dt = datetime(1970, 1, 1, 0, 0, 0, tzinfo=timezone.utc)
    elif dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc)
    else:
        dt = dt.replace(tzinfo=timezone.utc)

    if dt.year < 1900:
        day = min(dt.day, 28) if dt.month == 2 else dt.day
        dt = dt.replace(year=1970, day=day)
    elif dt.year > 2100:
        day = min(dt.day, 28) if dt.month == 2 else dt.day
        dt = dt.replace(year=2100, day=day)

    asctime_str = time.asctime(dt.utctimetuple())
    line = f"From {sender} {asctime_str}\n"
    return line.encode("ascii", errors="replace")


def format_mboxrd_entry(raw_bytes: bytes, msg=None, flags=None) -> bytes:
    """
    Format raw RFC 5322 message bytes as a complete `mboxrd` entry:
      1. Inject/normalize X-Mozilla-Status, Status, X-Status headers if `flags` are provided.
      2. Strip any leading `From ` envelope line if present.
      3. Escape any `^>*From ` lines in the message body (`>From `).
      4. Prepend `From <sender> <asctime>\\n` and ensure a trailing blank line (`\\n\\n`).
    """
    normalized = normalize_eml_bytes_with_flags(raw_bytes, flags)

    if normalized.startswith(b"From "):
        nl = normalized.find(b"\n")
        if nl != -1:
            normalized = normalized[nl + 1:]

    # Normalize CRLF to LF for standard mbox storage and clean regex matching
    normalized = normalized.replace(b"\r\n", b"\n")

    if msg is None:
        try:
            msg = BytesParser(policy=policy.default).parsebytes(normalized)
        except Exception:
            msg = email.message_from_bytes(b"Subject: (no subject)\n\n", policy=policy.default)

    sep = normalized.find(b"\n\n")
    if sep != -1:
        hdr_block = normalized[:sep + 2]
        body_block = normalized[sep + 2:]
        body_block = MBOXRD_BODY_FROM_RE.sub(rb">\1", body_block)
        payload = hdr_block + body_block
    else:
        payload = MBOXRD_BODY_FROM_RE.sub(rb">\1", normalized)

    if not payload.endswith(b"\n"):
        payload += b"\n"
    payload += b"\n"

    from_line = build_mbox_from_line(msg)
    return from_line + payload


def _open_mbox_stream(path: str, mode: str = "wb") -> IO[bytes]:
    """Open a plain or compressed (.gz, .bz2, .xz) mbox file for binary writing."""
    lp = win_long_path(path)
    parent = os.path.dirname(lp)
    if parent:
        os.makedirs(parent, exist_ok=True)
    lower = path.lower()
    if lower.endswith(".gz"):
        return gzip.open(lp, mode)
    if lower.endswith(".bz2"):
        return bz2.open(lp, mode)
    if lower.endswith(".xz"):
        return lzma.open(lp, mode)
    return open(lp, mode)


def _split_mbox_ext(path: str) -> Tuple[str, str]:
    """Split `.mbox.gz` / `.mbox.bz2` / `.mbox.xz` / `.mbox` into `(stem, ext)`."""
    lower = path.lower()
    for compound in (".mbox.gz", ".mbox.bz2", ".mbox.xz"):
        if lower.endswith(compound):
            return path[:-len(compound)], path[-len(compound):]
    return os.path.splitext(path)


class MboxWriterPool:
    """
    Manages open `.mbox` output file handles and optional size-based rotation (`--max-mbox-gb`).
    Maintains an LRU cache of at most `max_open` simultaneously open handles to prevent
    file descriptor exhaustion (e.g. EMFILE / Too many open files) and excessive RAM
    usage from concurrent compression buffers.
    """
    def __init__(self, max_gb: float = 0.0, overwrite: bool = True, max_open: int = 64):
        self.max_bytes = int(max_gb * (1024**3)) if max_gb and max_gb > 0 else 0
        self.overwrite = overwrite
        self.max_open = max(4, max_open)
        self._states: Dict[str, Dict] = {}
        self._open_order: List[str] = []

    def _part_path(self, base_path: str, part: int) -> str:
        if part <= 1:
            return base_path
        stem, ext = _split_mbox_ext(base_path)
        return f"{stem}__part{part}{ext}"

    def _close_handle(self, norm_key: str) -> None:
        state = self._states.get(norm_key)
        if state and state.get("fh") is not None:
            try:
                state["fh"].close()
            except Exception:
                pass
            state["fh"] = None
        if norm_key in self._open_order:
            self._open_order.remove(norm_key)

    def _ensure_open_fh(self, norm_key: str, state: Dict, mode: str):
        if state.get("fh") is not None:
            if norm_key in self._open_order:
                self._open_order.remove(norm_key)
            self._open_order.append(norm_key)
            return state["fh"]

        while len(self._open_order) >= self.max_open:
            oldest_key = self._open_order.pop(0)
            oldest_state = self._states.get(oldest_key)
            if oldest_state and oldest_state.get("fh") is not None:
                try:
                    oldest_state["fh"].close()
                except Exception:
                    pass
                oldest_state["fh"] = None

        fh = _open_mbox_stream(state["path"], mode=mode)
        state["fh"] = fh
        self._open_order.append(norm_key)
        return fh

    def write_entry(self, target_path: str, entry_bytes: bytes) -> str:
        norm_key = os.path.normcase(os.path.abspath(target_path))
        state = self._states.get(norm_key)
        entry_len = len(entry_bytes)

        if state is None:
            part = 1
            actual_path = self._part_path(target_path, part)
            mode = "wb" if self.overwrite else "ab"
            init_bytes = 0
            if not self.overwrite and os.path.isfile(win_long_path(actual_path)):
                try:
                    init_bytes = os.path.getsize(win_long_path(actual_path))
                except OSError:
                    init_bytes = 0
            state = {"fh": None, "path": actual_path, "bytes": init_bytes, "part": part, "base_path": target_path}
            self._states[norm_key] = state
            fh = self._ensure_open_fh(norm_key, state, mode=mode)
        elif self.max_bytes and state["bytes"] > 0 and (state["bytes"] + entry_len > self.max_bytes):
            self._close_handle(norm_key)
            state["part"] += 1
            actual_path = self._part_path(state["base_path"], state["part"])
            state.update({"path": actual_path, "bytes": 0})
            fh = self._ensure_open_fh(norm_key, state, mode="wb" if self.overwrite else "ab")
        else:
            fh = self._ensure_open_fh(norm_key, state, mode="ab")

        fh.write(entry_bytes)
        state["bytes"] += entry_len
        return state["path"]

    def close_all(self) -> List[str]:
        paths = []
        for state in self._states.values():
            paths.append(state["path"])
            if state.get("fh") is not None:
                try:
                    state["fh"].close()
                except Exception:
                    pass
                state["fh"] = None
        self._open_order.clear()
        return paths


def resolve_mbox_dest_path(
    out_dir: str,
    layout: str,
    folder_parts: List[str],
    year: int,
    month: int,
    default_name: str = "Inbox",
) -> str:
    """
    Determine the target `.mbox` file path inside `out_dir` according to `layout`:
      - 'folder': `<out_dir>/<Folder>/<Subfolder>.mbox`
      - 'year':   `<out_dir>/<Folder>/<YYYY>.mbox` (or `<out_dir>/<YYYY>.mbox` if no folders)
      - 'month':  `<out_dir>/<Folder>/<YYYY>/<YYYY-MM>.mbox`
      - 'flat':   `<out_dir>/<default_name>.mbox`
    """
    clean_parts = [sanitize_name(p, default="Folder") for p in (folder_parts or []) if p and p != "."]

    if layout == "flat":
        return os.path.join(out_dir, f"{sanitize_name(default_name, default='Inbox')}.mbox")

    if layout == "year":
        if clean_parts:
            return os.path.join(out_dir, *clean_parts, f"{year:04d}.mbox")
        return os.path.join(out_dir, f"{year:04d}.mbox")

    if layout == "month":
        if clean_parts:
            return os.path.join(out_dir, *clean_parts, f"{year:04d}", f"{year:04d}-{month:02d}.mbox")
        return os.path.join(out_dir, f"{year:04d}", f"{year:04d}-{month:02d}.mbox")

    # Default: 'folder' layout
    if not clean_parts:
        return os.path.join(out_dir, f"{sanitize_name(default_name, default='Inbox')}.mbox")
    if len(clean_parts) == 1:
        return os.path.join(out_dir, f"{clean_parts[0]}.mbox")
    return os.path.join(out_dir, *clean_parts[:-1], f"{clean_parts[-1]}.mbox")


def export_to_mbox(
    src_path: str,
    target_path: str,
    single_file_mode: bool = False,
    layout: str = "folder",
    start_year: Optional[int] = None,
    end_year: Optional[int] = None,
    include_deleted: bool = False,
    max_mbox_gb: float = 0.0,
    overwrite: bool = True,
    logger: Optional[AuditLogger] = None,
    dedup: str = "none",
    resume: Optional[Any] = None,
    parse_gmail: bool = True,
) -> Tuple[int, int]:
    """
    Universal converter that streams any supported mailbox format via
    `mail_formats.iter_mail_records` and writes standard `mboxrd` file(s).
    Returns `(written_count, skipped_count)`.
    """
    parser = BytesParser(policy=policy.default)
    pool = MboxWriterPool(max_gb=max_mbox_gb, overwrite=overwrite)
    written = 0
    skipped = 0
    skipped_date = 0
    skipped_deleted = 0
    skipped_duplicate = 0
    failed = 0
    total_bytes = 0
    started = time.time()
    idx = 0

    seen_keys: Set[str] = set()
    seen_mid_signatures: Dict[str, Tuple[str, str, str]] = {}
    if resume:
        resume_source = resume if isinstance(resume, str) and os.path.isfile(resume) else (logger.audit_log_path if logger else None)
        if resume_source and os.path.isfile(win_long_path(resume_source)):
            existing = AuditLogger.load_completed_keys(resume_source)
            seen_keys.update(existing)
            if logger:
                logger.info(f"[*] Resuming: loaded {len(existing):,} previously completed message signatures from {resume_source}")
        elif resume is not True and logger:
            logger.warn(f"Resume source manifest not found: {resume}")

    src_stem = os.path.splitext(os.path.basename(os.path.normpath(src_path)))[0] or "Inbox"

    try:
        for idx, rec in enumerate(iter_mail_records(src_path, include_deleted=True, parse_gmail=parse_gmail), 1):
            raw_bytes = rec.raw_bytes
            folder_display = os.sep.join(rec.folder_parts) if rec.folder_parts else "Root"
            if not raw_bytes:
                skipped += 1
                failed += 1
                continue

            try:
                msg = parser.parsebytes(raw_bytes)
            except Exception:
                msg = email.message_from_bytes(b"Subject: (no subject)\r\n\r\n", policy=policy.default)

            mid = safe_get_header(msg, "Message-ID", "").strip()
            subj_val = safe_get_header(msg, "Subject", "(no subject)").strip()
            date_val = safe_get_header(msg, "Date", "").strip()
            is_dup = False
            dup_key = None

            if dedup != "none" or resume:
                sha = compute_sha256(raw_bytes) if (dedup in ("sha256", "auto") or not mid or resume) else ""
                if dedup == "message-id":
                    dup_key = f"mid:{mid}" if mid else None
                    if dup_key and dup_key in seen_keys:
                        is_dup = True
                elif dedup == "sha256":
                    dup_key = f"sha256:{sha}" if sha else None
                    if dup_key and dup_key in seen_keys:
                        is_dup = True
                else:  # auto or resume
                    if resume and mid and f"mid:{mid}" in seen_keys:
                        is_dup = True
                        dup_key = f"mid:{mid}"
                    elif resume and sha and f"sha256:{sha}" in seen_keys:
                        is_dup = True
                        dup_key = f"sha256:{sha}"
                    elif dedup == "auto":
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
                        size_bytes=len(raw_bytes),
                        attachment_count=0,
                        notes="Skipped duplicate message (dedup/resume)",
                    ))
                continue

            if dup_key:
                seen_keys.add(dup_key)
                if mid:
                    seen_mid_signatures[mid] = (sha, subj_val, date_val)
                    if dedup in ("sha256", "auto") or resume:
                        sha = compute_sha256(raw_bytes) if not sha else sha
                        seen_keys.add(f"mid:{mid}")
                        seen_keys.add(f"sha256:{sha}")

            msg_flags_tuple = extract_message_flags(msg)
            flags = rec.flags if rec.flags is not None else msg_flags_tuple
            is_del = (
                rec.is_deleted
                or msg_flags_tuple[4]
                or (isinstance(flags, dict) and flags.get("is_deleted", False))
            )
            if is_del and not include_deleted:
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
                        size_bytes=len(raw_bytes),
                        attachment_count=0,
                        notes="Message marked deleted/expunged",
                    ))
                continue

            year, month = pick_year_month(msg)
            if (start_year and year < start_year) or (end_year and year > end_year):
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
                        size_bytes=len(raw_bytes),
                        attachment_count=0,
                        notes=f"Filtered by date range ({start_year}-{end_year})",
                    ))
                continue

            if single_file_mode:
                dest_mbox = target_path
            else:
                dest_mbox = resolve_mbox_dest_path(
                    target_path,
                    layout=layout,
                    folder_parts=rec.folder_parts,
                    year=year,
                    month=month,
                    default_name=src_stem,
                )

            try:
                entry_bytes = format_mboxrd_entry(raw_bytes, msg=msg, flags=flags)
                actual_dest = pool.write_entry(dest_mbox, entry_bytes)
                written += 1
                total_bytes += len(entry_bytes)
                if logger and logger.audit_log_path:
                    logger.record_message(AuditRecord(
                        timestamp=datetime.now(timezone.utc).isoformat(),
                        status="SUCCESS",
                        source_file=rec.source_name,
                        source_folder=folder_display,
                        target_location=actual_dest,
                        message_id=safe_get_header(msg, "Message-ID", ""),
                        subject=safe_get_header(msg, "Subject", "(no subject)"),
                        date=safe_get_header(msg, "Date", ""),
                        sender=safe_get_header(msg, "From", ""),
                        recipients="; ".join(safe_get_all_headers(msg, "To")),
                        size_bytes=len(entry_bytes),
                        attachment_count=0,
                        sha256=compute_sha256(raw_bytes) if logger.audit_log_path else "",
                    ))
            except Exception as e:
                skipped += 1
                failed += 1
                # Zero-Data-Loss Emergency Fallback: preserve unwritten bytes to disk
                fallback_base = target_path if not single_file_mode else os.path.dirname(target_path)
                fb_dir = os.path.join(fallback_base or ".", "_unwritten_messages")
                ensure_dir(fb_dir)
                subj = sanitize_name(safe_get_header(msg, "Subject", "msg") or "msg", max_len=60)
                fb_path = os.path.join(fb_dir, f"unwritten_{written + failed}_{subj}.eml")
                try:
                    with open(win_long_path(fb_path), "wb") as fb:
                        fb.write(raw_bytes)
                    if logger:
                        logger.warn(f"Emergency preservation: saved unwritten message to {fb_path}")
                except Exception:
                    pass

                if logger:
                    logger.error(f"Failed to write mbox entry to {dest_mbox}", exc=e)
                    logger.record_failure(
                        source_file=rec.source_name,
                        source_folder=folder_display,
                        message_id=safe_get_header(msg, "Message-ID", ""),
                        subject=safe_get_header(msg, "Subject", "(no subject)"),
                        error_type=type(e).__name__,
                        error_details=str(e),
                    )

            if written % 250 == 0:
                if logger and not logger.quiet:
                    sys.stdout.write(f"\rExported {written:,} messages (skipped {skipped:,})...")
                    sys.stdout.flush()
                elif not logger:
                    print(f"\rExported {written:,} messages (skipped {skipped:,})...", end="", flush=True)
    finally:
        pool.close_all()

    if written >= 250 and (not logger or not logger.quiet):
        print()

    if logger:
        logger.print_reconciliation_summary(
            source=src_path,
            target=target_path,
            start_time=started,
            total_scanned=idx,
            exported=written,
            skipped_date=skipped_date,
            skipped_deleted=skipped_deleted,
            failed=failed,
            total_bytes=total_bytes,
            skipped_duplicate=skipped_duplicate,
        )

    return written, skipped


def build_arg_parser(prog: Optional[str] = None) -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog=prog,
        description=(
            "mail2mbox — Convert any mail format (MBOX, Thunderbird ImapMail/Mail, "
            "Maildir/Maildir++, Apple Mail .emlx/.mbox, Outlook .msg/.olm/.pst, MH, MMDF, "
            "Babyl, EML, or .gz/.bz2/.xz/.zip/.tar* archives) into standard mboxrd .mbox file(s)."
        ),
    )
    ap.add_argument("src_pos", nargs="?", help="Source mail file, directory, or archive")
    ap.add_argument("out_pos", nargs="?", help="Output .mbox (.mbox.gz/.bz2/.xz) file or destination directory")
    ap.add_argument("--src", "--mbox", dest="src_opt", help="Source mail file, directory, or archive")
    ap.add_argument("--out", dest="out_file", help="Single output .mbox (.mbox.gz/.bz2/.xz) file path")
    ap.add_argument("--out-dir", dest="out_dir", help="Destination directory for multi-mbox output")
    ap.add_argument(
        "--layout",
        choices=["folder", "year", "month", "flat"],
        default="folder",
        help="MBOX directory organization when exporting to a folder (default: folder)",
    )
    ap.add_argument("--max-mbox-gb", type=float, default=0.0, help="Rotate to __part2.mbox when an mbox exceeds ~GB (0=unlimited)")
    ap.add_argument("--start-year", type=int, default=None, help="Only export messages from this year (inclusive)")
    ap.add_argument("--end-year", type=int, default=None, help="Only export messages up to this year (inclusive)")
    ap.add_argument("--include-deleted", action="store_true", help="Include messages marked deleted/expunged")
    ap.add_argument("--append", action="store_true", help="Append to existing .mbox file(s) instead of overwriting")
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
        "--no-gmail-labels",
        action="store_true",
        help="Disable automatic Google Takeout X-Gmail-Labels folder hierarchy reconstruction",
    )
    ap.add_argument("--audit-log", help="Path to write structured audit manifest (.csv or .jsonl)")
    ap.add_argument("--error-log", help="Path to write structured failure report (.jsonl)")
    ap.add_argument("--log-file", help="Path to write persistent operational log file")
    ap.add_argument("--dry-run", action="store_true", help="Perform pre-flight inspection and report without modifying disk")
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

    src = os.path.abspath(src)
    if not os.path.exists(win_long_path(src)):
        print(f"Error: Source path '{src}' does not exist.", file=sys.stderr)
        sys.exit(1)

    if args.dry_run:
        perform_dry_run(src)
        return

    target = args.out_file or args.out_dir or args.out_pos
    if not target:
        ap.error("Both a source path (positional or --src) and an output target (positional, --out, or --out-dir) are required.")

    target = os.path.abspath(target)

    single_file_mode = bool(args.out_file) or (not args.out_dir and is_mbox_output_filename(target))
    if not single_file_mode:
        os.makedirs(win_long_path(target), exist_ok=True)

    logger = AuditLogger(
        audit_log_path=args.audit_log,
        error_log_path=args.error_log,
        log_file_path=args.log_file,
        quiet=args.quiet,
        verbose=args.verbose,
    )

    try:
        logger.info(f"Starting mail2mbox export from: {src}")
        logger.info(f"Destination: {target} ({'single file' if single_file_mode else f'directory, layout={args.layout}'})")

        written, skipped = export_to_mbox(
            src_path=src,
            target_path=target,
            single_file_mode=single_file_mode,
            layout=args.layout,
            start_year=args.start_year,
            end_year=args.end_year,
            include_deleted=args.include_deleted,
            max_mbox_gb=args.max_mbox_gb,
            overwrite=(not args.append),
            logger=logger,
            dedup=args.dedup,
            resume=args.resume,
            parse_gmail=not args.no_gmail_labels,
        )
    finally:
        logger.close()


if __name__ == "__main__":
    main()
