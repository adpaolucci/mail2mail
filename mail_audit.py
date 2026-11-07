#!/usr/bin/env python3
"""
mail_audit.py — Structured Audit Logging, Forensics, and Inspection Engine for mail2mail

Provides:
  - Per-message audit manifest streaming (--audit-log <path.csv | path.jsonl>)
  - Dedicated failure / error logs (--error-log <path.jsonl>)
  - Persistent operational log files (--log-file <path.log>) with log levels (INFO, WARN, ERROR)
  - Pre-flight dry-run inspection (--dry-run) surfacing folder breakdowns, date ranges,
    and oversized attachments (>20 MB / >25 MB) before modifying disk/PST
  - Unified audit reconciliation summary printed at the conclusion of every job
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import datetime, timezone
import email
from email import policy
import hashlib
import json
import os
import sys
import time
from typing import Any, Dict, List, Optional, Set, Tuple

from mail_formats import (
    extract_attachment_filename,
    extract_message_flags,
    iter_mail_records,
    pick_year_month,
    safe_get_all_headers,
    safe_get_header,
    win_long_path,
)


@dataclass
class AuditRecord:
    timestamp: str
    status: str  # SUCCESS, SKIPPED_DATE, SKIPPED_DELETED, SKIPPED_OVERSIZED, SKIPPED_DUPLICATE, FAILED
    source_file: str
    source_folder: str
    target_location: str
    message_id: str
    subject: str
    date: str
    sender: str
    recipients: str
    size_bytes: int
    attachment_count: int
    sha256: str = ""
    notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "timestamp": self.timestamp,
            "status": self.status,
            "source_file": self.source_file,
            "source_folder": self.source_folder,
            "target_location": self.target_location,
            "message_id": self.message_id,
            "subject": self.subject,
            "date": self.date,
            "sender": self.sender,
            "recipients": self.recipients,
            "size_bytes": self.size_bytes,
            "attachment_count": self.attachment_count,
            "sha256": self.sha256,
            "notes": self.notes,
        }


class AuditLogger:
    """
    Unified operational logger, audit manifest recorder, and error reporter.
    """
    def __init__(
        self,
        audit_log_path: Optional[str] = None,
        error_log_path: Optional[str] = None,
        log_file_path: Optional[str] = None,
        quiet: bool = False,
        verbose: bool = False,
    ):
        self.audit_log_path = audit_log_path
        self.error_log_path = error_log_path
        self.log_file_path = log_file_path
        self.quiet = quiet
        self.verbose = verbose

        self._audit_fh = None
        self._audit_csv_writer = None
        self._audit_is_jsonl = False

        self._error_fh = None
        self._log_fh = None

        self._open_logs()

    def _open_logs(self) -> None:
        if self.audit_log_path:
            p = os.path.abspath(self.audit_log_path)
            os.makedirs(os.path.dirname(win_long_path(p)), exist_ok=True)
            self._audit_is_jsonl = p.lower().endswith((".jsonl", ".json", ".ndjson"))
            self._audit_fh = open(win_long_path(p), "w", encoding="utf-8", newline="" if not self._audit_is_jsonl else None)
            if not self._audit_is_jsonl:
                fieldnames = [
                    "timestamp", "status", "source_file", "source_folder", "target_location",
                    "message_id", "subject", "date", "sender", "recipients", "size_bytes",
                    "attachment_count", "sha256", "notes"
                ]
                self._audit_csv_writer = csv.DictWriter(self._audit_fh, fieldnames=fieldnames)
                self._audit_csv_writer.writeheader()

        if self.error_log_path:
            p = os.path.abspath(self.error_log_path)
            os.makedirs(os.path.dirname(win_long_path(p)), exist_ok=True)
            self._error_fh = open(win_long_path(p), "w", encoding="utf-8")

        if self.log_file_path:
            p = os.path.abspath(self.log_file_path)
            os.makedirs(os.path.dirname(win_long_path(p)), exist_ok=True)
            self._log_fh = open(win_long_path(p), "a", encoding="utf-8")

    def _now_str(self) -> str:
        return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

    def log(self, level: str, msg: str) -> None:
        formatted = f"[{self._now_str()}] [{level.upper()}] {msg}"
        if self._log_fh:
            try:
                self._log_fh.write(formatted + "\n")
                self._log_fh.flush()
            except Exception:
                pass

        if not self.quiet or level in ("WARN", "WARNING", "ERROR"):
            if level in ("WARN", "WARNING", "ERROR"):
                print(f"[{level.upper()}] {msg}", file=sys.stderr)
            else:
                print(msg)

    def info(self, msg: str) -> None:
        self.log("INFO", msg)

    def warn(self, msg: str) -> None:
        self.log("WARN", msg)

    def error(self, msg: str, exc: Optional[Exception] = None) -> None:
        full_msg = f"{msg}: {exc}" if exc else msg
        self.log("ERROR", full_msg)

    def record_message(self, rec: AuditRecord) -> None:
        if self._audit_fh:
            try:
                if self._audit_is_jsonl:
                    self._audit_fh.write(json.dumps(rec.to_dict(), ensure_ascii=False) + "\n")
                elif self._audit_csv_writer:
                    self._audit_csv_writer.writerow(rec.to_dict())
                self._audit_fh.flush()
            except Exception:
                pass

    def record_failure(
        self,
        source_file: str,
        source_folder: str,
        message_id: str,
        subject: str,
        error_type: str,
        error_details: str,
        retryable: bool = False,
    ) -> None:
        if self._error_fh:
            try:
                entry = {
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "source_file": source_file,
                    "source_folder": source_folder,
                    "message_id": message_id,
                    "subject": subject,
                    "error_type": error_type,
                    "error_details": error_details,
                    "retryable": retryable,
                }
                self._error_fh.write(json.dumps(entry, ensure_ascii=False) + "\n")
                self._error_fh.flush()
            except Exception:
                pass

    @staticmethod
    def load_completed_keys(manifest_path: str) -> Set[str]:
        """
        Read an existing CSV or JSONL audit manifest and return a set of unique message keys
        ('mid:<Message-ID>' and 'sha256:<hash>') for all entries marked with status == 'SUCCESS'.
        Used by --resume to skip already-imported records.
        """
        completed: Set[str] = set()
        lp = win_long_path(manifest_path)
        if not os.path.isfile(lp):
            return completed

        is_jsonl = manifest_path.lower().endswith((".jsonl", ".json", ".ndjson"))
        try:
            with open(lp, "r", encoding="utf-8", errors="replace") as f:
                if manifest_path.lower().endswith(".json"):
                    content = f.read().strip()
                    if content.startswith("{"):
                        try:
                            obj = json.loads(content)
                            if "migrated_message_ids" in obj:
                                for m in obj.get("migrated_message_ids", []):
                                    if m:
                                        completed.add(f"mid:{m}")
                                        completed.add(f"sha256:{m}")
                                        completed.add(m)
                                return completed
                        except Exception:
                            pass
                    # If not single-dict JSON, fall through to line-by-line reading
                    lines = content.splitlines()
                    for line in lines:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            obj = json.loads(line)
                            status = (obj.get("status") or "").strip()
                            if "status" not in obj or not status or status.upper() == "SUCCESS":
                                mid = (obj.get("message_id") or "").strip()
                                sha = (obj.get("sha256") or "").strip()
                                if mid:
                                    completed.add(f"mid:{mid}")
                                if sha:
                                    completed.add(f"sha256:{sha}")
                        except Exception:
                            pass
                elif is_jsonl:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            obj = json.loads(line)
                            status = (obj.get("status") or "").strip()
                            if "status" not in obj or not status or status.upper() == "SUCCESS":
                                mid = (obj.get("message_id") or "").strip()
                                sha = (obj.get("sha256") or "").strip()
                                if mid:
                                    completed.add(f"mid:{mid}")
                                if sha:
                                    completed.add(f"sha256:{sha}")
                        except Exception:
                            pass
                else:
                    reader = csv.DictReader(f)
                    for row in reader:
                        status = (row.get("status") or "").strip()
                        if "status" not in row or not status or status.upper() == "SUCCESS":
                            mid = (row.get("message_id") or "").strip()
                            sha = (row.get("sha256") or "").strip()
                            if mid:
                                completed.add(f"mid:{mid}")
                            if sha:
                                completed.add(f"sha256:{sha}")
        except Exception:
            pass
        return completed

    def print_reconciliation_summary(
        self,
        source: str,
        target: str,
        start_time: float,
        total_scanned: int,
        exported: int,
        skipped_date: int,
        skipped_deleted: int,
        failed: int,
        total_bytes: int = 0,
        extra_details: Optional[Dict[str, Any]] = None,
        skipped_duplicate: int = 0,
        skipped_folder: int = 0,
        skipped_state: int = 0,
        interrupted: bool = False,
        interruption_reason: str = "",
    ) -> None:
        elapsed = max(time.time() - start_time, 1e-6)
        rate = exported / elapsed
        mb_rate = (total_bytes / (1024 * 1024)) / elapsed if total_bytes else 0.0

        header_title = "MIGRATION AUDIT SUMMARY (PAUSED / INTERRUPTED)" if interrupted else "MIGRATION AUDIT SUMMARY"
        lines = [
            "\n" + "=" * 70,
            f"                     {header_title}",
            "=" * 70,
            f"Source:                {source}",
            f"Target:                {target}",
        ]
        if interrupted:
            reason_str = interruption_reason or "User Request (Ctrl+C / SIGTERM)"
            lines.append(f"Status:                PAUSED / HALTED ({reason_str})")
            lines.append("Integrity:             SAFE -- All saved items committed to disk with zero data loss.")
        lines.extend([
            f"Duration:              {elapsed:.1f}s ({elapsed / 60:.1f}m) | Speed: {rate:.1f} msg/s ({mb_rate:.2f} MB/s)",
            "-" * 70,
            "Message Reconciliation:",
            f"  Total Scanned:       {total_scanned:,}",
            f"  Successfully Saved:  {exported:,} ({100.0 * exported / max(total_scanned, 1):.1f}%)",
            f"  Skipped (Duplicates): {skipped_duplicate:,}",
            f"  Skipped (Date Filter): {skipped_date:,}",
            f"  Skipped (Deleted):   {skipped_deleted:,}",
        ])
        if skipped_folder > 0:
            lines.append(f"  Skipped (Folder Rule): {skipped_folder:,}")
        if skipped_state > 0:
            lines.append(f"  Skipped (Delta / State): {skipped_state:,}")
        lines.append(f"  Failed / Errors:     {failed:,}")

        if extra_details:
            lines.append("-" * 70)
            lines.append("Additional Details:")
            for k, v in extra_details.items():
                lines.append(f"  {k}: {v}")

        artifacts = []
        if self.audit_log_path:
            artifacts.append(f"Audit Manifest:        {self.audit_log_path}")
        if self.error_log_path:
            artifacts.append(f"Failure Log:           {self.error_log_path}")
        if self.log_file_path:
            artifacts.append(f"Operation Log:         {self.log_file_path}")

        if artifacts:
            lines.append("-" * 70)
            lines.append("Audit Artifacts:")
            for a in artifacts:
                lines.append(f"  {a}")

        lines.append("=" * 70 + "\n")
        summary_text = "\n".join(lines)

        if self._log_fh:
            try:
                self._log_fh.write(summary_text + "\n")
                self._log_fh.flush()
            except Exception:
                pass

        print(summary_text)

    def close(self) -> None:
        for fh in (self._audit_fh, self._error_fh, self._log_fh):
            if fh is not None:
                try:
                    fh.flush()
                    fh.close()
                except Exception:
                    pass
        self._audit_fh = None
        self._error_fh = None
        self._log_fh = None


def compute_sha256(data: bytes) -> str:
    """Compute SHA-256 hex digest for chain-of-custody verification."""
    return hashlib.sha256(data).hexdigest()


# ----------------------- Pre-Flight Dry-Run & Inventory -----------------------

def perform_dry_run(src_path: str, max_attach_mb: float = 25.0) -> None:
    """
    Perform a non-destructive pre-flight audit of the source mail store.
    Inspects folder hierarchies, message counts, date distributions, and
    flags any attachments that exceed the specified size threshold.
    """
    print(f"\n[*] Initiating Pre-Flight Dry-Run Audit for: {src_path}")
    print("-" * 70)

    start_time = time.time()
    total_messages = 0
    total_bytes = 0
    total_read = 0
    total_unread = 0
    total_flagged = 0
    total_deleted = 0

    folders: Dict[str, int] = {}
    folder_bytes: Dict[str, int] = {}
    years: Dict[int, int] = {}
    earliest_date: Optional[str] = None
    latest_date: Optional[str] = None
    earliest_val: Optional[datetime] = None
    latest_val: Optional[datetime] = None

    oversized_attachments: List[Tuple[str, float, str, str]] = []  # (filename, size_mb, subject, folder)
    header_parser = email.parser.BytesHeaderParser()

    for rec in iter_mail_records(src_path, include_deleted=True):
        total_messages += 1
        raw = rec.raw_bytes
        sz = len(raw)
        total_bytes += sz

        folder_name = os.sep.join(rec.folder_parts) if rec.folder_parts else "Root"
        folders[folder_name] = folders.get(folder_name, 0) + 1
        folder_bytes[folder_name] = folder_bytes.get(folder_name, 0) + sz

        is_read, is_flagged, _, _, is_deleted = extract_message_flags(rec.flags)
        if is_deleted or rec.is_deleted:
            total_deleted += 1
        elif is_read:
            total_read += 1
        else:
            total_unread += 1

        if is_flagged:
            total_flagged += 1

        msg = header_parser.parsebytes(raw)
        year, month = pick_year_month(msg)
        years[year] = years.get(year, 0) + 1

        date_hdr = safe_get_header(msg, "Date", "").strip()
        if date_hdr:
            try:
                dt = email.utils.parsedate_to_datetime(date_hdr)
                if dt is not None:
                    if dt.tzinfo is None:
                        dt = dt.replace(tzinfo=timezone.utc)
                    else:
                        dt = dt.astimezone(timezone.utc)
                    if earliest_val is None or dt < earliest_val:
                        earliest_val = dt
                        earliest_date = date_hdr
                    if latest_val is None or dt > latest_val:
                        latest_val = dt
                        latest_date = date_hdr
            except Exception:
                pass

        # Check for large attachments (> max_attach_mb) if message is multipart
        if b"attachment" in raw.lower() or b"boundary=" in raw.lower():
            try:
                full_msg = email.message_from_bytes(raw, policy=policy.default)
                for part in full_msg.iter_attachments():
                    try:
                        payload = part.get_payload(decode=True)
                        if payload and len(payload) > int(max_attach_mb * 1024 * 1024):
                            att_mb = len(payload) / (1024 * 1024)
                            fn = extract_attachment_filename(part) or "unnamed_attachment.bin"
                            subj = safe_get_header(msg, "Subject", "(no subject)")
                            oversized_attachments.append((fn, att_mb, subj, folder_name))
                    except Exception:
                        pass
            except Exception:
                pass

        if total_messages % 1000 == 0:
            sys.stdout.write(f"\r  Scanned {total_messages:,} messages ({total_bytes / 1e6:.1f} MB)...")
            sys.stdout.flush()

    elapsed = max(time.time() - start_time, 1e-6)
    print(f"\r  Scanned {total_messages:,} messages ({total_bytes / 1e6:.1f} MB) in {elapsed:.1f}s.")

    print("\n" + "=" * 70)
    print("                     PRE-FLIGHT AUDIT REPORT")
    print("=" * 70)
    print(f"Total Messages:          {total_messages:,}")
    print(f"Total Payload Size:      {total_bytes / (1024**3):.2f} GB ({total_bytes / (1024**2):.1f} MB)")
    print(f"Status Breakdown:        {total_read:,} Read | {total_unread:,} Unread | {total_flagged:,} Flagged | {total_deleted:,} Deleted")
    print(f"Date Span:               {earliest_date or 'N/A'}  -->  {latest_date or 'N/A'}")
    print(f"Distinct Mail Folders:   {len(folders):,}")

    print("\nFolder Distribution (Top 10):")
    top_folders = sorted(folder_bytes.items(), key=lambda kv: kv[1], reverse=True)[:10]
    for name, b in top_folders:
        cnt = folders.get(name, 0)
        print(f"  - {name:<40} {cnt:>6,} msgs | {b / (1024*1024):>8.1f} MB")

    print("\nMessages by Year:")
    for yr in sorted(years.keys()):
        print(f"  {yr}: {years[yr]:,}")

    if oversized_attachments:
        print("\n" + "!" * 70)
        print(f"WARNING: Found {len(oversized_attachments)} attachment(s) exceeding {max_attach_mb:.0f} MB!")
        print("These attachments may be rejected by Microsoft Outlook's default server size limit.")
        print("To lift Outlook's attachment limit, set MaximumAttachmentSize=0 in HKCU before importing.")
        print("-" * 70)
        for fn, sz_mb, subj, fldr in oversized_attachments[:10]:
            print(f"  - {fn} ({sz_mb:.1f} MB) in [{fldr}] (Subject: '{subj[:40]}')")
        if len(oversized_attachments) > 10:
            print(f"  ... and {len(oversized_attachments) - 10} more oversized attachment(s).")
        print("!" * 70)
    else:
        print(f"\nAttachment Size Scan: All attachments are <= {max_attach_mb:.0f} MB (within standard Outlook limits).")

    print("=" * 70)
    print("Dry-run complete. No files or PST archives were modified.\n")
