#!/usr/bin/env python3
"""
mail_filter.py — Advanced Filtering, Date Normalization, Size Parsing, and
Persistent Migration State Tracking for mail2mail (mail2pst, mail2eml, mail2mbox).

Provides:
  1. DateFilter: Granular date-range filtering (--since / --after, --until / --before,
     --start-year, --end-year) with timezone-safe UTC normalization.
  2. FolderFilter: Folder inclusion/exclusion filtering (--include-folder, --exclude-folder)
     supporting wildcards (glob / fnmatch) and hierarchy segment matching.
  3. SizeParser: Human-readable size parsing (e.g. '20GB', '500MB', '15.5G') for PST,
     MBOX, and directory splitting.
  4. MigrationState: Persistent SQLite / JSONL migration state cache (--state-file)
     enabling true incremental / delta cutover sync across repeated runs.
"""

from __future__ import annotations

from datetime import datetime, time, timezone
import email.utils
import fnmatch
import json
import os
import re
import sqlite3
from typing import Any, Dict, List, Optional, Set, Tuple


# ----------------------- Size Parsing -----------------------

_SIZE_UNIT_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*([a-zA-Z]*)\s*$")

def parse_size_bytes(val: Any, default_unit: str = "GB") -> int:
    """
    Parse a size string or numeric value into integer bytes.
    Examples:
      '20GB' -> 21,474,836,480
      '500MB' -> 524,288,000
      '15.5G' -> 16,642,998,272
      15.0    -> 16,106,127,360 (default_unit='GB')
      0       -> 0 (unlimited)
    """
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        if val <= 0:
            return 0
        if default_unit.upper() in ("GB", "G"):
            return int(val * (1024**3))
        if default_unit.upper() in ("MB", "M"):
            return int(val * (1024**2))
        return int(val)

    s = str(val).strip()
    if not s or s in ("0", "none", "unlimited", "off"):
        return 0

    m = _SIZE_UNIT_RE.match(s)
    if not m:
        try:
            return int(float(s) * (1024**3))
        except ValueError:
            raise ValueError(f"Invalid size specification: '{val}'")

    num_str, unit_str = m.groups()
    num = float(num_str)
    unit = unit_str.upper() if unit_str else default_unit.upper()

    if unit in ("B", "BYTES", ""):
        return int(num)
    elif unit in ("K", "KB", "KIB"):
        return int(num * 1024)
    elif unit in ("M", "MB", "MIB"):
        return int(num * (1024**2))
    elif unit in ("G", "GB", "GIB"):
        return int(num * (1024**3))
    elif unit in ("T", "TB", "TIB"):
        return int(num * (1024**4))
    else:
        raise ValueError(f"Unrecognized size unit in '{val}': '{unit_str}'")


# ----------------------- Date Parsing & Filtering -----------------------

def parse_cli_datetime(val: Any, is_end_of_day: bool = False) -> Optional[datetime]:
    """
    Parse a user CLI date/datetime string into a timezone-aware UTC datetime.
    Supports:
      - YYYY-MM-DD
      - YYYY-MM-DD HH:MM[:SS]
      - YYYY-MM-DDTHH:MM[:SS]
      - RFC 2822 / 5322 date strings (e.g. '15 Jan 2023 12:00:00 +0000')
      - YYYY (year only)
    When `is_end_of_day=True` and no time component was given, defaults to 23:59:59.999999 UTC.
    """
    if not val:
        return None
    if isinstance(val, datetime):
        if val.tzinfo is None:
            return val.replace(tzinfo=timezone.utc)
        return val.astimezone(timezone.utc)

    s = str(val).strip()
    if not s:
        return None

    # 1. Year only: '2023'
    if re.fullmatch(r"\d{4}", s):
        year = int(s)
        if is_end_of_day:
            return datetime(year, 12, 31, 23, 59, 59, 999999, tzinfo=timezone.utc)
        return datetime(year, 1, 1, 0, 0, 0, 0, tzinfo=timezone.utc)

    # 2. Date only: '2023-05-15' or '2023/05/15'
    date_m = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if date_m:
        y, m, d = int(date_m.group(1)), int(date_m.group(2)), int(date_m.group(3))
        if is_end_of_day:
            return datetime(y, m, d, 23, 59, 59, 999999, tzinfo=timezone.utc)
        return datetime(y, m, d, 0, 0, 0, 0, tzinfo=timezone.utc)

    # 3. ISO Datetime: '2023-05-15T14:30:00' or '2023-05-15 14:30:00'
    iso_clean = s.replace(" ", "T")
    try:
        dt = datetime.fromisoformat(iso_clean)
        if dt.tzinfo is None:
            return dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        pass

    # 4. RFC 2822 / 5322 fallback
    try:
        dt = email.utils.parsedate_to_datetime(s)
        if dt is not None:
            if dt.tzinfo is None:
                return dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
    except Exception:
        pass

    raise ValueError(f"Unable to parse date/time string: '{val}'. Supported formats: YYYY-MM-DD, YYYY-MM-DDTHH:MM:SS, RFC 2822.")


class DateFilter:
    """
    Timezone-aware date filter checking messages against --since, --until,
    --start-year, and --end-year bounds.
    """
    def __init__(
        self,
        since: Optional[Any] = None,
        until: Optional[Any] = None,
        start_year: Optional[int] = None,
        end_year: Optional[int] = None,
    ):
        self.since_dt: Optional[datetime] = parse_cli_datetime(since, is_end_of_day=False)
        self.until_dt: Optional[datetime] = parse_cli_datetime(until, is_end_of_day=True)
        self.start_year = start_year
        self.end_year = end_year

        # If start_year was given and since_dt was not, unify into since_dt
        if self.start_year and not self.since_dt:
            self.since_dt = datetime(self.start_year, 1, 1, 0, 0, 0, 0, tzinfo=timezone.utc)
        # If end_year was given and until_dt was not, unify into until_dt
        if self.end_year and not self.until_dt:
            self.until_dt = datetime(self.end_year, 12, 31, 23, 59, 59, 999999, tzinfo=timezone.utc)

    @property
    def is_active(self) -> bool:
        return self.since_dt is not None or self.until_dt is not None

    def matches(self, dt: Optional[datetime], fallback_year: Optional[int] = None) -> bool:
        """Return True if dt (or fallback_year) is within the configured window."""
        if not self.is_active:
            return True

        if dt is None:
            if fallback_year is not None:
                if self.since_dt and fallback_year < self.since_dt.year:
                    return False
                if self.until_dt and fallback_year > self.until_dt.year:
                    return False
                return True
            # No date information at all when filtering is requested -> skip
            return False

        # Normalize dt to UTC
        if dt.tzinfo is None:
            dt_utc = dt.replace(tzinfo=timezone.utc)
        else:
            dt_utc = dt.astimezone(timezone.utc)

        if self.since_dt and dt_utc < self.since_dt:
            return False
        if self.until_dt and dt_utc > self.until_dt:
            return False

        return True

    def describe(self) -> str:
        parts = []
        if self.since_dt:
            parts.append(f"since {self.since_dt.isoformat()}")
        if self.until_dt:
            parts.append(f"until {self.until_dt.isoformat()}")
        return " and ".join(parts) if parts else "all dates"


# ----------------------- Folder Filtering -----------------------

class FolderFilter:
    """
    Folder filter supporting both inclusion and exclusion rules with glob wildcards.
    Matches against full relative folder paths (e.g. '[Gmail]/Trash') as well as
    individual path segments (e.g. 'Trash').
    """
    def __init__(
        self,
        include_patterns: Optional[List[str]] = None,
        exclude_patterns: Optional[List[str]] = None,
    ):
        self.include_patterns: List[str] = [p.strip().replace("\\", "/") for p in (include_patterns or []) if p.strip()]
        self.exclude_patterns: List[str] = [p.strip().replace("\\", "/") for p in (exclude_patterns or []) if p.strip()]

    @property
    def is_active(self) -> bool:
        return bool(self.include_patterns or self.exclude_patterns)

    def _matches_pattern(self, target: str, pattern: str) -> bool:
        t_low = target.lower()
        p_low = pattern.lower()
        # Direct equality or glob match
        if t_low == p_low or fnmatch.fnmatch(t_low, p_low):
            return True
        # If pattern has no slash, match basename or individual segments
        if "/" not in pattern and "\\" not in pattern:
            segments = [s.lower() for s in re.split(r"[/]+", target) if s]
            if any(seg == p_low or fnmatch.fnmatch(seg, p_low) for seg in segments):
                return True
        return False

    def matches(self, folder_parts: List[str]) -> bool:
        """
        Return True if folder_parts passes inclusion and exclusion rules.
        """
        if not self.is_active:
            return True

        folder_str = "/".join(p for p in folder_parts if p) or "Inbox"

        # 1. Inclusion check: if include patterns exist, must match at least one
        if self.include_patterns:
            matched_include = False
            for pat in self.include_patterns:
                if self._matches_pattern(folder_str, pat):
                    matched_include = True
                    break
            if not matched_include:
                return False

        # 2. Exclusion check: if matches any exclude pattern, reject
        if self.exclude_patterns:
            for pat in self.exclude_patterns:
                if self._matches_pattern(folder_str, pat):
                    return False

        return True

    def describe(self) -> str:
        parts = []
        if self.include_patterns:
            parts.append(f"include: {', '.join(self.include_patterns)}")
        if self.exclude_patterns:
            parts.append(f"exclude: {', '.join(self.exclude_patterns)}")
        return "; ".join(parts) if parts else "all folders"


# ----------------------- Migration State Cache -----------------------

class MigrationState:
    """
    Persistent migration state tracking engine supporting SQLite database
    or JSON Lines checkpoint streams. Enables high-speed incremental migrations
    by recording processed Message-IDs and SHA-256 hashes.
    """
    def __init__(self, state_path: Optional[str] = None):
        self.state_path = os.path.abspath(state_path) if state_path else None
        self._is_sqlite = False
        self._is_jsonl = False
        self._db: Optional[sqlite3.Connection] = None
        self._jsonl_fh = None
        self._seen_keys: Set[str] = set()
        self._pending_records: List[Tuple[str, str, str, str, str, str]] = []
        self._batch_size = 500

        if self.state_path:
            os.makedirs(os.path.dirname(self.state_path) or ".", exist_ok=True)
            low = self.state_path.lower()
            if low.endswith((".sqlite", ".sqlite3", ".db")):
                self._is_sqlite = True
                self._init_sqlite()
            else:
                self._is_jsonl = True
                self._init_jsonl()

    @property
    def is_active(self) -> bool:
        return self.state_path is not None

    def _init_sqlite(self) -> None:
        self._db = sqlite3.connect(self.state_path, timeout=30.0)
        self._db.execute("PRAGMA journal_mode=WAL;")
        self._db.execute("PRAGMA synchronous=NORMAL;")
        self._db.execute(
            """
            CREATE TABLE IF NOT EXISTS completed_records (
                key TEXT PRIMARY KEY,
                message_id TEXT,
                sha256 TEXT,
                source_file TEXT,
                folder TEXT,
                timestamp TEXT
            );
            """
        )
        self._db.execute("CREATE INDEX IF NOT EXISTS idx_mid ON completed_records(message_id);")
        self._db.execute("CREATE INDEX IF NOT EXISTS idx_sha ON completed_records(sha256);")
        self._db.commit()

        # Pre-load keys into memory set for instant O(1) checks during streaming
        cursor = self._db.execute("SELECT key, message_id, sha256 FROM completed_records;")
        for k, mid, sha in cursor:
            if k:
                self._seen_keys.add(k)
            if mid:
                self._seen_keys.add(f"mid:{mid}")
            if sha:
                self._seen_keys.add(f"sha256:{sha}")

    def _init_jsonl(self) -> None:
        if os.path.isfile(self.state_path):
            with open(self.state_path, "r", encoding="utf-8", errors="replace") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        obj = json.loads(line)
                        k = obj.get("key")
                        if k:
                            self._seen_keys.add(k)
                        mid = obj.get("message_id")
                        if mid:
                            self._seen_keys.add(f"mid:{mid}")
                        sha = obj.get("sha256")
                        if sha:
                            self._seen_keys.add(f"sha256:{sha}")
                    except Exception:
                        pass
        self._jsonl_fh = open(self.state_path, "a", encoding="utf-8", errors="replace", buffering=1)

    def is_completed(self, mid: str = "", sha256: str = "") -> bool:
        """Check if message has already been successfully migrated in a prior pass."""
        if not self.is_active:
            return False
        if mid and f"mid:{mid}" in self._seen_keys:
            return True
        if sha256 and f"sha256:{sha256}" in self._seen_keys:
            return True
        return False

    def record_success(
        self,
        mid: str = "",
        sha256: str = "",
        source_file: str = "",
        folder: str = "",
    ) -> None:
        """Mark a message as successfully migrated and stage for persistence."""
        if not self.is_active:
            return

        ts = datetime.now(timezone.utc).isoformat()
        primary_key = f"mid:{mid}" if mid else (f"sha256:{sha256}" if sha256 else "")
        if not primary_key:
            return

        if mid:
            self._seen_keys.add(f"mid:{mid}")
        if sha256:
            self._seen_keys.add(f"sha256:{sha256}")

        if self._is_sqlite and self._db:
            self._pending_records.append((primary_key, mid, sha256, source_file, folder, ts))
            if len(self._pending_records) >= self._batch_size:
                self.flush()
        elif self._is_jsonl and self._jsonl_fh:
            rec = {
                "key": primary_key,
                "message_id": mid,
                "sha256": sha256,
                "source_file": source_file,
                "folder": folder,
                "timestamp": ts,
            }
            self._jsonl_fh.write(json.dumps(rec) + "\n")

    def flush(self) -> None:
        """Commit all pending records to disk."""
        if self._is_sqlite and self._db and self._pending_records:
            try:
                self._db.executemany(
                    """
                    INSERT OR IGNORE INTO completed_records
                    (key, message_id, sha256, source_file, folder, timestamp)
                    VALUES (?, ?, ?, ?, ?, ?);
                    """,
                    self._pending_records,
                )
                self._db.commit()
            except Exception:
                pass
            self._pending_records.clear()
        elif self._is_jsonl and self._jsonl_fh:
            try:
                self._jsonl_fh.flush()
            except Exception:
                pass

    def close(self) -> None:
        """Flush and close database / stream handles."""
        self.flush()
        if self._db:
            try:
                self._db.close()
            except Exception:
                pass
            self._db = None
        if self._jsonl_fh:
            try:
                self._jsonl_fh.close()
            except Exception:
                pass
            self._jsonl_fh = None

    def __bool__(self) -> bool:
        return True

    def __len__(self) -> int:
        return len(self._seen_keys)
