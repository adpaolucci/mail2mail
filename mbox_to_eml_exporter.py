#!/usr/bin/env python3
"""
mbox_to_eml_exporter.py — Backward-compatible wrapper around mail2eml.py (part of mail2mail).
Prefer running `mail2eml.py` or `mail2mail.py eml` directly.
"""

from mail2eml import (
    MSG_FLAG_EXPUNGED,
    WINDOWS_RESERVED_NAMES,
    eml_bytes,
    ensure_dir,
    export_to_eml,
    fits_limits,
    get_header_str,
    is_thunderbird_expunged,
    main,
    pick_year_month,
    safe_name,
    unique_eml_name,
    win_long_path,
)

__all__ = [
    "MSG_FLAG_EXPUNGED",
    "WINDOWS_RESERVED_NAMES",
    "eml_bytes",
    "ensure_dir",
    "export_to_eml",
    "fits_limits",
    "get_header_str",
    "is_thunderbird_expunged",
    "main",
    "pick_year_month",
    "safe_name",
    "unique_eml_name",
    "win_long_path",
]

if __name__ == "__main__":
    main()
