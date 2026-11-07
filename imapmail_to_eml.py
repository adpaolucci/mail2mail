#!/usr/bin/env python3
"""
imapmail_to_eml.py — Backward-compatible wrapper around mail2eml.py (part of mail2mail).
Prefer running `mail2eml.py` or `mail2mail.py eml` directly.
"""

import os
import sys

from mail2eml import (
    FORBIDDEN_CHARS,
    IGNORED_DIR_SUFFIXES,
    MSG_FLAG_EXPUNGED,
    MSG_FLAG_FORWARDED,
    MSG_FLAG_MARKED,
    MSG_FLAG_READ,
    MSG_FLAG_REPLIED,
    WINDOWS_RESERVED_NAMES,
    decode_modified_utf7,
    export_mbox_stream,
    export_multi_format_source,
    is_probable_mbox,
    main as _mail2eml_main,
    normalize_eml_bytes_with_flags,
    parse_thunderbird_flags,
    sanitize_filename,
    traverse_and_export,
    update_status_headers,
    win_long_path,
)

__all__ = [
    "FORBIDDEN_CHARS",
    "IGNORED_DIR_SUFFIXES",
    "MSG_FLAG_EXPUNGED",
    "MSG_FLAG_FORWARDED",
    "MSG_FLAG_MARKED",
    "MSG_FLAG_READ",
    "MSG_FLAG_REPLIED",
    "WINDOWS_RESERVED_NAMES",
    "decode_modified_utf7",
    "export_mbox_stream",
    "export_multi_format_source",
    "is_probable_mbox",
    "main",
    "normalize_eml_bytes_with_flags",
    "parse_thunderbird_flags",
    "sanitize_filename",
    "traverse_and_export",
    "update_status_headers",
    "win_long_path",
]


def main(argv=None):
    args = list(sys.argv[1:] if argv is None else argv)
    if "--naming" not in args and not any(a.startswith("--naming=") for a in args):
        args.extend(["--naming", "index"])
    if "--layout" not in args and not any(a.startswith("--layout=") for a in args):
        args.extend(["--layout", "folder"])
    _mail2eml_main(args, prog=os.path.basename(sys.argv[0]))


if __name__ == "__main__":
    main()
