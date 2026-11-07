#!/usr/bin/env python3
"""
mail2mail.py — Unified Universal Email & Mailbox Converter

Converts between any supported email/mailbox format (MBOX, Thunderbird ImapMail/Mail,
Maildir/Maildir++, Apple Mail .emlx/.mbox, Outlook .msg, Outlook for Mac .olm,
Outlook .pst, MH, MMDF, Babyl, EML, and .gz/.bz2/.xz/.zip/.tar* archives) and:
  - .eml directory hierarchies (`mail2mail eml ...` or `mail2eml`)
  - .mbox / .mbox.gz / .mbox.bz2 / .mbox.xz files (`mail2mail mbox ...` or `mail2mbox`)
  - Microsoft Outlook .pst archives (`mail2mail pst ...` or `mail2pst`)

Examples:
  # Subcommand syntax
  python mail2mail.py eml "D:\\Mail\\inbox.mbox.gz" "D:\\Export_EML" --layout month
  python mail2mail.py mbox "D:\\Export_EML" "D:\\Backups\\AllMail.mbox.gz"
  python mail2mail.py pst "D:\\Mail\\ImapMail" "D:\\PSTs\\Mailbox.pst"
  python mail2mail.py pst --src "D:\\Export_EML" --out-dir "D:\\PSTs" --split-by year --max-pst-gb 15

  # Auto-detected target syntax (inferred from destination extension or --to)
  python mail2mail.py "D:\\Mail\\export.olm" "D:\\PSTs\\MacMail.pst"
  python mail2mail.py "D:\\PSTs\\MacMail.pst" "D:\\Backups\\MacMail.mbox.gz"
  python mail2mail.py "D:\\PSTs\\MacMail.pst" "D:\\Export_EML"
"""

from __future__ import annotations

import sys
from typing import List, Optional, Tuple

import mail2eml
import mail2mbox


PST_EXCLUSIVE_FLAGS = {
    "--pst",
    "--split-by",
    "--splits",
    "--max-pst-gb",
    "--pst-root",
    "--flatten",
    "--overwrite-pst",
    "--flush-every",
    "--count-every",
}

MBOX_EXCLUSIVE_FLAGS = {
    "--max-mbox-gb",
    "--append",
}

EML_EXCLUSIVE_FLAGS = {
    "--naming",
    "--max-per-dir",
    "--max-dir-bytes",
    "--sanitize-filenames",
    "--progress-every",
}

FLAGS_WITH_VALUE = {
    "--src",
    "--mbox",
    "--out",
    "--pst",
    "--out-dir",
    "--base-name",
    "--split-by",
    "--splits",
    "--max-pst-gb",
    "--max-mbox-gb",
    "--pst-root",
    "--layout",
    "--start-year",
    "--end-year",
    "--flush-every",
    "--count-every",
    "--naming",
    "--max-per-dir",
    "--max-dir-bytes",
    "--progress-every",
}


def _print_top_help() -> None:
    print(
        "usage: mail2mail {eml,mbox,pst} [options] ...\n"
        "   or: mail2mail <src> <dest.pst | dest.mbox[.gz] | out_dir> [--to {eml,mbox,pst}] [options]\n\n"
        "mail2mail — Universal Email & Mailbox Converter (MBOX, Thunderbird ImapMail/Mail,\n"
        "Maildir/Maildir++, Apple Mail .emlx/.mbox, Outlook .msg/.olm/.pst, MH, MMDF, Babyl,\n"
        "EML, and .gz/.bz2/.xz/.zip/.tar* archives <-> EML, MBOX, or Outlook PST).\n\n"
        "Commands:\n"
        "  eml  (or mail2eml)    Convert any supported mail source into .eml files (No Outlook required)\n"
        "  mbox (or mail2mbox)   Convert any supported mail source into .mbox / .mbox.gz files\n"
        "  pst  (or mail2pst)    Convert/import any supported mail source into Outlook .pst file(s)\n\n"
        "Run `mail2mail eml --help`, `mail2mail mbox --help`, or `mail2mail pst --help` for full options."
    )


def _infer_mode(argv: List[str]) -> Tuple[str, List[str]]:
    """
    Determine whether argv targets 'eml', 'mbox', or 'pst' and strip any explicit
    subcommand ('eml'/'mbox'/'pst') or '--to {eml,mbox,pst}' flag.
    """
    if not argv or argv[0] in ("-h", "--help"):
        return "help", argv

    first = argv[0].lower()
    if first in ("eml", "to-eml", "mail2eml"):
        return "eml", argv[1:]
    if first in ("mbox", "to-mbox", "mail2mbox"):
        return "mbox", argv[1:]
    if first in ("pst", "to-pst", "mail2pst"):
        return "pst", argv[1:]

    # Check for explicit --to {eml,mbox,pst}
    cleaned: List[str] = []
    explicit_to: Optional[str] = None
    i = 0
    while i < len(argv):
        arg = argv[i]
        if arg == "--to" and i + 1 < len(argv):
            explicit_to = argv[i + 1].lower().lstrip(".")
            i += 2
            continue
        if arg.startswith("--to="):
            explicit_to = arg.split("=", 1)[1].lower().lstrip(".")
            i += 1
            continue
        cleaned.append(arg)
        i += 1

    if explicit_to in ("eml", "mbox", "pst"):
        return explicit_to, cleaned

    # Check for mode-exclusive CLI flags and collect destination targets
    positionals: List[str] = []
    explicit_out: Optional[str] = None
    j = 0
    while j < len(cleaned):
        arg = cleaned[j]
        flag_name = arg.split("=", 1)[0]
        if flag_name in PST_EXCLUSIVE_FLAGS:
            return "pst", cleaned
        if flag_name in MBOX_EXCLUSIVE_FLAGS:
            return "mbox", cleaned
        if flag_name in EML_EXCLUSIVE_FLAGS:
            return "eml", cleaned

        if flag_name == "--out":
            if "=" in arg:
                explicit_out = arg.split("=", 1)[1]
            elif j + 1 < len(cleaned):
                explicit_out = cleaned[j + 1]
                j += 2
                continue
        elif flag_name in FLAGS_WITH_VALUE and "=" not in arg:
            j += 2
            continue
        elif not arg.startswith("-"):
            positionals.append(arg)
        j += 1

    target_candidate = explicit_out or (positionals[1] if len(positionals) >= 2 else (positionals[0] if len(positionals) == 1 else ""))
    if target_candidate:
        lower_target = target_candidate.lower()
        if lower_target.endswith(".pst"):
            return "pst", cleaned
        if mail2mbox.is_mbox_output_filename(lower_target):
            return "mbox", cleaned

    return "eml", cleaned


def main(argv: Optional[List[str]] = None) -> None:
    for _stream in (sys.stdout, sys.stderr):
        if hasattr(_stream, "reconfigure"):
            try:
                _stream.reconfigure(errors="replace")
            except Exception:
                pass

    args = list(sys.argv[1:] if argv is None else argv)
    mode, forwarded = _infer_mode(args)

    if mode == "help":
        _print_top_help()
        return

    if mode == "pst":
        import mail2pst
        mail2pst.main(forwarded, prog="mail2mail pst")
    elif mode == "mbox":
        mail2mbox.main(forwarded, prog="mail2mail mbox")
    else:
        mail2eml.main(forwarded, prog="mail2mail eml")


if __name__ == "__main__":
    main()
