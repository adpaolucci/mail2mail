#!/usr/bin/env python3
"""
mail2mail.py — Unified Universal Email & Mailbox Converter

Converts between any supported email/mailbox format (MBOX, Thunderbird ImapMail/Mail,
Maildir/Maildir++, Apple Mail .emlx/.mbox, Outlook .msg, Outlook for Mac .olm,
MH, MMDF, Babyl, EML, and .gz/.bz2/.xz/.zip/.tar* archives) and either:
  - .eml directory hierarchies (`mail2mail eml ...` or `mail2eml`)
  - Microsoft Outlook .pst archives (`mail2mail pst ...` or `mail2pst`)

Examples:
  # Subcommand syntax
  python mail2mail.py eml "D:\\Mail\\inbox.mbox.gz" "D:\\Export_EML" --layout month
  python mail2mail.py pst "D:\\Mail\\ImapMail" "D:\\PSTs\\Mailbox.pst"
  python mail2mail.py pst --src "D:\\Export_EML" --out-dir "D:\\PSTs" --split-by year --max-pst-gb 15

  # Auto-detected target syntax
  python mail2mail.py "D:\\Mail\\export.olm" "D:\\PSTs\\MacMail.pst"
  python mail2mail.py "D:\\Mail\\export.olm" "D:\\Export_EML" --to eml
"""

from __future__ import annotations

import sys
from typing import List, Optional, Tuple

import mail2eml


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


def _print_top_help() -> None:
    print(
        "usage: mail2mail {eml,pst} [options] ...\n"
        "   or: mail2mail <src> <dest.pst | out_eml_dir> [--to {eml,pst}] [options]\n\n"
        "mail2mail — Universal Email & Mailbox Converter (MBOX, Thunderbird ImapMail/Mail,\n"
        "Maildir/Maildir++, Apple Mail .emlx/.mbox, Outlook .msg/.olm, MH, MMDF, Babyl,\n"
        "EML, and .gz/.bz2/.xz/.zip/.tar* archives -> EML or Outlook PST).\n\n"
        "Commands:\n"
        "  eml (or mail2eml)   Convert any supported mail source into .eml files (No Outlook required)\n"
        "  pst (or mail2pst)   Convert/import any supported mail source into Outlook .pst file(s)\n\n"
        "Run `python mail2mail.py eml --help` or `python mail2mail.py pst --help` for full options."
    )


def _infer_mode(argv: List[str]) -> Tuple[str, List[str]]:
    """
    Determine whether argv targets 'eml' or 'pst' and strip any explicit
    subcommand ('eml'/'pst') or '--to {eml,pst}' flag.
    """
    if not argv or argv[0] in ("-h", "--help"):
        return "help", argv

    first = argv[0].lower()
    if first in ("eml", "to-eml", "mail2eml"):
        return "eml", argv[1:]
    if first in ("pst", "to-pst", "mail2pst"):
        return "pst", argv[1:]

    # Check for explicit --to {eml,pst}
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

    if explicit_to in ("eml", "pst"):
        return explicit_to, cleaned

    # Check for PST-specific flags or .pst file extension in arguments
    for arg in cleaned:
        flag_name = arg.split("=", 1)[0]
        if flag_name in PST_EXCLUSIVE_FLAGS:
            return "pst", cleaned
        if arg.lower().endswith(".pst"):
            return "pst", cleaned

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
    else:
        mail2eml.main(forwarded, prog="mail2mail eml")


if __name__ == "__main__":
    main()
