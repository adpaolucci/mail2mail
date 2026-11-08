#!/usr/bin/env python3
"""
mail2mcp.py — Model Context Protocol (MCP) Server for mail2mail

Exposes mail2mail's universal email migration, forensic auditing, format sniffing,
and volume splitting capabilities to AI agents (Claude Desktop, Cursor, Antigravity,
AutoGPT) over standard JSON-RPC 2.0 stdio transport.
"""

from __future__ import annotations

import csv
import hashlib
import json
import os
import platform
import subprocess
import sys
from typing import Any, Dict, List, Optional

import mail_audit
import mail_formats

MCP_PROTOCOL_VERSION = "2024-11-05"
SERVER_NAME = "mail2mail-mcp"
SERVER_VERSION = "2.2.0"


def get_system_info() -> Dict[str, Any]:
    """Inspect local operating environment and toolchain availability."""
    is_windows = sys.platform == "win32"
    has_win32com = False
    outlook_available = False

    if is_windows:
        try:
            import win32com.client  # type: ignore

            has_win32com = True
        except ImportError:
            pass

        try:
            import winreg

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Outlook.Application"):
                outlook_available = True
        except (OSError, ImportError):
            pass

    return {
        "server": SERVER_NAME,
        "version": SERVER_VERSION,
        "python_version": platform.python_version(),
        "platform": sys.platform,
        "os_name": platform.system(),
        "windows_com_available": has_win32com,
        "microsoft_outlook_installed": outlook_available,
        "pst_export_supported_locally": outlook_available,
        "pst_note": (
            "Local PST export is available via Outlook COM."
            if outlook_available
            else (
                "PST creation in mail2mail requires Microsoft Outlook on Windows. "
                "For native cross-platform MS-PST binary generation on Linux/macOS without Outlook, "
                "use the companion pure-Rust tool 'mail2rust' (https://github.com/adpaolucci/mail2rust)."
            )
        ),
        "supported_formats": [
            "MBOX (mboxo, mboxrd)",
            "Thunderbird (ImapMail/Mail, .sbd, .msf)",
            "Maildir / Maildir++ (cur, new, tmp, :2,FRS flags)",
            "Apple Mail (.emlx, .mbox, .imapmbox)",
            "Outlook .msg (MS-CFB Compound File Binary)",
            "Outlook for Mac .olm (ZIP + XML reconstruction)",
            "Outlook .pst (via Outlook MAPI on Windows)",
            "MH / Sylpheed / Claws Mail (.mh_sequences)",
            "MMDF (\x01\x01\x01\x01) & Emacs Babyl",
            "JMAP RFC 8620 / RFC 8621 (Fastmail)",
            "RFC 5322 .eml Directory Trees",
            "Compressed Archives (.gz, .bz2, .xz, .zip, .tar)",
        ],
    }


def classify_mailbox_format(path: str) -> Dict[str, Any]:
    """Classify mailbox format from file header magic bytes or directory layout."""
    norm_path = os.path.normpath(path)
    if not os.path.exists(norm_path):
        raise FileNotFoundError(f"Path does not exist: {path}")

    is_dir = os.path.isdir(norm_path)
    size_bytes = 0

    if not is_dir:
        size_bytes = os.path.getsize(norm_path)
        with open(norm_path, "rb") as f:
            header = f.read(512)
        fmt = mail_formats._sniff_stream_format(header)
        ext = os.path.splitext(norm_path)[1].lower()

        format_names = {
            "pst": "Microsoft Outlook Personal Storage Table (Unicode PST)",
            "msg": "Microsoft Outlook Item (Compound File Binary Format)",
            "zip": "ZIP Compressed Archive (e.g. Google Takeout)",
            "gz": "Gzip Compressed Mailbox",
            "bz2": "Bzip2 Compressed Mailbox",
            "xz": "LZMA/XZ Compressed Mailbox",
            "tar": "TAR Unix Archive",
            "mmdf": "MMDF Mailbox Format",
            "babyl": "Emacs Babyl Mailbox Format",
            "jmap": "JMAP RFC 8620 / 8621 JSON Payload",
            "mbox": "Standard MBOX Mailbox (mboxo / mboxrd)",
            "emlx": "Apple Mail Message (.emlx)",
            "eml": "Standard RFC 5322 MIME Message (.eml)",
        }
        name = format_names.get(fmt, f"Generic ({ext or 'unknown'})")
        return {
            "path": norm_path,
            "format": fmt.upper(),
            "format_name": name,
            "size_bytes": size_bytes,
            "size_human": f"{size_bytes / (1024 * 1024):.2f} MB",
            "is_directory": False,
            "recommended_action": "mail2mail_convert or mail2mail_audit",
        }
    else:
        total_files = 0
        has_cur = False
        has_new = False
        has_sbd = False
        has_eml = False
        has_msg = False
        has_numbered = False

        for root, dirs, files in os.walk(norm_path):
            total_files += len(files)
            for d in dirs:
                dl = d.lower()
                if dl == "cur":
                    has_cur = True
                elif dl == "new":
                    has_new = True
                elif dl.endswith(".sbd"):
                    has_sbd = True
            for f in files:
                fl = f.lower()
                try:
                    size_bytes += os.path.getsize(os.path.join(root, f))
                except OSError:
                    pass
                if fl.endswith(".eml"):
                    has_eml = True
                elif fl.endswith(".msg"):
                    has_msg = True
                elif fl.isdigit():
                    has_numbered = True

        if has_cur and has_new:
            fmt = "MAILDIR"
            name = "Maildir / Maildir++ Directory Hierarchy"
        elif has_sbd:
            fmt = "THUNDERBIRD"
            name = "Mozilla Thunderbird Profile / ImapMail Directory"
        elif norm_path.lower().endswith(".mbox") or norm_path.lower().endswith(".imapmbox"):
            fmt = "APPLE_MBOX"
            name = "Apple Mail Directory Bundle"
        elif has_numbered:
            fmt = "MH"
            name = "MH / Sylpheed / Claws Mail Numbered Message Store"
        elif has_eml:
            fmt = "EML_TREE"
            name = "RFC 5322 EML Directory Tree"
        elif has_msg:
            fmt = "MSG_TREE"
            name = "Outlook .MSG File Directory Tree"
        else:
            fmt = "DIRECTORY"
            name = "Generic Directory Mailbox Store"

        return {
            "path": norm_path,
            "format": fmt,
            "format_name": name,
            "size_bytes": size_bytes,
            "size_human": f"{size_bytes / (1024 * 1024):.2f} MB",
            "file_count": total_files,
            "is_directory": True,
            "recommended_action": "mail2mail_convert",
        }


def run_convert(
    source: str,
    destination: str,
    target_format: Optional[str] = None,
    dedup: bool = False,
    since: Optional[str] = None,
    until: Optional[str] = None,
    bates_prefix: Optional[str] = None,
    audit_log: Optional[str] = None,
) -> Dict[str, Any]:
    """Execute mailbox conversion across formats."""
    norm_src = os.path.normpath(source)
    norm_dest = os.path.normpath(destination)
    if not os.path.exists(norm_src):
        raise FileNotFoundError(f"Source path not found: {source}")

    target = target_format.lower() if target_format else None
    if not target:
        dest_lower = norm_dest.lower()
        if dest_lower.endswith(".pst"):
            target = "pst"
        elif dest_lower.endswith((".mbox", ".mbox.gz", ".mbox.bz2", ".mbox.xz")):
            target = "mbox"
        else:
            target = "eml"

    if target == "pst":
        info = get_system_info()
        if not info["pst_export_supported_locally"]:
            return {
                "status": "error",
                "error_type": "PST_REQUIRES_OUTLOOK",
                "message": (
                    "PST conversion in mail2mail requires Microsoft Outlook on Windows via pywin32 COM. "
                    "On non-Windows platforms (Linux/macOS) or systems without Outlook, please export to 'eml' or 'mbox', "
                    "or use the companion pure-Rust tool 'mail2rust' (https://github.com/adpaolucci/mail2rust) "
                    "which generates 64-bit Unicode MS-PST binary containers natively with zero Outlook dependency."
                ),
                "source": norm_src,
                "destination": norm_dest,
            }

    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mail2mail.py")
    cmd = [sys.executable, script_path, target, norm_src, norm_dest]

    if dedup:
        cmd.append("--dedup")
    if since:
        cmd.extend(["--since", since])
    if until:
        cmd.extend(["--until", until])
    if bates_prefix:
        cmd.extend(["--bates-prefix", bates_prefix])
    if audit_log:
        cmd.extend(["--audit-log", audit_log])

    proc = subprocess.run(cmd, capture_output=True, text=True, errors="replace")
    if proc.returncode != 0:
        raise RuntimeError(f"Conversion failed (code {proc.returncode}):\n{proc.stderr or proc.stdout}")

    return {
        "status": "success",
        "source": norm_src,
        "destination": norm_dest,
        "target_format": target,
        "cli_output": proc.stdout.strip(),
    }


def run_audit(
    path: str,
    output_csv: Optional[str] = None,
    case_number: Optional[str] = "CASE-2026-M2M",
    examiner: Optional[str] = "Autonomous Agent",
    bates_prefix: Optional[str] = None,
) -> Dict[str, Any]:
    """Perform cryptographic forensic audit and Bates numbering."""
    norm_path = os.path.normpath(path)
    if not os.path.exists(norm_path):
        raise FileNotFoundError(f"Path does not exist: {path}")

    total_items = 0
    total_bytes = 0
    sha256_list: List[str] = []
    current_bates = 1000

    rows = []
    for rec in mail_formats.iter_mail_records(norm_path):
        total_items += 1
        raw = rec.raw_bytes
        size = len(raw)
        total_bytes += size
        digest = hashlib.sha256(raw).hexdigest()
        sha256_list.append(digest)

        bates_num = f"{bates_prefix}-{current_bates:07d}" if bates_prefix else "N/A"
        current_bates += 1

        folder_str = "/".join(rec.folder_parts) if rec.folder_parts else "Inbox"
        rows.append({
            "item_index": total_items,
            "bates_number": bates_num,
            "folder": folder_str,
            "source_name": rec.source_name,
            "size_bytes": size,
            "sha256": digest,
        })

    combined_hash = hashlib.sha256("".join(sha256_list).encode("ascii")).hexdigest() if sha256_list else ""
    bates_start = rows[0]["bates_number"] if rows and bates_prefix else "N/A"
    bates_end = rows[-1]["bates_number"] if rows and bates_prefix else "N/A"

    if output_csv:
        csv_path = os.path.normpath(output_csv)
        if os.path.dirname(csv_path):
            os.makedirs(os.path.dirname(os.path.abspath(csv_path)), exist_ok=True)
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["item_index", "bates_number", "folder", "source_name", "size_bytes", "sha256"],
            )
            writer.writeheader()
            writer.writerows(rows)

    return {
        "status": "success",
        "case_number": case_number,
        "examiner": examiner,
        "path": norm_path,
        "total_items": total_items,
        "total_size_bytes": total_bytes,
        "total_size_human": f"{total_bytes / (1024 * 1024):.2f} MB",
        "root_sha256": combined_hash,
        "bates_start": bates_start,
        "bates_end": bates_end,
        "output_csv": output_csv,
    }


def run_split(
    source: str,
    destination_dir: str,
    max_volume_mb: Optional[int] = 20000,
    split_by: Optional[str] = "size",
    target_format: Optional[str] = "eml",
) -> Dict[str, Any]:
    """Split oversized mailbox into volumes by size or year."""
    norm_src = os.path.normpath(source)
    norm_dest = os.path.normpath(destination_dir)
    os.makedirs(norm_dest, exist_ok=True)

    target = (target_format or "eml").lower()
    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mail2mail.py")
    cmd = [sys.executable, script_path, target, norm_src, norm_dest]

    if split_by == "year":
        cmd.extend(["--split-by", "year"])
    else:
        if target == "pst":
            cmd.extend(["--max-pst-gb", str(max(1, (max_volume_mb or 20000) // 1024))])
        elif target == "mbox":
            cmd.extend(["--max-mbox-size", f"{max_volume_mb}MB"])
        else:
            cmd.extend(["--max-dir-size", f"{max_volume_mb}MB"])

    proc = subprocess.run(cmd, capture_output=True, text=True, errors="replace")
    if proc.returncode != 0:
        raise RuntimeError(f"Split failed (code {proc.returncode}):\n{proc.stderr or proc.stdout}")

    return {
        "status": "success",
        "source": norm_src,
        "destination_dir": norm_dest,
        "split_by": split_by,
        "target_format": target,
        "cli_output": proc.stdout.strip(),
    }


def get_tools_manifest() -> List[Dict[str, Any]]:
    """Return JSON schema definitions for all exposed tools."""
    return [
        {
            "name": "mail2mail_info",
            "description": "Inspect system environment, Python version, platform, Outlook COM availability, and supported formats.",
            "inputSchema": {
                "type": "object",
                "properties": {},
            },
        },
        {
            "name": "mail2mail_sniff",
            "description": "Inspect magic bytes and classify mailbox format (PST, MSG, MBOX, EML, EMLX, OLM, Maildir, Thunderbird, etc.).",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to mailbox file or directory"}
                },
                "required": ["path"],
            },
        },
        {
            "name": "mail2mail_convert",
            "description": "Convert mailbox between formats (EML, MBOX, PST) with optional deduplication and date filtering.",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "source": {"type": "string", "description": "Source mailbox file or directory"},
                    "destination": {"type": "string", "description": "Target destination file or directory"},
                    "target_format": {"type": "string", "enum": ["eml", "mbox", "pst"], "description": "Target format"},
                    "dedup": {"type": "boolean", "description": "Deduplicate identical messages (default: false)"},
                    "since": {"type": "string", "description": "Optional ISO/YYYY-MM-DD date filter"},
                    "until": {"type": "string", "description": "Optional ISO/YYYY-MM-DD date filter"},
                    "bates_prefix": {"type": "string", "description": "Optional Bates numbering prefix"},
                    "audit_log": {"type": "string", "description": "Optional audit log path"},
                },
                "required": ["source", "destination"],
            },
        },
        {
            "name": "mail2mail_audit",
            "description": "Forensic audit: computes cryptographic SHA-256 item hashes, Bates numbering, and CSV manifest.",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Mailbox file or directory to audit"},
                    "output_csv": {"type": "string", "description": "Optional CSV manifest export path"},
                    "case_number": {"type": "string", "description": "Case reference (default: CASE-2026-M2M)"},
                    "examiner": {"type": "string", "description": "Examiner name or agent handle"},
                    "bates_prefix": {"type": "string", "description": "Bates stamp prefix"},
                },
                "required": ["path"],
            },
        },
        {
            "name": "mail2mail_split",
            "description": "Split large mailbox into volume chunks by size or year while preserving folder hierarchies.",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "source": {"type": "string", "description": "Source mailbox file or folder"},
                    "destination_dir": {"type": "string", "description": "Directory where split archives are created"},
                    "max_volume_mb": {"type": "integer", "description": "Maximum size per volume in MB"},
                    "split_by": {"type": "string", "enum": ["size", "year"], "description": "Split mode"},
                    "target_format": {"type": "string", "enum": ["eml", "mbox", "pst"], "description": "Target format"},
                },
                "required": ["source", "destination_dir"],
            },
        },
    ]


def handle_jsonrpc(req: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Process a single JSON-RPC 2.0 request."""
    req_id = req.get("id")
    method = req.get("method")
    if not method:
        if req_id is not None:
            return {"jsonrpc": "2.0", "id": req_id, "error": {"code": -32600, "message": "Missing method"}}
        return None

    if method == "initialize":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "protocolVersion": MCP_PROTOCOL_VERSION,
                "capabilities": {"tools": {}},
                "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
            },
        }

    if method == "notifications/initialized":
        return None

    if method == "ping":
        return {"jsonrpc": "2.0", "id": req_id, "result": {}}

    if method == "tools/list":
        return {"jsonrpc": "2.0", "id": req_id, "result": {"tools": get_tools_manifest()}}

    if method == "tools/call":
        params = req.get("params") or {}
        name = params.get("name")
        args = params.get("arguments") or {}

        try:
            if name == "mail2mail_info":
                res = get_system_info()
            elif name == "mail2mail_sniff":
                res = classify_mailbox_format(args["path"])
            elif name == "mail2mail_convert":
                res = run_convert(
                    source=args["source"],
                    destination=args["destination"],
                    target_format=args.get("target_format"),
                    dedup=args.get("dedup", False),
                    since=args.get("since"),
                    until=args.get("until"),
                    bates_prefix=args.get("bates_prefix"),
                    audit_log=args.get("audit_log"),
                )
            elif name == "mail2mail_audit":
                res = run_audit(
                    path=args["path"],
                    output_csv=args.get("output_csv"),
                    case_number=args.get("case_number", "CASE-2026-M2M"),
                    examiner=args.get("examiner", "Autonomous Agent"),
                    bates_prefix=args.get("bates_prefix"),
                )
            elif name == "mail2mail_split":
                res = run_split(
                    source=args["source"],
                    destination_dir=args["destination_dir"],
                    max_volume_mb=args.get("max_volume_mb", 20000),
                    split_by=args.get("split_by", "size"),
                    target_format=args.get("target_format", "eml"),
                )
            else:
                return {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "error": {"code": -32601, "message": f"Method not found: {name}"},
                }

            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": json.dumps(res, indent=2)}],
                    "isError": False,
                },
            }
        except Exception as e:
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": f"Error: {e}"}],
                    "isError": True,
                },
            }

    return {
        "jsonrpc": "2.0",
        "id": req_id,
        "error": {"code": -32601, "message": f"Method not found: {method}"},
    }


def run_stdio_server() -> None:
    """Run standard I/O newline-delimited JSON-RPC loop."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8")

    for line in sys.stdin:
        trimmed = line.strip()
        if not trimmed:
            continue
        try:
            req = json.loads(trimmed)
            resp = handle_jsonrpc(req)
            if resp is not None:
                sys.stdout.write(json.dumps(resp) + "\n")
                sys.stdout.flush()
        except json.JSONDecodeError as err:
            err_resp = {
                "jsonrpc": "2.0",
                "id": None,
                "error": {"code": -32700, "message": f"Parse error: {err}"},
            }
            sys.stdout.write(json.dumps(err_resp) + "\n")
            sys.stdout.flush()


def main(argv: Optional[List[str]] = None) -> None:
    args = argv if argv is not None else sys.argv[1:]
    if "--tools" in args:
        print(json.dumps(get_tools_manifest(), indent=2))
        return

    if "--help" in args or "-h" in args:
        print(
            "mail2mail MCP Server (Model Context Protocol)\n\n"
            "Usage:\n"
            "  python mail2mcp.py          Run stdio JSON-RPC server\n"
            "  python mail2mcp.py --tools  Print tool schemas in JSON format\n"
        )
        return

    run_stdio_server()


if __name__ == "__main__":
    main()
