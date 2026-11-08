#!/usr/bin/env python3
"""
mail2jmap.py — Pure Python RFC 8620 / RFC 8621 JMAP Client & Mail Ingestion Engine

Connects to Fastmail or any standards-compliant JMAP (JSON Meta Application Protocol)
server (RFC 8620 Core, RFC 8621 Mail) to synchronize mailboxes directly into:
  - Microsoft Outlook .pst archives (via mail2pst)
  - Standard .mbox / .mbox.gz files (via mail2mbox)
  - Clean .eml directory trees (via mail2eml)

Zero third-party dependencies — built entirely on Python's standard library
(urllib.request, json, ssl, email).

Examples:
  # Ingest entire Fastmail account into an Outlook PST
  python mail2jmap.py --token "fmu1-xxxx-xxxx" --out "fastmail.pst"

  # Sync specific folders with incremental delta checkpointing
  python mail2jmap.py --token "fmu1-xxxx-xxxx" --folders "Inbox,Archive" --checkpoint "sync.json" --out "fastmail.mbox"

  # Connect to custom JMAP server (e.g. Stalwart, Apache James, Cyrus)
  python mail2jmap.py --session-url "https://mail.example.com/.well-known/jmap" --token "secret" --out "out_eml/"
"""

from __future__ import annotations

import argparse
import email.message
import email.utils
import json
import os
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from typing import Any, Callable, Dict, Iterator, List, Optional, Set, Tuple

import mail_formats
from mail_formats import MailRecord, default_flags


# Default Fastmail session URL
FASTMAIL_SESSION_URL = "https://api.fastmail.com/jmap/session"


class JmapError(Exception):
    """Base exception for JMAP protocol and network failures."""
    pass


class JmapSession:
    """Represents an RFC 8620 JMAP Session Resource."""

    def __init__(
        self,
        session_url: str,
        token: str,
        insecure: bool = False,
    ) -> None:
        self.session_url = session_url
        self.token = token
        self.insecure = insecure
        self.api_url: str = ""
        self.download_url: str = ""
        self.upload_url: str = ""
        self.primary_account_id: str = ""
        self.accounts: Dict[str, Any] = {}
        self.capabilities: Dict[str, Any] = {}
        self.raw_session: Dict[str, Any] = {}

    def _create_ssl_context(self) -> ssl.SSLContext:
        ctx = ssl.create_default_context()
        if self.insecure:
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
        return ctx

    def request(
        self,
        method: str,
        url: str,
        json_data: Optional[Dict[str, Any]] = None,
    ) -> Tuple[int, Dict[str, Any]]:
        """Perform an HTTP request with JMAP Bearer Token authentication."""
        headers = {
            "Authorization": f"Bearer {self.token}",
            "User-Agent": "mail2mail/3.3.0 (Python RFC8620/8621 JMAP Client)",
            "Accept": "application/json",
        }
        body_bytes = None
        if json_data is not None:
            headers["Content-Type"] = "application/json"
            body_bytes = json.dumps(json_data).encode("utf-8")

        req = urllib.request.Request(url, data=body_bytes, headers=headers, method=method)
        ctx = self._create_ssl_context()

        try:
            with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
                status = resp.status
                raw = resp.read().decode("utf-8")
                return status, json.loads(raw) if raw.strip() else {}
        except urllib.error.HTTPError as e:
            raw_err = e.read().decode("utf-8", errors="replace")
            raise JmapError(f"HTTP {e.code} ({e.reason}): {raw_err}") from e
        except Exception as e:
            raise JmapError(f"Network error communicating with JMAP server: {e}") from e

    def connect(self) -> None:
        """Fetch and parse the RFC 8620 Session object."""
        status, data = self.request("GET", self.session_url)
        if status != 200:
            raise JmapError(f"Session discovery failed with HTTP {status}: {data}")

        self.raw_session = data
        self.api_url = data.get("apiUrl", "")
        self.download_url = data.get("downloadUrl", "")
        self.upload_url = data.get("uploadUrl", "")
        self.capabilities = data.get("capabilities", {})
        self.accounts = data.get("accounts", {})

        primary_accounts = data.get("primaryAccounts", {})
        self.primary_account_id = primary_accounts.get("urn:ietf:params:jmap:mail") or ""
        if not self.primary_account_id and self.accounts:
            self.primary_account_id = next(iter(self.accounts.keys()))

        if not self.primary_account_id:
            raise JmapError("No mail-capable JMAP account found in session response.")
        if not self.api_url:
            raise JmapError("JMAP session did not specify an apiUrl.")


class JmapClient:
    """Client for RFC 8620 and RFC 8621 operations."""

    def __init__(
        self,
        session_url: str = FASTMAIL_SESSION_URL,
        token: str = "",
        insecure: bool = False,
    ) -> None:
        self.session = JmapSession(session_url, token, insecure=insecure)

    @classmethod
    def fastmail(cls, token: str) -> JmapClient:
        return cls(FASTMAIL_SESSION_URL, token)

    def connect(self) -> None:
        self.session.connect()

    @property
    def account_id(self) -> str:
        return self.session.primary_account_id

    def call(self, method_calls: List[List[Any]]) -> List[List[Any]]:
        """Invoke RFC 8620 batch method calls against apiUrl."""
        payload = {
            "using": [
                "urn:ietf:params:jmap:core",
                "urn:ietf:params:jmap:mail",
            ],
            "methodCalls": method_calls,
        }
        _, resp = self.session.request("POST", self.session.api_url, json_data=payload)
        return resp.get("methodResponses", [])

    def list_mailboxes(self, account_id: Optional[str] = None) -> List[Dict[str, Any]]:
        """Retrieve all mailboxes via Mailbox/get."""
        acc_id = account_id or self.account_id
        responses = self.call([
            ["Mailbox/get", {"accountId": acc_id, "ids": None}, "m0"]
        ])
        for name, args, _ in responses:
            if name == "Mailbox/get":
                return args.get("list", [])
        return []

    def query_emails(
        self,
        mailbox_id: str,
        limit: int = 200,
        position: int = 0,
        account_id: Optional[str] = None,
    ) -> List[str]:
        """Query email IDs within a mailbox via Email/query."""
        acc_id = account_id or self.account_id
        responses = self.call([
            [
                "Email/query",
                {
                    "accountId": acc_id,
                    "filter": {"inMailbox": mailbox_id},
                    "sort": [{"property": "receivedAt", "isAscending": True}],
                    "limit": limit,
                    "position": position,
                },
                "q0",
            ]
        ])
        for name, args, _ in responses:
            if name == "Email/query":
                return args.get("ids", [])
        return []

    def get_emails(
        self,
        email_ids: List[str],
        account_id: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """Fetch email entities via Email/get."""
        if not email_ids:
            return []
        acc_id = account_id or self.account_id
        responses = self.call([
            [
                "Email/get",
                {
                    "accountId": acc_id,
                    "ids": email_ids,
                    "properties": [
                        "id", "blobId", "threadId", "mailboxIds", "keywords", "size",
                        "receivedAt", "sentAt", "messageId", "inReplyTo", "references",
                        "subject", "from", "to", "cc", "bcc", "replyTo",
                        "textBody", "htmlBody", "bodyValues", "attachments",
                    ],
                    "bodyProperties": ["partId", "blobId", "size", "type"],
                    "fetchTextBodyValues": True,
                    "fetchHTMLBodyValues": True,
                },
                "e0",
            ]
        ])
        for name, args, _ in responses:
            if name == "Email/get":
                return args.get("list", [])
        return []

    def iter_all_items(
        self,
        folder_filter: Optional[List[str]] = None,
        checkpoint: Optional[Dict[str, Any]] = None,
        max_per_folder: Optional[int] = None,
    ) -> Iterator[Tuple[Dict[str, Any], Dict[str, str]]]:
        """Stream raw JMAP email dictionaries along with mailbox mapping."""
        mailboxes = self.list_mailboxes()
        mailbox_map: Dict[str, str] = {}
        for mb in mailboxes:
            mailbox_map[mb.get("id", "")] = mb.get("name", "Inbox")

        migrated_ids: Set[str] = set()
        if checkpoint and "migrated_message_ids" in checkpoint:
            migrated_ids = set(checkpoint["migrated_message_ids"])

        for mb in mailboxes:
            mb_name = mb.get("name", "Inbox")
            mb_id = mb.get("id", "")

            if folder_filter:
                if not any(f.lower() == mb_name.lower() for f in folder_filter):
                    continue

            position = 0
            page_size = 200
            folder_count = 0

            while True:
                remaining = (max_per_folder - folder_count) if max_per_folder else page_size
                if remaining <= 0:
                    break
                curr_limit = min(page_size, remaining)

                ids = self.query_emails(mb_id, limit=curr_limit, position=position)
                if not ids:
                    break

                fetch_ids = [i for i in ids if i not in migrated_ids]
                if fetch_ids:
                    items = self.get_emails(fetch_ids)
                    for item in items:
                        migrated_ids.add(item.get("id", ""))
                        if checkpoint is not None:
                            checkpoint.setdefault("migrated_message_ids", []).append(item.get("id", ""))
                            checkpoint["total_messages_migrated"] = checkpoint.get("total_messages_migrated", 0) + 1
                        folder_count += 1
                        yield item, mailbox_map

                position += len(ids)
                if len(ids) < curr_limit:
                    break

    def iter_all_records(
        self,
        folder_filter: Optional[List[str]] = None,
        checkpoint: Optional[Dict[str, Any]] = None,
        max_per_folder: Optional[int] = None,
    ) -> Iterator[MailRecord]:
        """Stream MailRecord instances across all mailboxes."""
        for item, mb_map in self.iter_all_items(folder_filter, checkpoint, max_per_folder):
            yield jmap_item_to_mail_record(item, mb_map)


def jmap_item_to_email_message(
    item: Dict[str, Any],
    mailbox_map: Optional[Dict[str, str]] = None,
) -> email.message.EmailMessage:
    """Convert an RFC 8621 JMAP Email dictionary into a Python EmailMessage."""
    msg = email.message.EmailMessage()

    # Subject
    msg["Subject"] = item.get("subject") or "(No Subject)"

    # Addresses
    if item.get("from"):
        f = item["from"][0]
        msg["From"] = email.utils.formataddr((f.get("name") or "", f.get("email") or ""))

    if item.get("to"):
        msg["To"] = ", ".join(
            email.utils.formataddr((r.get("name") or "", r.get("email") or "")) for r in item["to"]
        )
    if item.get("cc"):
        msg["Cc"] = ", ".join(
            email.utils.formataddr((r.get("name") or "", r.get("email") or "")) for r in item["cc"]
        )
    if item.get("bcc"):
        msg["Bcc"] = ", ".join(
            email.utils.formataddr((r.get("name") or "", r.get("email") or "")) for r in item["bcc"]
        )
    if item.get("replyTo"):
        msg["Reply-To"] = ", ".join(
            email.utils.formataddr((r.get("name") or "", r.get("email") or "")) for r in item["replyTo"]
        )

    # Identifiers
    msg_ids = item.get("messageId") or []
    if msg_ids:
        msg["Message-ID"] = msg_ids[0]
    else:
        msg["Message-ID"] = f"<{item.get('id', 'unknown')}@jmap.local>"

    in_reply = item.get("inReplyTo") or []
    if in_reply:
        msg["In-Reply-To"] = in_reply[0]

    refs = item.get("references") or []
    if refs:
        msg["References"] = " ".join(refs)

    # Dates
    date_str = item.get("sentAt") or item.get("receivedAt")
    if date_str:
        try:
            dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
            msg["Date"] = email.utils.format_datetime(dt)
        except Exception:
            msg["Date"] = email.utils.formatdate(timeval=None, localtime=False, usegmt=True)
    else:
        msg["Date"] = email.utils.formatdate(timeval=None, localtime=False, usegmt=True)

    # Body values
    body_values = item.get("bodyValues") or {}
    text_parts = item.get("textBody") or []
    html_parts = item.get("htmlBody") or []

    text_content = ""
    for part in text_parts:
        pid = part.get("partId")
        if pid and pid in body_values:
            text_content += body_values[pid].get("value", "")

    html_content = ""
    for part in html_parts:
        pid = part.get("partId")
        if pid and pid in body_values:
            html_content += body_values[pid].get("value", "")

    if html_content and text_content:
        msg.set_content(text_content)
        msg.add_alternative(html_content, subtype="html")
    elif html_content:
        msg.set_content(html_content, subtype="html")
    else:
        msg.set_content(text_content or "")

    return msg


def jmap_item_to_mail_record(
    item: Dict[str, Any],
    mailbox_map: Optional[Dict[str, str]] = None,
) -> MailRecord:
    """Transform an RFC 8621 JMAP Email dictionary into a unified MailRecord."""
    mb_map = mailbox_map or {}
    folder_parts: List[str] = []

    mailbox_ids = item.get("mailboxIds") or {}
    for mid in mailbox_ids:
        if mid in mb_map:
            folder_parts = [mb_map[mid]]
            break
        mid_low = mid.lower()
        if "inbox" in mid_low:
            folder_parts = ["Inbox"]
            break
        elif "sent" in mid_low:
            folder_parts = ["Sent"]
            break
        elif "archive" in mid_low:
            folder_parts = ["Archive"]
            break
        elif "trash" in mid_low:
            folder_parts = ["Trash"]
            break

    if not folder_parts:
        folder_parts = ["Inbox"]

    msg = jmap_item_to_email_message(item, mb_map)
    raw_bytes = msg.as_bytes()

    keywords = item.get("keywords") or {}
    is_read = "$seen" in keywords
    is_flagged = "$flagged" in keywords
    is_replied = "$answered" in keywords
    is_draft = "$draft" in keywords
    is_deleted = "$deleted" in keywords

    # Timestamp
    mtime = None
    rec_at = item.get("receivedAt") or item.get("sentAt")
    if rec_at:
        try:
            mtime = datetime.fromisoformat(rec_at.replace("Z", "+00:00")).timestamp()
        except Exception:
            pass

    return MailRecord(
        raw_bytes=raw_bytes,
        folder_parts=folder_parts,
        source_name=f"{item.get('id', 'msg')}.eml",
        is_deleted=is_deleted,
        flags={
            "is_read": is_read,
            "is_replied": is_replied,
            "is_flagged": is_flagged,
            "is_forwarded": False,
            "is_deleted": is_deleted,
            "is_draft": is_draft,
        },
        file_mtime=mtime,
    )


def iter_jmap_records(
    file_or_data: Any,
    folder_parts: Optional[List[str]] = None,
    include_deleted: bool = False,
) -> Iterator[MailRecord]:
    """
    Parse a JMAP JSON file, stream, or raw text and yield MailRecords.
    Supports single JMAP Email entities, lists of Email entities,
    or RFC 8620 batch method responses (Email/get).
    """
    if isinstance(file_or_data, (str, bytes)):
        if isinstance(file_or_data, str) and os.path.exists(file_or_data):
            with open(file_or_data, "r", encoding="utf-8", errors="replace") as fp:
                data = json.load(fp)
        elif isinstance(file_or_data, bytes):
            data = json.loads(file_or_data.decode("utf-8", errors="replace"))
        else:
            data = json.loads(file_or_data)
    elif hasattr(file_or_data, "read"):
        content = file_or_data.read()
        if isinstance(content, bytes):
            content = content.decode("utf-8", errors="replace")
        data = json.loads(content)
    else:
        data = file_or_data

    # Extract email dictionaries
    email_items: List[Dict[str, Any]] = []

    if isinstance(data, dict):
        if "methodResponses" in data:
            for resp in data["methodResponses"]:
                if isinstance(resp, list) and len(resp) >= 2 and resp[0] == "Email/get":
                    email_items.extend(resp[1].get("list", []))
        elif "list" in data and isinstance(data["list"], list):
            email_items.extend(data["list"])
        elif "id" in data and ("mailboxIds" in data or "subject" in data or "bodyValues" in data):
            email_items.append(data)
    elif isinstance(data, list):
        for entry in data:
            if isinstance(entry, dict) and "id" in entry:
                email_items.append(entry)

    for item in email_items:
        rec = jmap_item_to_mail_record(item)
        if folder_parts:
            rec.folder_parts = list(folder_parts)
        if rec.is_deleted and not include_deleted:
            continue
        yield rec


def export_jmap(
    session_url: str,
    token: str,
    output: str,
    target_format: Optional[str] = None,
    folders: Optional[List[str]] = None,
    checkpoint_file: Optional[str] = None,
    max_per_folder: Optional[int] = None,
    insecure: bool = False,
) -> int:
    """Connect to a JMAP server, stream all emails, and export to PST, MBOX, or EML."""
    client = JmapClient(session_url=session_url, token=token, insecure=insecure)
    print(f"[*] Connecting to JMAP session: {session_url}...")
    client.connect()
    print(f"[*] Authenticated as primary mail account: {client.account_id}")

    chk: Dict[str, Any] = {}
    if checkpoint_file and os.path.isfile(checkpoint_file):
        try:
            with open(checkpoint_file, "r", encoding="utf-8") as fp:
                chk = json.load(fp)
            print(f"[*] Resuming from checkpoint: {checkpoint_file} ({chk.get('total_messages_migrated', 0)} previously migrated)")
        except Exception as e:
            print(f"[!] Warning: could not load checkpoint file ({e}), starting fresh.")
            chk = {}

    out_low = output.lower()
    inferred_fmt = target_format or ("pst" if out_low.endswith(".pst") else ("mbox" if out_low.endswith((".mbox", ".mbox.gz")) else "eml"))

    raw_items: List[Dict[str, Any]] = []
    for item, mb_map in client.iter_all_items(folder_filter=folders, checkpoint=chk, max_per_folder=max_per_folder):
        raw_items.append(item)

    total_count = len(raw_items)
    print(f"[*] Retrieved {total_count} emails via JMAP.")

    if total_count == 0:
        if checkpoint_file:
            chk["timestamp"] = datetime.now(timezone.utc).isoformat()
            try:
                with open(checkpoint_file, "w", encoding="utf-8") as fp:
                    json.dump(chk, fp, indent=2)
            except Exception:
                pass
        return 0

    import tempfile
    with tempfile.NamedTemporaryFile("w", suffix=".jmap", encoding="utf-8", delete=False) as tf:
        temp_path = tf.name
        jmap_export = {
            "using": ["urn:ietf:params:jmap:core", "urn:ietf:params:jmap:mail"],
            "methodResponses": [
                [
                    "Email/get",
                    {
                        "accountId": client.account_id,
                        "list": raw_items,
                    },
                    "e0",
                ]
            ],
        }
        json.dump(jmap_export, tf, indent=2)

    try:
        if inferred_fmt == "pst":
            import mail2pst
            print(f"[*] Exporting to Outlook PST: {output}...")
            mail2pst.main([temp_path, output], prog="mail2jmap")
        elif inferred_fmt == "mbox":
            import mail2mbox
            print(f"[*] Exporting to MBOX: {output}...")
            mail2mbox.main([temp_path, output], prog="mail2jmap")
        else:
            import mail2eml
            print(f"[*] Exporting to EML directory: {output}...")
            mail2eml.main([temp_path, output], prog="mail2jmap")
    finally:
        try:
            os.remove(temp_path)
        except OSError:
            pass

    if checkpoint_file:
        chk["timestamp"] = datetime.now(timezone.utc).isoformat()
        chk["status"] = "completed"
        try:
            with open(checkpoint_file, "w", encoding="utf-8") as fp:
                json.dump(chk, fp, indent=2)
            print(f"[*] Saved final checkpoint state to: {checkpoint_file}")
        except Exception as e:
            print(f"[!] Warning: Failed to save checkpoint file: {e}")

    return total_count


def main(argv: Optional[List[str]] = None) -> None:
    parser = argparse.ArgumentParser(
        prog="mail2jmap",
        description="Synchronize email directly from Fastmail or RFC 8620/8621 JMAP servers into PST, MBOX, or EML.",
    )
    parser.add_argument(
        "--session-url",
        default=FASTMAIL_SESSION_URL,
        help=f"JMAP session resource URL (default: {FASTMAIL_SESSION_URL})",
    )
    parser.add_argument(
        "--token",
        required=True,
        help="JMAP API Bearer token (e.g. Fastmail API token)",
    )
    parser.add_argument(
        "--out", "-o",
        required=True,
        help="Output destination path (.pst, .mbox, or directory for .eml)",
    )
    parser.add_argument(
        "--folders",
        help="Comma-separated list of mailbox folder names to sync (e.g. 'Inbox,Sent,Archive')",
    )
    parser.add_argument(
        "--checkpoint",
        help="Path to JSON checkpoint file for resume and incremental delta synchronization",
    )
    parser.add_argument(
        "--max-per-folder",
        type=int,
        help="Maximum messages to retrieve per mailbox",
    )
    parser.add_argument(
        "--insecure",
        action="store_true",
        help="Disable SSL/TLS certificate verification (for testing/local development)",
    )

    args = parser.parse_args(argv)
    folder_list = [f.strip() for f in args.folders.split(",") if f.strip()] if args.folders else None

    try:
        count = export_jmap(
            session_url=args.session_url,
            token=args.token,
            output=args.out,
            folders=folder_list,
            checkpoint_file=args.checkpoint,
            max_per_folder=args.max_per_folder,
            insecure=args.insecure,
        )
        print(f"[+] Successfully migrated {count} messages via JMAP to: {args.out}")
    except KeyboardInterrupt:
        print("\n[!] Migration paused by user.")
        sys.exit(130)
    except Exception as e:
        print(f"\n[ERROR] JMAP migration failed: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
