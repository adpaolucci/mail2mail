# Changelog

All notable changes to `mail2mail` are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.2.0] - 2026-11-07

### Added
- **Fastmail & JMAP (RFC 8620 / RFC 8621) Native Client & Specification (`mail2jmap.py`, `mail2mail jmap`)**:
  - Implemented pure-Python, zero-third-party-dependency JMAP client supporting Fastmail and standard JMAP servers (Stalwart, Apache James, Cyrus).
  - Implemented automatic JMAP session discovery (`/jmap/session`), mailbox listing (`Mailbox/get`), email ID querying (`Email/query`), and batch email item retrieval (`Email/get`).
  - Added direct mapping of JMAP JSON payloads (including keywords `$seen`, `$flagged`, `$answered`, `$draft`, threads, and structured body values) to unified `MailRecord` and `EmailMessage` models.
  - Added CLI subcommand `mail2jmap` / `mail2mail jmap --token <token> --out <dest>` with folder filtering and incremental delta sync checkpointing.
  - Added automatic sniffing and ingestion of static `.jmap` and `.jmap.json` backup files in `mail_formats.py`.
  - Added comprehensive RFC compliance specification in `docs/rfc_compliance/RFC8620_RFC8621_JMAP.md` and automated test suite in `tests/test_jmap.py`.
- **Graceful Interruption, Pause/Resume & Signal Handling (`mail_shutdown.py`)**:
  - Coordinates graceful stops across all tools on `SIGINT` (Ctrl+C), `SIGTERM`, and `SIGBREAK` (Windows Console Break).
  - Traps interruption signals, finishes writing the active in-flight message cleanly, commits all batched state, unmounts all Outlook PST stores, and flushes audit logs.
  - Interactive TTY pause & resume: press `p` or `[Space]` to suspend processing, `[Space]` or `[Enter]` to resume, and `q` to stop safely.
  - Double-tap Ctrl+C detection for immediate emergency force termination (`os._exit(130)`).
  - Displays reconciliation summary and resumption guidance on interruption, exiting with standard code `130`.
- **Advanced Date & Folder Filtering (`mail_filter.py`)**:
  - Granular date filtering (`--since` / `--after`, `--until` / `--before`) supporting ISO 8601 (`YYYY-MM-DD`, `YYYY-MM-DDTHH:MM:SS`), RFC 2822 (`15 Jan 2024 10:00:00 +0000`), and single-year strings with timezone-aware UTC comparison.
  - Folder inclusion & exclusion rules (`--include-folder <glob>`, `--exclude-folder <glob>`) with case-insensitive glob and path-segment matching.
- **Human-Friendly Capacity & Size Splitting**:
  - `parse_size_bytes()` supporting human units (`20GB`, `500MB`, `15G`, `4g`, `1024K`, `1000B`) across `--split-size`, `--split-pst`, `--max-pst-size`, and `--max-mbox-size`.
- **Persistent Incremental / Delta Migration State Cache**:
  - `--state-file <path.sqlite | path.jsonl>` recording processed `Message-ID`s and SHA-256 hashes to enable instant incremental delta cutovers skipping previously migrated messages in $O(1)$ time.
- **Auditing, Pre-Flight Dry-Run & Forensics (`mail_audit.py`)**:
  - `--dry-run` inspection mode detailing volume, date spans, folder trees, and oversized attachment warnings (>20 MB / 25 MB).
  - Chain-of-custody structured manifests (`--audit-log` `.csv` / `.jsonl`) with message SHA-256 hashes.
  - Dedicated failure logging (`--error-log` `.jsonl`) and operational console mirroring (`--log-file`).
  - In-flight deduplication (`--dedup {auto, message-id, sha256}`).
  - Resumption from existing manifests (`--resume [manifest]`).
- **Zero-Data-Loss Safety & Adversarial Hardening**:
  - Emergency unwritten message spools (`_unwritten_messages`, `_unimported_messages`).
  - Emergency unattached file preservation (`_unattached_files`) when attachments exceed Outlook server limits.
  - Monotonically increasing collision counters on filename collisions.
  - Deep path sanitization and extended-length (`\\?\`) bypass on Windows.
  - Fuzz-tested against 60,000+ adversarial malformed inputs with 0 crashes.

---

## [2.1.0] - 2026-11-06

### Added
- **`mail2mbox` (`mail2mail mbox`)**: Universal exporter converting any supported mail format (including `.eml` folder trees, `.pst`, `.olm`, `.msg`, Maildir++, and Apple Mail `.emlx`) into standard RFC 4155 / `mboxrd` `.mbox` or compressed `.mbox.gz` / `.mbox.bz2` / `.mbox.xz` files, with `--layout {folder,year,month,flat}` and `--max-mbox-gb` size rotation.
- **Outlook `.pst` Input Support (`iter_pst_records`)**: Read `.pst` files via Outlook MAPI for `.pst` $\to$ `.eml`, `.pst` $\to$ `.mbox.gz`, or `.pst` $\to$ year/size-split `.pst` conversions.
- **RFC 3501 §5.1.3 `encode_imap_utf7`**: Added IMAP4rev1 Modified UTF-7 encoder alongside `decode_imap_utf7` for full bidirectional folder name round-trip compliance.
- **Automated Test Suite (`tests/`)**: Added `unittest` / `pytest`-compatible test suite (`test_rfc_compliance.py`, `test_mail_formats.py`, `test_exporters_and_cli.py`).
- **RFC & Format Specifications Reference (`docs/rfc_compliance/`)**: Detailed compliance reference covering RFC 5322, RFC 2045–2047, RFC 2183/2231/2392, RFC 3501, RFC 4155, [MS-CFB], [MS-OXMSG], [MS-OXPROPS], Maildir++, Apple Mail `.emlx`, and `.olm`.

### Changed
- Consolidated codebase into 5 clean modules (`mail2mail.py`, `mail2eml.py`, `mail2mbox.py`, `mail2pst.py`, `mail_formats.py`) and deduplicated shared RFC/header/flag helpers into `mail_formats.py`.

### Fixed
- **RFC 2046 §5.2.1 `message/rfc822` Attachments**: Added `_walk_outer_parts` in `mail2pst.py` so `multipart/*` containers are traversed recursively while encapsulated `message/rfc822` parts are treated as atomic leaf attachments.

---

## [2.0.0] - 2026-11-06

### Added
- Renamed project to **`mail2mail`** with unified `mail2mail`, `mail2eml`, and `mail2pst` CLI entry points and `pyproject.toml` packaging.
- **Multi-Format Support (`mail_formats.py`)**:
  - **Maildir & Maildir++** (`cur/`, `new/`, `tmp/`, `:2,FRSPT` and Windows `;2,FRSPT`, `.Folder.Subfolder`).
  - **Apple Mail** (`.emlx`, `.partial.emlx`, `*.mbox` / `*.imapmbox` bundles, XML `<plist>` 32-bit flags bitmask, and `Data/0..9/Messages` shard collapsing).
  - **Outlook `.msg`** (pure-Python [MS-CFB] OLE2 + [MS-OXMSG] parser for `.msg` $\to$ `.eml`/`.mbox` without Outlook, plus native MAPI `OpenSharedItem` transfer for `.msg` $\to$ `.pst`).
  - **Outlook for Mac `.olm`** (`com.microsoft.__Messages` XML hierarchy, HTML/plain bodies, and ZIP attachments).
  - **MH / Claws Mail / Sylpheed** (`.mh_sequences` sequence ranges and numbered message files).
  - **MMDF** (`\x01\x01\x01\x01` delimiter) and **Emacs Rmail Babyl** (`BABYL OPTIONS:`).
  - **Compressed Files & Archives** (`.gz`, `.bz2`, `.xz`, `.zip`, `.tar`, `.tar.gz`, `.tgz`, `.tar.bz2`, `.tar.xz`).

### Fixed
- **MAPI `PT_SYSTIME` UTC Timestamps**: Fixed `to_mapi_utc` to pass naive UTC `datetime` values to `pywin32`, preventing `VT_DATE` local-timezone shifts.
- **MAPI Property Ordering**: Set low-level MAPI properties (`PR_TRANSPORT_MESSAGE_HEADERS_W`, `PR_SENDER_*`, `PR_SENT_REPRESENTING_*`) *after* `HTMLBody`/`Body` so Outlook Object Model body sync does not reset them.
- **Thunderbird `X-Mozilla-Status` Bitmasks**: Corrected `MSG_FLAG_MARKED` (`0x0004`), `MSG_FLAG_EXPUNGED` (`0x0008`), and `MSG_FLAG_FORWARDED` (`0x1000`).
- **Safe PST Flush (`--flush-every`)**: Fixed `PstRouter` detach/reattach so `--flush-every` never deletes or overwrites the active `.pst` file.
- **Resilient Header & COM Handling**: Added `safe_get_header` / `safe_get_all_headers` for malformed RFC 5322 headers, `@com_retry` for `RPC_E_CALL_REJECTED`, and Windows `MAX_PATH` protection.
