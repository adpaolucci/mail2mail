# `mail2mail` — Universal Email & Mailbox Converter (`mail2eml` / `mail2mbox` / `mail2pst`)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-3776AB.svg)](pyproject.toml)
[![RFC Compliant](https://img.shields.io/badge/RFC-5322%20%7C%202045--2047%20%7C%203501%20%7C%204155-success.svg)](docs/rfc_compliance/README.md)

Convert between **MBOX**, **Thunderbird `ImapMail` / `Mail`** (mbox and Maildir modes), **Maildir / Maildir++**, **Apple Mail (`.emlx` / `.mbox` bundles)**, **Outlook `.msg`**, **Outlook for Mac `.olm`**, **Outlook `.pst`**, **MH / Sylpheed / Claws Mail**, **MMDF**, **Emacs Babyl**, **`.eml` trees**, and **compressed archives (`.gz`, `.bz2`, `.xz`, `.zip`, `.tar.gz`, `.tgz`)** into `.eml` folder hierarchies (`mail2eml`), standard `mboxrd` `.mbox` / `.mbox.gz` files (`mail2mbox`), or Microsoft Outlook **.pst** archives (`mail2pst`) with full folder hierarchy preservation, message flag mapping, size/year splitting, and resilient Outlook MAPI integration.

---

## Commands & Architecture

| Command / Module | Purpose | Requires Outlook? |
| :--- | :--- | :--- |
| [`mail2mail.py`](mail2mail.py) (`mail2mail`) | Unified master CLI (`mail2mail eml ...`, `mail2mail mbox ...`, `mail2mail pst ...`, or auto-detected from destination extension) | Only for `.pst` input/output |
| [`mail2eml.py`](mail2eml.py) (`mail2eml`) | Convert any supported mail format or archive into `.eml` files (`folder`, `year`, `month`, or `flat` layout) | No (unless reading `.pst`) |
| [`mail2mbox.py`](mail2mbox.py) (`mail2mbox`) | Convert any supported mail format or archive into standard `mboxrd` `.mbox` (or `.mbox.gz` / `.mbox.bz2` / `.mbox.xz`) file(s) | No (unless reading `.pst`) |
| [`mail2pst.py`](mail2pst.py) (`mail2pst`) | Convert/import any supported mail format or archive into one or more Outlook `.pst` files (single PST, year-split, `N`-way split, or auto-rotated at `--max-pst-gb`) | Yes (`pywin32`) |
| [`mail_formats.py`](mail_formats.py) | Shared multi-format detection, streaming, flag normalization, and folder hierarchy engine | No (Stdlib only; `pywin32` for `.pst` input) |

---

## Supported Email & Mailbox Formats

All commands share [`mail_formats.py`](mail_formats.py) and automatically detect any of the following formats:

| Format | Supported Inputs | Metadata & Status Flags Preserved |
| :--- | :--- | :--- |
| **1. Standard MBOX** (`mboxo` / `mboxrd`) | `.mbox` files, extensionless mbox files | `X-Mozilla-Status`, `Status`, `X-Status`, `>From ` un-escaping |
| **2. Thunderbird `ImapMail` / `Mail`** | Account folders, `ImapMail`/`Mail` roots, `.sbd` hierarchies (both **mbox** and **Maildir** modes) | `X-Mozilla-Status` (`Read`, `Replied`, `Flagged`, `Forwarded`, `Expunged`), IMAP Modified UTF-7 folder decoding |
| **3. Maildir & Maildir++** | `cur/`, `new/`, `tmp/`, `.Folder.Subfolder` Maildir++ trees | `:2,FRSPT` / `;2,FRSPT` filename flags (`S`=Seen, `R`=Replied, `F`=Flagged, `P`=Forwarded, `T`=Trashed), `new/` unread state |
| **4. Apple Mail (`.emlx` & `.mbox` Bundles)** | `.emlx`, `.partial.emlx`, macOS `*.mbox` / `*.imapmbox` / `*.popmbox` directory bundles | Strips byte-count prefix & parses XML `<plist>` `flags` bitmask (`Read`, `Deleted`, `Replied`, `Flagged`, `Forwarded`); collapses internal `Data/0..9/Messages` shards |
| **5. Outlook `.msg` (MS-CFB / MS-OXMSG)** | `.msg` files and folder trees | Native MAPI `OpenSharedItem` transfer when importing to PST, plus pure-Python MS-CFB OLE2 + MS-OXMSG parser for `.msg` → `.eml`/`.mbox` conversion without Outlook |
| **6. Outlook for Mac `.olm`** | `.olm` archive files | Extracts `com.microsoft.__Messages` folder hierarchy, HTML/plain bodies, inline `cid:` & file attachments, and Read/Flagged state |
| **7. Outlook `.pst` (Personal Folders)** | `.pst` files and directories containing `.pst` files | Mounts `.pst` via Outlook MAPI, walks all folders recursively, reconstructs RFC 5322 MIME with attachments, and preserves Read, Flagged, Replied, and Forwarded flags |
| **8. MH / Claws Mail / Sylpheed** | Directories with numbered message files (`1`, `2`, …) and `.mh_sequences` | Parses `.mh_sequences` (`unseen`, `flagged`/`tick`/`mark`, `replied`/`answered`, `forwarded`, `deleted`) |
| **9. MMDF & Emacs Babyl** | `\x01\x01\x01\x01`-delimited MMDF files and `BABYL OPTIONS:` Rmail files | Streaming MMDF parser and Babyl attribute header (`unseen`, `answered`, `flagged`, `forwarded`, `deleted`) mapping |
| **10. Compressed Files & Archives** | `.gz`, `.bz2`, `.xz`, `.zip`, `.tar`, `.tar.gz`, `.tgz`, `.tar.bz2`, `.tar.xz` | Streams compressed mailboxes (`.mbox.gz`, Google Takeout `.zip`/`.tgz`, zipped Maildir/EML/MSG trees) without manual extraction |

---

## Installation & Requirements

- **Python 3.9+**
- **Windows** with classic **Microsoft Outlook** desktop (M365 / 2021 / 2019 / 2016) when reading or writing `.pst` files (`mail2pst` or `.pst` input)

```bash
pip install -r requirements.txt
# Or install as CLI commands (mail2mail, mail2eml, mail2mbox, mail2pst):
pip install -e .
```

### System & Environment Preparation Tips (Windows & Outlook)

To ensure maximum throughput, zero data loss, and uninterrupted multi-gigabyte migrations, apply the following Windows and Outlook environment preparations:

#### 1. Enable Windows Long Path Support (`MAX_PATH` 260-char Bypass)
By default, Windows enforces a legacy 260-character maximum path limit. Deeply nested mail trees (e.g. `Thunderbird/ImapMail/account/[Gmail]/...` or long message subjects) can exceed this. While `mail2mail` uses extended-length prefixes (`\\?\`) internally, enabling system-wide long paths prevents issues with third-party utilities, standard Python scripts, and Windows Explorer:
```powershell
# Run in PowerShell as Administrator:
New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force
```

#### 2. Remove Outlook's 20 MB / 25 MB Attachment Size Limit
By default, classic desktop Outlook rejects attaching files larger than 20–25 MB with error `(4096, "The file you're attaching is bigger than the server allows...")`. To lift this limit and allow unlimited attachment sizes in Outlook:
```powershell
# Set MaximumAttachmentSize to 0 (unlimited) in HKCU:
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences" -Name "MaximumAttachmentSize" -Value 0 -PropertyType DWORD -Force
```
*(Replace `16.0` with `15.0` for Outlook 2013 or `14.0` for Outlook 2010).*

#### 3. Expand Outlook's 50 GB Single-PST Capacity Limit (Optional)
Classic Outlook's default Unicode `.pst` format ceiling is 50 GB (warning at 47.5 GB, raising COM error `0x8004011D` when exceeded). If your migration targets a single large PST beyond 50 GB rather than using `--max-pst-gb 15` or `--split-by year`, expand the MAPI store limit (e.g. to 100 GB):
```powershell
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\PST"
New-Item -Path $regPath -Force | Out-Null
# Set maximum size to 100 GB (102400 MB) and warning threshold to 95 GB (97280 MB):
New-ItemProperty -Path $regPath -Name "MaxLargeFileSize" -Value 102400 -PropertyType DWORD -Force
New-ItemProperty -Path $regPath -Name "WarnLargeFileSize" -Value 97280 -PropertyType DWORD -Force
```

#### 4. Add Real-Time Antivirus Folder Exclusion (Windows Defender)
When migrating large mail archives (50,000–500,000+ `.eml` or `.msg` files), real-time antivirus scanning intercepts every file creation, reducing I/O throughput by 5× to 10×. Adding a temporary folder exclusion on your migration directory significantly boosts conversion speed:
```powershell
# Run in PowerShell as Administrator to exclude the migration workspace:
Add-MpPreference -ExclusionPath "D:\Migration"
```

#### 5. Put Outlook in "Work Offline" Mode
Before starting a large PST import, open Outlook and toggle **Send / Receive** -> **Work Offline**. This prevents background Exchange, Microsoft 365, or IMAP synchronization from competing for MAPI locks or slowing down COM automation calls.

#### 6. Architecture (Bitness) Matching
The Python interpreter's bitness must match your installed Microsoft Outlook's bitness:
- **64-bit Python** -> **64-bit Outlook** (standard default on modern Windows 10/11 & Microsoft 365)
- **32-bit Python** -> **32-bit Outlook** (older Click-to-Run or legacy Office installs)

Verify your Python bitness in PowerShell:
```powershell
python -c "import struct; print(struct.calcsize('P') * 8)"
```
This value (32 or 64) must match Outlook's architecture under *File -> Office Account -> About Outlook*.

---

## Recommended Migration Workspace Structure

For clean, auditable, and easily resumable conversions—especially on multi-gigabyte mailboxes—we recommend establishing a dedicated migration workspace (e.g. on a fast SSD drive `D:\Migration`):

```text
D:\Migration\
├── 00_Source\                  <-- Raw source archives (Read-Only: Takeout .zip, .mbox, Thunderbird)
│   ├── GoogleTakeout_Part1.zip
│   ├── GoogleTakeout_Part2.zip
│   ├── Thunderbird_Profile\
│   └── OldArchive.olm
│
├── 01_Output_PST\              <-- Destination for generated Outlook .pst files
│   ├── Mailbox_2022.pst
│   ├── Mailbox_2023.pst
│   └── Mailbox_2024.pst
│
├── 02_Output_EML\              <-- Destination for .eml hierarchies (when exporting to EML)
│   └── [Folder Tree / YYYY / MM]
│
├── 03_Output_MBOX\             <-- Destination for .mbox / .mbox.gz backups
│   └── AllMail.mbox.gz
│
└── 04_Audit_Logs\              <-- Chain-of-custody audit manifests and failure reports
    ├── audit_manifest.csv      <-- Per-message SHA-256 hashes, timestamps, statuses
    ├── failures.jsonl          <-- Machine-readable error logs (if any unparseable mail)
    └── migration_ops.log       <-- Verbose operational console log mirror
```

**Why this structure is recommended:**
- **Source Immutability**: Keeping raw sources in `00_Source` ensures original mail archives are never modified, deleted, or mixed with output data.
- **Easy Antivirus Exclusion**: You can whitelist the single parent folder (`D:\Migration`) in Windows Defender without exposing other personal directories.
- **Resumability & Audit Isolation**: Pointing `--resume` and `--audit-log` to `04_Audit_Logs\audit_manifest.csv` keeps migration state cleanly decoupled from data files, allowing seamless resumption after system restarts.
- **I/O Performance Optimization**: High-throughput write folders (`01_Output_PST`, `02_Output_EML`) and scratch temp space can reside on a fast local NVMe SSD, while large static source files in `00_Source` can sit on higher-capacity secondary storage.

---

## Usage Guide

### 1) Unified `mail2mail` CLI

`mail2mail.py` dispatches to `mail2eml`, `mail2mbox`, or `mail2pst` either via subcommands (`eml`, `mbox`, `pst`), `--to {eml,mbox,pst}`, or automatically inferred from the destination filename extension:

```cmd
:: Convert any source -> Outlook PST
python mail2mail.py pst "D:\Migration\00_Source\ImapMail" "D:\Migration\01_Output_PST\Mailbox.pst"
python mail2mail.py "D:\Migration\00_Source\export.olm" "D:\Migration\01_Output_PST\MacMail.pst"

:: Convert any source (including PST or EML) -> MBOX or compressed .mbox.gz
python mail2mail.py mbox "D:\Migration\01_Output_PST\MacMail.pst" "D:\Migration\03_Output_MBOX\MacMail.mbox.gz"
python mail2mail.py "D:\Migration\02_Output_EML" "D:\Migration\03_Output_MBOX\AllMail.mbox"

:: Convert any source (including PST or MBOX) -> EML hierarchy
python mail2mail.py eml "D:\Migration\00_Source\inbox.mbox.gz" "D:\Migration\02_Output_EML" --layout month
python mail2mail.py "D:\Migration\01_Output_PST\MacMail.pst" "D:\Migration\02_Output_EML"
```

---

### 2) Convert Any Format → EML (`mail2eml.py`)

#### Preserve 1:1 Folder Hierarchy (Thunderbird, Maildir, Apple Mail, `.olm`, `.pst`, `.zip`)
```cmd
python mail2eml.py ^
  "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" ^
  "D:\Migration\02_Output_EML" ^
  --layout folder ^
  --naming index
```

#### Organize by Year or Month with Date & Size Filters
```cmd
python mail2eml.py ^
  --src "D:\Migration\00_Source\inbox.mbox.gz" ^
  --out-dir "D:\Migration\02_Output_EML" ^
  --layout month ^
  --start-year 2015 ^
  --end-year 2025 ^
  --max-per-dir 50000
```

**Options (`mail2eml.py`):**
- `src` / `--src` / `--mbox`: Source mail directory, mailbox file, or archive.
- `out` / `--out-dir` / `--out`: Destination directory for `.eml` files.
- `--layout {folder,year,month,flat}`: Organize exported `.eml` files by source folder hierarchy, `YYYY`, `YYYY\MM`, or flat.
- `--naming {subject,index}`: Name `.eml` files as `<subject>__<hash>.eml` (`subject`) or `000001.eml` (`index`).
- `--since DATE` / `--after DATE`: Only export messages on or after this date (e.g. `2023-01-01`, `2023-06-15T09:00:00`, or RFC 2822).
- `--until DATE` / `--before DATE`: Only export messages on or before this date (e.g. `2024-12-31`).
- `--include-folder PATTERN`: Only export folders matching glob pattern(s) (can be specified multiple times; e.g. `--include-folder "*Inbox*" --include-folder "Projects/*"`).
- `--exclude-folder PATTERN`: Exclude folders matching glob pattern(s) (e.g. `--exclude-folder "*Trash*" --exclude-folder "*Spam*"`).
- `--state-file FILE`: Persistent SQLite (`.sqlite`, `.db`) or JSON Lines (`.jsonl`) state database for high-speed incremental / delta cutover syncs.
- `--split-size SIZE`: Roll over directory when reaching human-friendly size limit (e.g. `20GB`, `500MB`, `15G`).
- `--start-year YYYY` / `--end-year YYYY`: Filter messages by `Date` header year (inclusive).
- `--max-per-dir N`: Roll over to `__part2`, `__part3`, etc. when a directory reaches `N` files (`0` = unlimited).
- `--max-dir-bytes GB`: Roll over when a directory reaches `GB` gigabytes (`0` = unlimited).
- `--include-deleted`: Export messages marked as deleted/expunged.
- `--progress-every N`: Print progress every `N` messages (default: `1000`).
- `--audit-log FILE`: Generate structured audit manifest (`.csv` or `.jsonl`) with message metadata and SHA-256 hashes.
- `--error-log FILE`: Record machine-readable failure/warning entries (`.jsonl`).
- `--log-file FILE`: Mirror all operations with timestamps into a persistent log file.
- `--dedup [MODE]`: Deduplicate messages during migration (`auto`, `message-id`, or `sha256`; default: `none`).
- `--resume [FILE]`: Resume an interrupted export, skipping messages already recorded with `status=SUCCESS` in an existing audit manifest.
- `--no-gmail-labels`: Disable automatic Google Takeout `X-Gmail-Labels` folder hierarchy extraction.
- `--dry-run`: Pre-flight inspection reporting folder counts, date ranges, and oversized attachments without writing to disk.
- `-v` / `--verbose`: Enable verbose logging output.
- `-q` / `--quiet`: Suppress live progress counters, printing only the final summary and errors.

---

### 3) Convert Any Format → MBOX (`mail2mbox.py`)

#### Single `.mbox` or Compressed `.mbox.gz` Output
```cmd
python mail2mbox.py "D:\Migration\02_Output_EML" "D:\Migration\03_Output_MBOX\AllMail.mbox"
python mail2mbox.py "D:\Migration\01_Output_PST\Archive.pst" "D:\Migration\03_Output_MBOX\Archive.mbox.gz"
```

#### Multi-File Folder or Year MBOX Tree
```cmd
python mail2mbox.py "D:\Migration\00_Source\ImapMail" "D:\Migration\03_Output_MBOX\MboxTree" --layout folder
python mail2mbox.py "D:\Migration\00_Source\ImapMail" "D:\Migration\03_Output_MBOX\ByYear" --layout year --max-mbox-gb 4
```

**Options (`mail2mbox.py`):**
- `src` / `--src` / `--mbox`: Source mail directory, mailbox file, or archive.
- `out` / `--out` / `--out-dir`: Output `.mbox` (`.mbox.gz` / `.mbox.bz2` / `.mbox.xz`) file or destination directory.
- `--layout {folder,year,month,flat}`: Organize multi-mbox output by folder hierarchy, year, month, or flat (default: `folder`).
- `--since DATE` / `--after DATE`: Only export messages on or after this date (e.g. `2023-01-01`, `2023-06-15T09:00:00`, or RFC 2822).
- `--until DATE` / `--before DATE`: Only export messages on or before this date (e.g. `2024-12-31`).
- `--include-folder PATTERN`: Only export folders matching glob pattern(s) (can be specified multiple times; e.g. `--include-folder "*Inbox*" --include-folder "Projects/*"`).
- `--exclude-folder PATTERN`: Exclude folders matching glob pattern(s) (e.g. `--exclude-folder "*Trash*" --exclude-folder "*Spam*"`).
- `--state-file FILE`: Persistent SQLite (`.sqlite`, `.db`) or JSON Lines (`.jsonl`) state database for high-speed incremental / delta cutover syncs.
- `--split-size SIZE` / `--max-mbox-size SIZE`: Rotate to `__part2.mbox` when an mbox exceeds human-friendly size limit (e.g. `4GB`, `500MB`, `10G`).
- `--max-mbox-gb GB`: Rotate to `__part2.mbox` when an mbox exceeds `GB` gigabytes (`0` = unlimited).
- `--start-year YYYY` / `--end-year YYYY`: Filter messages by `Date` header year (inclusive).
- `--include-deleted`: Include messages marked as deleted/expunged.
- `--append`: Append to existing `.mbox` file(s) instead of overwriting.
- `--audit-log FILE`: Generate structured audit manifest (`.csv` or `.jsonl`) with message metadata and SHA-256 hashes.
- `--error-log FILE`: Record machine-readable failure/warning entries (`.jsonl`).
- `--log-file FILE`: Mirror all operations with timestamps into a persistent log file.
- `--dedup [MODE]`: Deduplicate messages during migration (`auto`, `message-id`, or `sha256`; default: `none`).
- `--resume [FILE]`: Resume an interrupted export, skipping messages already recorded with `status=SUCCESS` in an existing audit manifest.
- `--no-gmail-labels`: Disable automatic Google Takeout `X-Gmail-Labels` folder hierarchy extraction.
- `--dry-run`: Pre-flight inspection reporting folder counts, date ranges, and oversized attachments without writing to disk.
- `-v` / `--verbose` & `-q` / `--quiet`: Control console output verbosity.

---

### 4) Convert Any Format → Outlook PST (`mail2pst.py`)

#### Mode A: Direct Single-PST Output
```cmd
python mail2pst.py ^
  "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" ^
  "D:\Migration\01_Output_PST\Mailbox.pst"
```

#### Mode B: Split by Year (One PST per Year, Auto-Rotating at 15 GB)
```cmd
python mail2pst.py ^
  --src "D:\Migration\02_Output_EML" ^
  --out-dir "D:\Migration\01_Output_PST" ^
  --base-name emails ^
  --split-by year ^
  --max-pst-gb 15 ^
  --pst-root "Imported (EML)" ^
  --flush-every 5000 ^
  --count-every 200
```

#### Mode C: Even-Split into N PSTs by Total Size
```cmd
python mail2pst.py ^
  --src "D:\Migration\02_Output_EML" ^
  --out-dir "D:\Migration\01_Output_PST" ^
  --base-name emails ^
  --splits 6 ^
  --max-pst-gb 18 ^
  --pst-root "Imported (EML)"
```

**Options (`mail2pst.py`):**
- `src` / `--src` / `--mbox`: Source folder, mailbox file, or archive (`.eml`, `.emlx`, `.msg`, `.olm`, `.pst`, Maildir, MH, MBOX, MMDF, Babyl, `.zip`, `.tgz`, `.gz`).
- `out` / `--out` / `--pst`: Output `.pst` file path (for single-PST mode).
- `--out-dir DIR`: Output directory for multi-PST split/rotated mode.
- `--base-name NAME`: Base filename prefix when using `--out-dir` (default: `emails`).
- `--split-by year`: Create one PST per message year (`<base>_<YYYY>.pst`, rolling over to `<base>_<YYYY>_part2.pst` if exceeding `--max-pst-gb`).
- `--splits N`: Divide total byte volume across `N` sequential PST files (`<base>.pst`, `<base>_part2.pst`, …).
- `--split-pst SIZE` / `--max-pst-size SIZE`: Maximum size per PST in human units (e.g. `15GB`, `20G`, `500MB`) before rotating to the next part.
- `--max-pst-gb GB`: Maximum size in GB per PST before rotating to the next part (default: `15.0`).
- `--since DATE` / `--after DATE`: Only export messages on or after this date (e.g. `2023-01-01`, `2023-06-15T09:00:00`, or RFC 2822).
- `--until DATE` / `--before DATE`: Only export messages on or before this date (e.g. `2024-12-31`).
- `--include-folder PATTERN`: Only export folders matching glob pattern(s) (can be specified multiple times; e.g. `--include-folder "*Inbox*" --include-folder "Projects/*"`).
- `--exclude-folder PATTERN`: Exclude folders matching glob pattern(s) (e.g. `--exclude-folder "*Trash*" --exclude-folder "*Spam*"`).
- `--state-file FILE`: Persistent SQLite (`.sqlite`, `.db`) or JSON Lines (`.jsonl`) state database for high-speed incremental / delta cutover syncs.
- `--pst-root NAME`: Top-level folder inside each PST (default: root `""` in single-PST mode, `"Imported (EML)"` in `--out-dir` mode).
- `--flatten`: Place all messages directly inside `--pst-root` without recreating source subfolders.
- `--start-year YYYY` / `--end-year YYYY`: Filter messages by `Date` header year (inclusive).
- `--overwrite-pst`: Delete and recreate target PST files if they already exist on disk.
- `--include-deleted`: Import messages marked as deleted/expunged.
- `--flush-every N`: Detach and re-attach the current PST every `N` items so Windows Explorer updates file size (`0` = off).
- `--count-every N`: Print live folder item count every `N` items (default: `200`, `0` = off).
- `--audit-log FILE`: Generate structured audit manifest (`.csv` or `.jsonl`) with message metadata and SHA-256 hashes.
- `--error-log FILE`: Record machine-readable failure/warning entries (`.jsonl`).
- `--log-file FILE`: Mirror all operations with timestamps into a persistent log file.
- `--dedup [MODE]`: Deduplicate messages during migration (`auto`, `message-id`, or `sha256`; default: `none`).
- `--resume [FILE]`: Resume an interrupted migration, skipping messages already recorded with `status=SUCCESS` in an existing audit manifest.
- `--no-gmail-labels`: Disable automatic Google Takeout `X-Gmail-Labels` folder hierarchy extraction.
- `--dry-run`: Pre-flight inspection reporting folder counts, date ranges, and oversized attachments without writing to disk/PST.
- `-v` / `--verbose` & `-q` / `--quiet`: Control console output verbosity.

### 5) Fastmail & Modern JMAP RFC 8620/8621 Ingestion (`mail2jmap.py` or `mail2mail jmap`)
Directly ingest email from **Fastmail** or any RFC 8620/8621 compliant JMAP server (Stalwart, Apache James, Cyrus) using standard API Bearer tokens:

```bash
# Ingest entire Fastmail account into an Outlook PST
python mail2jmap.py --token "fmu1-xxxx-xxxx" --out "D:\PSTs\Fastmail.pst"

# Sync specific folders into a compressed MBOX file with incremental checkpointing
python mail2jmap.py --token "fmu1-xxxx-xxxx" --folders "Inbox,Archive" --checkpoint "fastmail_sync.json" --out "fastmail.mbox.gz"

# Unified CLI invocation
mail2mail jmap --token "fmu1-xxxx-xxxx" --out "fastmail_backup.pst"

# Ingest static exported .jmap or .jmap.json backup files
mail2mail eml "fastmail_export.jmap" "D:\Export_EML"
```

Key JMAP Capabilities:
- **Zero Third-Party Dependencies**: Native Python client using standard library `urllib.request` and `json`.
- **RFC 8620 Core Discovery**: Automatic session negotiation via `https://api.fastmail.com/jmap/session` or custom URLs.
- **RFC 8621 Mail Batching**: High-throughput querying (`Email/query`) and retrieval (`Email/get`) fetching up to 200–500 emails per request.
- **Rich Body Extraction**: Pulls `textBody` and `htmlBody` directly from structured JSON payloads.
- **Status Flag & Keyword Mapping**: `$seen`, `$flagged`, `$answered`, and `$draft` keywords mapped directly to MAPI/MIME status flags.
- **Incremental Resumption**: Checkpoint tracking (`--checkpoint`) compatible with `mail2rust`.

---

### 6) Model Context Protocol (MCP) Server for AI Agents (`mail2mail mcp` or `mail2mcp.py`)

`mail2mail` includes a built-in [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server allowing LLM agents (Claude Desktop, Cursor, Google Antigravity, AutoGPT) to inspect, classify, convert, audit, and split mailbox archives over standard JSON-RPC 2.0 stdio.

#### Exposed Agent Tools:
1. **`mail2mail_info`**: Inspects platform, Python version, and Outlook COM availability.
2. **`mail2mail_sniff`**: Classifies 15+ email archive formats via magic bytes.
3. **`mail2mail_convert`**: Converts mailboxes between EML, MBOX, and PST (with graceful cross-platform handling).
4. **`mail2mail_audit`**: Forensic auditing: calculates SHA-256 Merkle root, applies Bates stamping, and exports CSV manifests.
5. **`mail2mail_split`**: Splits large mailboxes by size or year into chunked volumes.

#### Running the MCP Server:
```bash
# Run stdio MCP server for agent clients
mail2mail mcp

# Or via dedicated script
python mail2mcp.py

# Inspect tool JSON schemas
mail2mail mcp --tools
```

#### Claude Desktop & Cursor Configuration:
Add to `claude_desktop_config.json` or `.cursor/mcp.json`:
```json
{
  "mcpServers": {
    "mail2mail": {
      "command": "python",
      "args": ["-m", "mail2mcp"]
    }
  }
}
```

> **Note on PST Export**: PST creation in `mail2mail` requires Microsoft Outlook on Windows via `pywin32` COM. For headless cloud VMs or cross-platform (Linux/macOS) native MS-PST binary generation with zero Outlook dependency, see our companion project [**`mail2rust`**](https://github.com/adpaolucci/mail2rust).

---

## Auditability, Pre-Flight Dry-Run & Forensic Logging

`mail2mail` includes built-in forensic auditing, manifest generation, and pre-flight inspection tools suited for compliance, legal eDiscovery, and enterprise migrations:

### 1) Pre-Flight Dry-Run Audit (`--dry-run`)
Inspect a source mail file, folder, or multi-GB archive before performing any writes:
```cmd
python mail2mail.py --dry-run "D:\Migration\00_Source\GoogleTakeout.zip"
```
Surfaces:
- Total message count, payload size (GB/MB), and status breakdown (Read, Unread, Flagged, Deleted).
- Date span (earliest and latest message date) and breakdown by year.
- Folder hierarchy and volume distribution.
- **Oversized Attachment Warnings**: Scans for attachments exceeding 20 MB / 25 MB that would trigger Outlook server limits.

### 2) Structured Audit Manifest (`--audit-log <manifest.csv | manifest.jsonl>`)
Generates an audit trail of every evaluated message:
- Fields: `timestamp`, `status` (`SUCCESS`, `SKIPPED_DATE`, `SKIPPED_DELETED`, `SKIPPED_DUPLICATE`, `FAILED`), `message_id`, `subject`, `date`, `sender`, `recipients`, `source_folder`, `target_location`, `size_bytes`, `attachment_count`, and `sha256` digest for chain-of-custody verification.

```cmd
python mail2mail.py pst "D:\Migration\00_Source\ImapMail" "D:\Migration\01_Output_PST\Mail.pst" --audit-log "D:\Migration\04_Audit_Logs\audit.csv" --log-file "D:\Migration\04_Audit_Logs\migration_ops.log"
```

### 3) Dedicated Failure Log (`--error-log <failures.jsonl>`)
Appends machine-readable JSON entries with exact message subjects, Message-IDs, error types, and exception details for any skipped or failed messages, eliminating manual log hunting.

### 4) In-Flight Deduplication (`--dedup {auto | message-id | sha256}`)
Prevents importing duplicate messages when scanning multi-part archives (e.g. split Takeout `.zip` files) or merging multiple backup roots. Duplicates are tracked and logged in the manifest as `SKIPPED_DUPLICATE`:
```cmd
python mail2mail.py pst "D:\Migration\00_Source\Takeout_Parts" "D:\Migration\01_Output_PST\Unified.pst" --dedup auto
```

### 5) Resuming Interrupted Migrations (`--resume [manifest]`)
Large migrations (hundreds of gigabytes) can be resumed after system reboots or power interruptions without duplicating already-imported records. `--resume` reads existing `SUCCESS` entries from the audit manifest and skips them on the next run:
```cmd
python mail2mail.py pst "D:\Migration\00_Source\LargeMailbox" "D:\Migration\01_Output_PST\Mailbox.pst" --audit-log "D:\Migration\04_Audit_Logs\audit.csv" --resume
```

### 6) Incremental Cutover & Delta Sync (`--state-file <path.sqlite | path.jsonl>`)
In production enterprise cutovers, an initial bulk migration is performed while users remain on the old mail server. During the final cutover weekend, you run an **incremental delta sync** that fast-skips all previously migrated messages and imports only newly arrived emails:
```cmd
:: Pass 1: Bulk historical migration (builds state database)
python mail2mail.py pst "D:\Migration\00_Source\LiveMailbox" "D:\Migration\01_Output_PST\Mailbox.pst" --state-file "D:\Migration\04_Audit_Logs\migration_state.sqlite"

:: Pass 2: Final cutover delta sync (processes only messages received since Pass 1)
python mail2mail.py pst "D:\Migration\00_Source\LiveMailbox" "D:\Migration\01_Output_PST\Mailbox.pst" --state-file "D:\Migration\04_Audit_Logs\migration_state.sqlite"
```
The state cache engine indexes both `Message-ID` headers and byte-accurate SHA-256 hashes using a high-concurrency SQLite WAL database or line-delimited JSONL stream, providing $O(1)$ lookup speed across millions of records.

### 7) Granular Date Range & Folder Filtering (`--since`, `--until`, `--include-folder`, `--exclude-folder`)
Selectively migrate date intervals or folder subtrees with timezone-aware ISO / RFC 2822 timestamps and case-insensitive glob patterns:
```cmd
:: Migrate only calendar year 2024, excluding Trash and Spam folders
python mail2mail.py pst "D:\Migration\00_Source\ImapMail" "D:\Migration\01_Output_PST\Mail2024.pst" ^
  --since "2024-01-01" ^
  --until "2024-12-31" ^
  --exclude-folder "*Trash*" ^
  --exclude-folder "*Spam*"

:: Target only specific department or project archives
python mail2mail.py eml "D:\Migration\00_Source\CompanyArchive.olm" "D:\Migration\02_Output_EML" ^
  --include-folder "Projects/*" ^
  --include-folder "*Finance*"
```

### 8) Human-Friendly Capacity & Size Splitting (`--split-pst`, `--split-size`)
Specify capacity thresholds using intuitive size units (`GB`, `MB`, `G`, `M`) instead of raw byte counts:
```cmd
:: Rotate to part2, part3 when PST reaches 15 GB
python mail2mail.py pst "D:\Migration\00_Source\LargeArchive" --out-dir "D:\Migration\01_Output_PST" --split-pst 15GB

:: Rotate mbox files at 4 GB
python mail2mail.py mbox "D:\Migration\02_Output_EML" "D:\Migration\03_Output_MBOX\MboxTree" --split-size 4G
```

### 9) Graceful Interruption, Pause/Resume & Signal Handling (`Ctrl+C`, `SIGTERM`)
`mail2mail` incorporates an interruption controller ([`mail_shutdown.py`](mail_shutdown.py)) ensuring zero data loss and clean state preservation when halting long-running migrations:
- **Signal Trapping (`Ctrl+C` / `SIGINT`, `SIGTERM`, `SIGBREAK`)**:
  - Instead of abruptly terminating mid-write, the controller allows the single in-flight email to finish saving cleanly.
  - Commits all batched records to the `--state-file` cache and flushes `--audit-log` manifests.
  - Gracefully unmounts all Outlook PST stores (`router.close_all()`) so Outlook flushes indices and releases file locks without leaving orphan handles or stuck 265 KB stores.
  - Flushes all open MBOX file streams and EML directory handles.
  - Prints the full reconciliation summary and displays a resumption banner showing the exact command needed to resume.
  - Returns standard exit code `130`.
- **Double-Tap Emergency Termination**:
  - Pressing `Ctrl+C` a second time triggers an immediate emergency exit (`os._exit(130)`).
- **Interactive Console Pause & Resume**:
  - When running in an interactive terminal (TTY), press **`p`** or **`[Space]`** at any point to temporarily pause the conversion.
  - Press **`[Space]`** or **`[Enter]`** to resume processing immediately.
  - Press **`q`** or **`Ctrl+C`** to halt gracefully and commit all progress to disk.

---

## Testing & RFC Compliance Reference

### Running the Automated Test Suite
The [`tests/`](tests/) package verifies RFC compliance, all 10 mailbox formats, and all CLI exporters (`mail2eml`, `mail2mbox`, `mail2pst`, `mail2mail`):

```bash
python -m unittest discover -s tests -v
```

### RFC & Format Specifications Reference (`docs/rfc_compliance/`)
Detailed technical documentation mapping IETF RFCs and Microsoft Open Specifications to the codebase is located in [`docs/rfc_compliance/`](docs/rfc_compliance/README.md):
- [`docs/rfc_compliance/README.md`](docs/rfc_compliance/README.md) — Master standards compliance matrix
- [`docs/rfc_compliance/RFC5322_RFC2045_MIME_AND_HEADERS.md`](docs/rfc_compliance/RFC5322_RFC2045_MIME_AND_HEADERS.md) — RFC 5322, RFC 2045–2047, RFC 2183/2231/2392 (`cid:` MHTML), S/MIME & PGP/MIME
- [`docs/rfc_compliance/RFC4155_MBOX_AND_MAILBOX_FORMATS.md`](docs/rfc_compliance/RFC4155_MBOX_AND_MAILBOX_FORMATS.md) — RFC 4155 (`mbox` / `mboxrd`), RFC 3501 §5.1.3 (IMAP Modified UTF-7), Thunderbird, Maildir++, Apple Mail `.emlx`, MH, MMDF, and Babyl
- [`docs/rfc_compliance/MS_OXMSG_CFBF_MAPI_AND_OLM.md`](docs/rfc_compliance/MS_OXMSG_CFBF_MAPI_AND_OLM.md) — [MS-CFB] (OLE2), [MS-OXMSG] (`.msg`), [MS-OXPROPS] (MAPI DASL properties & `PT_SYSTIME`), and Outlook for Mac `.olm`
- [`docs/rfc_compliance/RFC8620_RFC8621_JMAP.md`](docs/rfc_compliance/RFC8620_RFC8621_JMAP.md) — RFC 8620 (JMAP Core Architecture, Session Discovery, JSON Batching) & RFC 8621 (JMAP for Mail, Keywords, Threads, Fastmail & Modern Mail Servers)

---

## Troubleshooting

- **PST file size appears stuck at ~265 KB in Windows Explorer while running**
  - Outlook holds an open write handle on mounted `.pst` files and flushes metadata on unmount. `mail2pst` automatically detaches all PST stores on completion (and during `--flush-every`), at which point Explorer shows the final size.
- **Classic Outlook vs. "New Outlook"**
  - The web-based "New Outlook" for Windows does not support MAPI/COM automation (`Outlook.Application`). Switch to classic desktop Outlook before running `mail2pst` or reading `.pst` files.

---

## Credits & Attribution

`mail2mail` is built on a rich foundation of open-source projects, community fork contributions, and collaborative AI pair-programming:
- **Project Maintainer**: Anthony Paolucci ([@adpaolucci](https://github.com/adpaolucci))
- **AI Collaborators**: Argon 4 (Anthropic) & Antigravity (Google DeepMind)
- **Original Upstream Origin**: Rick Madson ([@madsonrick](https://github.com/madsonrick/mbox-to-pst-toolkit))
- **Community Fork Contributors**: [@gbordier](https://github.com/gbordier), [@HTCGS](https://github.com/HTCGS), [@enthusedamateur](https://github.com/enthusedamateur), [@dianlight](https://github.com/dianlight), [@icmaa](https://github.com/icmaa), [@Erzmaster](https://github.com/Erzmaster), [@ITKKDF](https://github.com/ITKKDF)
- **Related Prior Art**: `3en5` ([KALIT/Natanor IT Services](https://github.com/3en5/EML-to-PST)), `SSDECT` ([@ssdect](https://github.com/ssdect/eml-to-pst)), and `MimeKit`/`MsgKit`

For complete attributions, architectural histories, and technical details, see [**`CONTRIBUTORS.md`**](CONTRIBUTORS.md).

---

## License

Released under the [MIT License](LICENSE).
