# Universal Email & Mailbox → EML → PST Toolkit (`all2pst`)

Export and convert **MBOX**, **Thunderbird `ImapMail` / `Mail`** (mbox and Maildir modes), **Maildir / Maildir++**, **Apple Mail (`.emlx` / `.mbox` bundles)**, **Outlook `.msg`**, **Outlook for Mac `.olm`**, **MH / Sylpheed / Claws Mail**, **MMDF**, **Emacs Babyl**, **`.eml` trees**, and **compressed archives (`.gz`, `.bz2`, `.xz`, `.zip`, `.tar.gz`, `.tgz`)** into `.eml` folder hierarchies or directly into Microsoft Outlook **.pst** archives with full folder hierarchy preservation, message flag mapping, size/year splitting, and resilient Outlook MAPI integration.

---

## Supported Email & Mailbox Formats

All four CLI tools share the unified [`mail_formats.py`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py) reader and automatically detect any of the following formats (using only the Python standard library + `pywin32` for `.pst` output):

| Format | Supported Inputs | Metadata & Status Flags Preserved |
| :--- | :--- | :--- |
| **1. Standard MBOX** (`mboxo` / `mboxrd`) | `.mbox` files, extensionless mbox files | `X-Mozilla-Status`, `Status`, `X-Status`, `>From ` un-escaping |
| **2. Thunderbird `ImapMail` / `Mail`** | Account folders, `ImapMail`/`Mail` roots, `.sbd` hierarchies (both **mbox** and **Maildir** modes) | `X-Mozilla-Status` (`Read`, `Replied`, `Flagged`, `Forwarded`, `Expunged`), IMAP Modified UTF-7 folder decoding |
| **3. Maildir & Maildir++** | `cur/`, `new/`, `tmp/`, `.Folder.Subfolder` Maildir++ trees | `:2,FRSPT` / `;2,FRSPT` filename flags (`S`=Seen, `R`=Replied, `F`=Flagged, `P`=Forwarded, `T`=Trashed), `new/` unread state |
| **4. Apple Mail (`.emlx` & `.mbox` Bundles)** | `.emlx`, `.partial.emlx`, macOS `*.mbox` / `*.imapmbox` / `*.popmbox` directory bundles | Strips byte-count prefix & parses XML `<plist>` `flags` bitmask (`Read`, `Deleted`, `Replied`, `Flagged`, `Forwarded`); collapses internal `Data/0..9/Messages` shards |
| **5. Outlook `.msg` (MS-CFB / MS-OXMSG)** | `.msg` files and folder trees | Native MAPI `OpenSharedItem` transfer when importing to PST, plus pure-Python MS-CFB OLE2 + MS-OXMSG parser for `.msg` → `.eml` conversion without Outlook |
| **6. Outlook for Mac `.olm`** | `.olm` archive files | Extracts `com.microsoft.__Messages` folder hierarchy, HTML/plain bodies, inline `cid:` & file attachments, and Read/Flagged state |
| **7. MH / Claws Mail / Sylpheed** | Directories with numbered message files (`1`, `2`, …) and `.mh_sequences` | Parses `.mh_sequences` (`unseen`, `flagged`/`tick`/`mark`, `replied`/`answered`, `forwarded`, `deleted`) |
| **8. MMDF & Emacs Babyl** | `\x01\x01\x01\x01`-delimited MMDF files and `BABYL OPTIONS:` Rmail files | Streaming MMDF parser and Babyl attribute header (`unseen`, `answered`, `flagged`, `forwarded`, `deleted`) mapping |
| **9. Compressed Files & Archives** | `.gz`, `.bz2`, `.xz`, `.zip`, `.tar`, `.tar.gz`, `.tgz`, `.tar.bz2`, `.tar.xz` | Streams compressed mailboxes (`.mbox.gz`, Google Takeout `.zip`/`.tgz`, zipped Maildir/EML/MSG trees) without manual extraction |

---

## Included Scripts

| Script | Purpose | Requires Outlook? |
| :--- | :--- | :--- |
| [`mail_formats.py`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py) | Shared multi-format detection, streaming, flag normalization, and folder hierarchy engine | No (Stdlib only) |
| [`mbox_to_eml_exporter.py`](file:///c:/Users/andre/Desktop/all2pst/mbox_to_eml_exporter.py) | Export `.mbox` (or any supported mailbox/archive) into `.eml` directories (`year`, `month`, `flat`, or `folder` layout) with date filters and per-directory file/size caps | No (Stdlib only) |
| [`imapmail_to_eml.py`](file:///c:/Users/andre/Desktop/all2pst/imapmail_to_eml.py) | Stream a Thunderbird `ImapMail` / `Mail` directory, Maildir, Apple Mail `.mbox`, `.olm`, `.msg`, MH, or archive into a 1:1 `.eml` folder tree | No (Stdlib only) |
| [`eml_to_pst_import.py`](file:///c:/Users/andre/Desktop/all2pst/eml_to_pst_import.py) | Import a directory or archive of `.eml` (or `.emlx`, `.msg`, `.olm`, Maildir, MH, MBOX, `.zip`/`.tgz`) into one or more Outlook `.pst` files (split by year, even-split into `N` parts, or auto-rotated at `--max-pst-gb`) | Yes (`pywin32`) |
| [`imapmail_to_pst.py`](file:///c:/Users/andre/Desktop/all2pst/imapmail_to_pst.py) | Directly convert a Thunderbird `ImapMail` / `Mail` directory, Maildir, Apple Mail `.mbox`, `.olm`, `.msg`, MH, MBOX, or archive into an Outlook `.pst` archive | Yes (`pywin32`) |

---

## Key Capabilities & Hardening

- **Full Folder Hierarchy Preservation**
  - Traverses Thunderbird `.sbd` subfolder containers, Maildir++ `.Folder.Subfolder` directories, Apple Mail `.mbox` bundles, and `.olm` / `.zip` / `.tgz` archives.
  - Decodes RFC 3501 **IMAP Modified UTF-7** folder names (e.g., `Notes &- Invoices` → `Notes & Invoices`, `Caf&AOk-` → `Café`).
  - Mirrors relative folder structures inside target `.pst` archives (or imports flat when `--flatten` is specified).
- **Unified Message State Mapping**
  - Maps `Read`, `Replied`, `Flagged`, `Forwarded`, and `Deleted`/`Expunged` states across Thunderbird `X-Mozilla-Status`, Maildir `:2,FRSPT` flags, Apple Mail `.emlx` `<plist>` bitmasks, `.olm` XML attributes, `.msg` MAPI properties, MH `.mh_sequences`, and Babyl labels.
  - Automatically skips deleted/expunged/trashed messages by default (override with `--include-deleted`).
  - Preserves Read/Unread state (`MailItem.UnRead`), Follow-up Flags (`FlagStatus = olFlagMarked`), and Replied/Forwarded icons (`PR_LAST_VERB_EXECUTED`, `PR_ICON_INDEX`).
- **High-Fidelity Outlook MAPI Property Injection**
  - Creates items directly in the destination PST folder (`folder.Items.Add(0)`) so items never land in the default profile's `Drafts` folder.
  - Injects naive UTC `PT_SYSTIME` timestamps into `PR_MESSAGE_DELIVERY_TIME` (`0x0E060040`) and `PR_CLIENT_SUBMIT_TIME` (`0x00390040`) after body population so Outlook preserves exact sent/received dates.
  - Populates Unicode MAPI sender and representing properties (`0x0C1A001F`, `0x0C1F001F`, `0x0042001F`, `0x0065001F`), `PR_INTERNET_MESSAGE_ID_W` (`0x1035001F`), and `PR_TRANSPORT_MESSAGE_HEADERS_W` (`0x007D001F`).
  - Preserves inline `cid:` HTML images (`PR_ATTACH_CONTENT_ID_W` + `PR_ATTACH_FLAGS = 4`), `text/*` file attachments (`.csv`, `.txt`, `.ics`), and embedded `message/rfc822` forwarded emails with original filenames intact.
- **Large-Archive Reliability**
  - Automatic COM retry backoff with `pythoncom.PumpWaitingMessages()` when Outlook returns `RPC_E_CALL_REJECTED` (`0x80010001`).
  - Windows `\\?\` long-path support for deep folder trees and dynamic `MAX_PATH` protection for Outlook COM attachment staging.
  - Non-destructive `--flush-every` PST detach/re-attach cycle to force Windows Explorer file size updates without losing data.

---

## Requirements

- **Python 3.9+**
- **Windows** with classic **Microsoft Outlook** desktop (M365 / 2021 / 2019 / 2016) for `.pst` import scripts (`eml_to_pst_import.py` and `imapmail_to_pst.py`)
- `pywin32`:

```bash
pip install -r requirements.txt
```

### Outlook Preparation Tips
1. Ensure classic desktop Outlook has a default mail profile configured that opens without modal prompts.
2. Put Outlook in **Work Offline** mode while running large imports.
3. Avoid browsing or moving items inside the target `.pst` in Outlook while the import is active.

---

## Usage Guide

### 1) Export Mailbox / Archive → EML with Layout & Filters (`mbox_to_eml_exporter.py`)

```cmd
python mbox_to_eml_exporter.py ^
  --mbox "D:\Mail\inbox.mbox.gz" ^
  --out-dir "D:\Export_EML" ^
  --layout month ^
  --sanitize-filenames
```

**Options:**
- `--mbox` / `--src`: Path to source `.mbox`, `.mbox.gz`, `.olm`, `.msg`, `.emlx`, Maildir, MH, MMDF, Babyl, `.zip`, or `.tgz`.
- `--layout {year,month,flat,folder}`: Organize exported `.eml` files into `YYYY`, `YYYY\MM`, flat, or source folder hierarchy (default: `year`).
- `--start-year YYYY` / `--end-year YYYY`: Filter messages by `Date` header year (inclusive).
- `--max-per-dir N`: Roll over to `__part2`, `__part3`, etc. when a directory reaches `N` files (`0` = unlimited).
- `--max-dir-bytes GB`: Roll over when a directory reaches `GB` gigabytes (`0` = unlimited).
- `--include-deleted`: Export messages marked as deleted/expunged.
- `--progress-every N`: Print progress every `N` messages (default: `1000`).

---

### 2) Export Folder Hierarchy → EML (`imapmail_to_eml.py`)

Point `imap_src` at a Thunderbird `ImapMail` / `Mail` folder, Maildir, Apple Mail `.mbox` bundle, Outlook `.olm` archive, `.msg` directory, MH directory, or `.zip`/`.tgz` archive:

```cmd
python imapmail_to_eml.py ^
  "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" ^
  "D:\Export_EML"
```

**Options:**
- `--include-deleted`: Export messages marked as deleted/expunged.

---

### 3) Import EML / Any Format → PST with Splitting (`eml_to_pst_import.py`)

#### Split by Year (one PST per year, auto-rotating at 15 GB)
```cmd
python eml_to_pst_import.py ^
  --src "D:\Export_EML" ^
  --out-dir "D:\PSTs" ^
  --base-name emails ^
  --split-by year ^
  --max-pst-gb 15 ^
  --pst-root "Imported (EML)" ^
  --flush-every 5000 ^
  --count-every 200
```

#### Even-Split into N PSTs by Total Size
```cmd
python eml_to_pst_import.py ^
  --src "D:\Export_EML" ^
  --out-dir "D:\PSTs" ^
  --base-name emails ^
  --splits 6 ^
  --max-pst-gb 18 ^
  --pst-root "Imported (EML)"
```

**Options:**
- `--src`: Source folder, mailbox file, or archive (`.eml`, `.emlx`, `.msg`, `.olm`, Maildir, MH, MBOX, MMDF, Babyl, `.zip`, `.tgz`, `.gz`).
- `--split-by year`: Create one PST per message year (`<base>_<YYYY>_part1.pst`).
- `--splits N`: Divide total byte volume across `N` sequential PST files (`<base>_part1.pst`, …).
- `--max-pst-gb GB`: Maximum size in GB per PST before rotating to the next part (default: `15.0`).
- `--pst-root NAME`: Top-level folder inside each PST (default: `"Imported (EML)"`; pass `""` to place folders directly at the PST root).
- `--flatten`: Place all messages directly inside `--pst-root` without recreating source subfolders.
- `--overwrite-pst`: Delete and recreate target PST files if they already exist on disk.
- `--include-deleted`: Import messages marked as deleted/expunged.
- `--flush-every N`: Detach and re-attach the current PST every `N` items so Windows Explorer updates file size (`0` = off).
- `--count-every N`: Print live folder item count every `N` items (default: `200`, `0` = off).

---

### 4) Direct Folder Hierarchy / Archive → PST (`imapmail_to_pst.py`)

```cmd
python imapmail_to_pst.py ^
  "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" ^
  "D:\PSTs\Mailbox.pst"
```

**Options:**
- `--include-deleted`: Import messages marked as deleted/expunged.

---

## Troubleshooting

- **PST file size appears stuck at ~265 KB in Windows Explorer while running**
  - Outlook holds an open write handle on mounted `.pst` files and flushes metadata on unmount. The scripts automatically detach all PST stores in their `finally` blocks (and during `--flush-every`), at which point Explorer shows the final size.
- **Classic Outlook vs. "New Outlook"**
  - The web-based "New Outlook" for Windows does not support MAPI/COM automation (`Outlook.Application`). Switch to classic desktop Outlook before running `eml_to_pst_import.py` or `imapmail_to_pst.py`.
