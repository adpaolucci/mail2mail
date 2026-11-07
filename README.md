# MBOX & Thunderbird ImapMail → EML → PST Toolkit

Export large `.mbox` mailboxes or full Mozilla Thunderbird `ImapMail` / `Mail` profile trees into `.eml` files and import them into Microsoft Outlook **.pst** archives with full folder hierarchy preservation, Thunderbird message state mapping, size/year splitting, and resilient Outlook MAPI integration.

---

## Included Scripts

| Script | Purpose | Requires Outlook? |
| :--- | :--- | :--- |
| [`mbox_to_eml_exporter.py`](file:///c:/Users/andre/Desktop/all2pst/mbox_to_eml_exporter.py) | Export a single `.mbox` file into `.eml` directories (`year`, `month`, or `flat` layout) with date filters and per-directory file/size caps | No (Stdlib only) |
| [`imapmail_to_eml.py`](file:///c:/Users/andre/Desktop/all2pst/imapmail_to_eml.py) | Stream a Thunderbird `ImapMail` / `Mail` directory (or `.mbox` file) into a 1:1 `.eml` folder tree, decoding IMAP Modified UTF-7 names and normalizing status headers | No (Stdlib only) |
| [`eml_to_pst_import.py`](file:///c:/Users/andre/Desktop/all2pst/eml_to_pst_import.py) | Import a directory tree of `.eml` files into one or more Outlook `.pst` files (split by year, even-split into `N` parts, or auto-rotated at `--max-pst-gb`) | Yes (`pywin32`) |
| [`imapmail_to_pst.py`](file:///c:/Users/andre/Desktop/all2pst/imapmail_to_pst.py) | Directly convert a Thunderbird `ImapMail` / `Mail` directory (or single `.mbox` file) into an Outlook `.pst` archive | Yes (`pywin32`) |

---

## Key Capabilities & Hardening

- **Full Folder Hierarchy Preservation**
  - Traverses Thunderbird `.sbd` subfolder containers, account folders (`ImapMail`, `Mail/Local Folders`), and single `.mbox` files.
  - Decodes RFC 3501 **IMAP Modified UTF-7** folder names (e.g., `Notes &- Invoices` → `Notes & Invoices`, `Caf&AOk-` → `Café`).
  - Mirrors relative folder structures inside target `.pst` archives (or imports flat when `--flatten` is specified).
- **Accurate Thunderbird `X-Mozilla-Status` Bitmask Mapping**
  - Maps `0x0001` (`Read`), `0x0002` (`Replied`), `0x0004` (`Marked` / Flagged), `0x0008` (`Expunged` / Deleted), and `0x1000` (`Forwarded`).
  - Automatically skips uncompacted deleted/expunged (`0x0008`) messages by default (override with `--include-deleted`).
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

### 1) Export Single MBOX → EML (`mbox_to_eml_exporter.py`)

```cmd
python mbox_to_eml_exporter.py ^
  --mbox "D:\Mail\inbox.mbox" ^
  --out-dir "D:\Export_EML" ^
  --layout month ^
  --sanitize-filenames
```

**Options:**
- `--layout {year,month,flat}`: Organize exported `.eml` files into `YYYY`, `YYYY\MM`, or flat directories (default: `year`).
- `--start-year YYYY` / `--end-year YYYY`: Filter messages by `Date` header year (inclusive).
- `--max-per-dir N`: Roll over to `__part2`, `__part3`, etc. when a directory reaches `N` files (`0` = unlimited).
- `--max-dir-bytes GB`: Roll over when a directory reaches `GB` gigabytes (`0` = unlimited).
- `--include-deleted`: Export messages marked as expunged (`0x0008`) in `X-Mozilla-Status`.
- `--progress-every N`: Print progress every `N` messages (default: `1000`).

---

### 2) Export Thunderbird ImapMail / Mail → EML (`imapmail_to_eml.py`)

Point `imap_src` at an account folder (`...\ImapMail\imap.example.com`), the parent `ImapMail` or `Mail` folder, or a single `.mbox` file:

```cmd
python imapmail_to_eml.py ^
  "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" ^
  "D:\Export_EML"
```

**Options:**
- `--include-deleted`: Export uncompacted messages marked as expunged (`0x0008`) in `X-Mozilla-Status`.

---

### 3) Import EML Tree → PST (`eml_to_pst_import.py`)

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
- `--split-by year`: Create one PST per message year (`<base>_<YYYY>_part1.pst`).
- `--splits N`: Divide total `.eml` byte volume across `N` sequential PST files (`<base>_part1.pst`, …).
- `--max-pst-gb GB`: Maximum size in GB per PST before rotating to the next part (default: `15.0`).
- `--pst-root NAME`: Top-level folder inside each PST (default: `"Imported (EML)"`; pass `""` to place folders directly at the PST root).
- `--flatten`: Place all messages directly inside `--pst-root` without recreating source subfolders.
- `--overwrite-pst`: Delete and recreate target PST files if they already exist on disk.
- `--include-deleted`: Import messages marked expunged (`0x0008`) in `X-Mozilla-Status`.
- `--flush-every N`: Detach and re-attach the current PST every `N` items so Windows Explorer updates file size (`0` = off).
- `--count-every N`: Print live folder item count every `N` items (default: `200`, `0` = off).

---

### 4) Direct Thunderbird ImapMail / Mail → PST (`imapmail_to_pst.py`)

```cmd
python imapmail_to_pst.py ^
  "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" ^
  "D:\PSTs\Mailbox.pst"
```

**Options:**
- `--include-deleted`: Import uncompacted messages marked as expunged (`0x0008`) in `X-Mozilla-Status`.

---

## Troubleshooting

- ** PST file size appears stuck at ~265 KB in Windows Explorer while running**
  - Outlook holds an open write handle on mounted `.pst` files and flushes metadata on unmount. The scripts automatically detach all PST stores in their `finally` blocks (and during `--flush-every`), at which point Explorer shows the final size.
- **Classic Outlook vs. "New Outlook"**
  - The web-based "New Outlook" for Windows does not support MAPI/COM automation (`Outlook.Application`). Switch to classic desktop Outlook before running `eml_to_pst_import.py` or `imapmail_to_pst.py`.
