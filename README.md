# `mail2mail` — Universal Email & Mailbox Converter (`mail2eml` / `mail2pst`)

Convert between **MBOX**, **Thunderbird `ImapMail` / `Mail`** (mbox and Maildir modes), **Maildir / Maildir++**, **Apple Mail (`.emlx` / `.mbox` bundles)**, **Outlook `.msg`**, **Outlook for Mac `.olm`**, **MH / Sylpheed / Claws Mail**, **MMDF**, **Emacs Babyl**, **`.eml` trees**, and **compressed archives (`.gz`, `.bz2`, `.xz`, `.zip`, `.tar.gz`, `.tgz`)** into `.eml` folder hierarchies (`mail2eml`) or directly into Microsoft Outlook **.pst** archives (`mail2pst`) with full folder hierarchy preservation, message flag mapping, size/year splitting, and resilient Outlook MAPI integration.

---

## Commands & Architecture

| Command / Module | Purpose | Requires Outlook? |
| :--- | :--- | :--- |
| [`mail2mail.py`](file:///c:/Users/andre/Desktop/all2pst/mail2mail.py) (`mail2mail`) | Unified master CLI (`mail2mail eml ...`, `mail2mail pst ...`, or auto-detected from `.pst` destination) | Only for `pst` |
| [`mail2eml.py`](file:///c:/Users/andre/Desktop/all2pst/mail2eml.py) (`mail2eml`) | Convert any supported mail format or archive into `.eml` files (`folder`, `year`, `month`, or `flat` layout) | No (Stdlib only) |
| [`mail2pst.py`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py) (`mail2pst`) | Convert/import any supported mail format or archive into one or more Outlook `.pst` files (single PST, year-split, `N`-way split, or auto-rotated at `--max-pst-gb`) | Yes (`pywin32`) |
| [`mail_formats.py`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py) | Shared multi-format detection, streaming, flag normalization, and folder hierarchy engine | No (Stdlib only) |

> [!NOTE]
> The original script filenames ([`mbox_to_eml_exporter.py`](file:///c:/Users/andre/Desktop/all2pst/mbox_to_eml_exporter.py), [`imapmail_to_eml.py`](file:///c:/Users/andre/Desktop/all2pst/imapmail_to_eml.py), [`eml_to_pst_import.py`](file:///c:/Users/andre/Desktop/all2pst/eml_to_pst_import.py), and [`imapmail_to_pst.py`](file:///c:/Users/andre/Desktop/all2pst/imapmail_to_pst.py)) are preserved as backward-compatible wrappers around [`mail2eml.py`](file:///c:/Users/andre/Desktop/all2pst/mail2eml.py) and [`mail2pst.py`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py).

---

## Supported Email & Mailbox Formats

All commands share [`mail_formats.py`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py) and automatically detect any of the following formats:

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

## Installation & Requirements

- **Python 3.9+**
- **Windows** with classic **Microsoft Outlook** desktop (M365 / 2021 / 2019 / 2016) for `.pst` output (`mail2pst`)

```bash
pip install -r requirements.txt
# Or install as CLI commands (mail2mail, mail2eml, mail2pst):
pip install -e .
```

### Outlook Preparation Tips (for `mail2pst`)
1. Ensure classic desktop Outlook has a default mail profile configured that opens without modal prompts.
2. Put Outlook in **Work Offline** mode while running large imports.
3. Avoid browsing or moving items inside the target `.pst` in Outlook while the import is active.

---

## Usage Guide

### 1) Unified `mail2mail` CLI

`mail2mail.py` dispatches to `mail2eml` or `mail2pst` either via subcommands (`eml` / `pst`) or automatically when the target ends with `.pst`:

```cmd
:: Convert any source -> Outlook PST
python mail2mail.py pst "D:\Mail\ImapMail" "D:\PSTs\Mailbox.pst"
python mail2mail.py "D:\Mail\export.olm" "D:\PSTs\MacMail.pst"

:: Convert any source -> EML hierarchy
python mail2mail.py eml "D:\Mail\inbox.mbox.gz" "D:\Export_EML" --layout month
```

---

### 2) Convert Any Format → EML (`mail2eml.py`)

#### Preserve 1:1 Folder Hierarchy (Thunderbird, Maildir, Apple Mail, `.olm`, `.zip`)
```cmd
python mail2eml.py ^
  "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" ^
  "D:\Export_EML" ^
  --layout folder ^
  --naming index
```

#### Organize by Year or Month with Date & Size Filters
```cmd
python mail2eml.py ^
  --src "D:\Mail\inbox.mbox.gz" ^
  --out-dir "D:\Export_EML" ^
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
- `--start-year YYYY` / `--end-year YYYY`: Filter messages by `Date` header year (inclusive).
- `--max-per-dir N`: Roll over to `__part2`, `__part3`, etc. when a directory reaches `N` files (`0` = unlimited).
- `--max-dir-bytes GB`: Roll over when a directory reaches `GB` gigabytes (`0` = unlimited).
- `--include-deleted`: Export messages marked as deleted/expunged.
- `--progress-every N`: Print progress every `N` messages (default: `1000`).

---

### 3) Convert Any Format → Outlook PST (`mail2pst.py`)

#### Mode A: Direct Single-PST Output
```cmd
python mail2pst.py ^
  "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" ^
  "D:\PSTs\Mailbox.pst"
```

#### Mode B: Split by Year (One PST per Year, Auto-Rotating at 15 GB)
```cmd
python mail2pst.py ^
  --src "D:\Export_EML" ^
  --out-dir "D:\PSTs" ^
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
  --src "D:\Export_EML" ^
  --out-dir "D:\PSTs" ^
  --base-name emails ^
  --splits 6 ^
  --max-pst-gb 18 ^
  --pst-root "Imported (EML)"
```

**Options (`mail2pst.py`):**
- `src` / `--src` / `--mbox`: Source folder, mailbox file, or archive (`.eml`, `.emlx`, `.msg`, `.olm`, Maildir, MH, MBOX, MMDF, Babyl, `.zip`, `.tgz`, `.gz`).
- `out` / `--out` / `--pst`: Output `.pst` file path (for single-PST mode).
- `--out-dir DIR`: Output directory for multi-PST split/rotated mode.
- `--base-name NAME`: Base filename prefix when using `--out-dir` (default: `emails`).
- `--split-by year`: Create one PST per message year (`<base>_<YYYY>_part1.pst`).
- `--splits N`: Divide total byte volume across `N` sequential PST files (`<base>_part1.pst`, …).
- `--max-pst-gb GB`: Maximum size in GB per PST before rotating to the next part (default: `15.0`).
- `--pst-root NAME`: Top-level folder inside each PST (default: root `""` in single-PST mode, `"Imported (EML)"` in `--out-dir` mode).
- `--flatten`: Place all messages directly inside `--pst-root` without recreating source subfolders.
- `--start-year YYYY` / `--end-year YYYY`: Filter messages by `Date` header year (inclusive).
- `--overwrite-pst`: Delete and recreate target PST files if they already exist on disk.
- `--include-deleted`: Import messages marked as deleted/expunged.
- `--flush-every N`: Detach and re-attach the current PST every `N` items so Windows Explorer updates file size (`0` = off).
- `--count-every N`: Print live folder item count every `N` items (default: `200`, `0` = off).

---

## Troubleshooting

- **PST file size appears stuck at ~265 KB in Windows Explorer while running**
  - Outlook holds an open write handle on mounted `.pst` files and flushes metadata on unmount. `mail2pst` automatically detaches all PST stores on completion (and during `--flush-every`), at which point Explorer shows the final size.
- **Classic Outlook vs. "New Outlook"**
  - The web-based "New Outlook" for Windows does not support MAPI/COM automation (`Outlook.Application`). Switch to classic desktop Outlook before running `mail2pst`.
