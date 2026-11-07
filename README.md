# MBOX / Thunderbird ImapMail → EML → PST Toolkit

Export large `.mbox` mailboxes or Thunderbird `ImapMail` account hierarchies into filesystem `.eml` files and import them into Outlook **.pst** archives — with folder hierarchy preservation, Thunderbird read/flagged/replied status mapping, year/month layouts, filters, size limits, progress logs, and safe Outlook MAPI integration.

## Scripts Included

- **`mbox_to_eml_exporter.py`**: Single `.mbox` → `.eml` folders (`year`, `month`, or `flat` layout)
- **`imapmail_to_eml.py`**: Thunderbird `ImapMail` / `Mail` account tree (`.sbd` subfolders & mbox files) → `.eml` folder tree
- **`eml_to_pst_import.py`**: `.eml` folder tree → Outlook `.pst` archive(s) (with subfolder preservation, year/byte splitting, and size rotation)
- **`imapmail_to_pst.py`**: Thunderbird `ImapMail` / `Mail` account tree → single Outlook `.pst` archive directly

> Designed for multi-tens-of-GB mailboxes. Avoids the “Drafts” pitfall by creating items **directly** in the target PST folder (`folder.Items.Add(0)`).

---

## Features

**MBOX / ImapMail → EML**
- Single `.mbox` export (`mbox_to_eml_exporter.py`) with **year**, **month** (`YYYY/MM`), or **flat** layouts
- Full Thunderbird `ImapMail` hierarchy export (`imapmail_to_eml.py`) with `.sbd` subfolder traversal and **IMAP Modified UTF-7** folder name decoding
- Accurately parses Thunderbird `X-Mozilla-Status` bitmasks (`0x0001` Read, `0x0002` Replied, `0x0004` Flagged/Starred, `0x0008` Expunged/Deleted, `0x1000` Forwarded)
- Automatically skips expunged/deleted messages (`0x0008`) unless `--include-deleted` is specified
- Year range filters (`--start-year`, `--end-year`) and per-directory caps (**max files** or **max GB**)
- Safe RFC 2047 header decoding and Windows/macOS filename sanitization

**EML / ImapMail → PST (Outlook MAPI)**
- Mirrors source subfolder hierarchy inside the PST (or flat mode via `--flatten`)
- Preserves **Read/Unread**, **Follow-up Flag**, **Replied/Forwarded icons**, **UTC Delivery/Submit timestamps**, **Unicode Sender/Recipient properties**, **Internet Message-ID**, **Transport Headers**, and **Inline CID images**
- Split by **year** (one PST per year) or even-split into **N PSTs by bytes**
- **Max PST size** (e.g., 15–20 GB) with auto-rotate to `…_part2.pst`, `…_part3.pst`
- Safe periodic **flush** (detach/reattach without deleting data) so Windows Explorer refreshes file growth

---

## Requirements

- **Windows** with classic **Microsoft Outlook** desktop (M365 / 2021 / 2019 / 2016) for `.pst` creation
- **Python 3.9+**
- Python package: `pywin32` (required for PST scripts; EML exporters use only the Python standard library)

```bash
pip install -r requirements.txt
```

---

## Outlook Tips
- Put Outlook in **Work Offline** during imports.
- Ensure a default Outlook profile opens without prompts.
- Don’t browse/move items inside the target PST while the script runs.
- If Windows Explorer shows PST ~256 KB while attached, that’s normal; the script detaches PSTs at the end (and on `--flush-every`) so the OS updates file size.

---

## 1) Export MBOX → EML (`mbox_to_eml_exporter.py`)

```cmd
python mbox_to_eml_exporter.py ^
  --mbox "D:\Mail\inbox.mbox" ^
  --out-dir "D:\Export_EML" ^
  --layout month ^
  --sanitize-filenames
```

### Common layouts
- `--layout year` → `Export_EML\YYYY\*.eml`
- `--layout month` → `Export_EML\YYYY\MM\*.eml`
- `--layout flat` → `Export_EML\*.eml`

### Year filters & folder limits
```cmd
:: Only 2007–2016
python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout year --start-year 2007 --end-year 2016

:: Cap folders at 50k files OR ~9 GB
python mbox_to_eml_exporter.py --mbox "D:\Mail\inbox.mbox" --out-dir "D:\Export_EML" --layout month --max-per-dir 50000 --max-dir-bytes 9
```

---

## 2) Export Thunderbird ImapMail → EML (`imapmail_to_eml.py`)

```cmd
python imapmail_to_eml.py "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" "D:\Export_EML"
```

Pass `--include-deleted` if you want to export uncompacted messages marked as deleted (`0x0008`) in Thunderbird.

---

## 3) Import EML → PST (`eml_to_pst_import.py`)

### Split by year, cap 15 GB per PST
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

### Even-split into N PSTs by total size (no per-year split)
```cmd
python eml_to_pst_import.py ^
  --src "D:\Export_EML" ^
  --out-dir "D:\PSTs" ^
  --base-name emails ^
  --splits 6 ^
  --max-pst-gb 18 ^
  --pst-root "Imported (EML)"
```

### Useful flags
- `--split-by year` or `--splits N`
- `--max-pst-gb 15` (rotate to `part2` when exceeded)
- `--pst-root "Imported (EML)"` (top folder inside each PST; pass `""` to place folders directly at PST root)
- `--flatten` (place all messages directly in `--pst-root` without recreating relative subfolders)
- `--overwrite-pst` (delete and recreate target PST if it already exists on disk)
- `--include-deleted` (import messages marked expunged in `X-Mozilla-Status`)
- `--flush-every N` (detaches and re-attaches the active PST every `N` items so Explorer refreshes file size)
- `--count-every N` (prints Outlook folder item counts)

---

## 4) Direct Thunderbird ImapMail → PST (`imapmail_to_pst.py`)

```cmd
python imapmail_to_pst.py "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" "D:\PSTs\Mailbox.pst"
```
