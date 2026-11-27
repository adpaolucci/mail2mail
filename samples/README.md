# mail2mail Multi-Format Sample & Evaluation Suite

This directory contains small sample files for the formats **mail2mail** supports. Most are
genuine, standards-conformant fixtures (RFC 5322 EML, RFC 4155 mboxrd, Maildir, vCard, iCalendar,
MMDF, Babyl, EMLX, Google Takeout). The PST fixtures are **synthetic** and are described honestly
below.

> **CLI note.** The commands in the table are from the sibling `mail2rust` project and are shown for
> format-identification reference. The equivalent commands in this repository are `mail2pst`,
> `mail2eml` and `mail2mbox` (for example `mail2pst eml out.pst`).

## Included Sample Datasets

| Directory | Format | Description | Sniffer Identification | Reference CLI |
| :--- | :--- | :--- | :--- | :--- |
| `pst/sample_unicode.pst` | **PST** | **Synthetic fixture, not a valid Outlook PST.** See `pst/README.md`. | `Outlook Personal Folders (PST)` | `mail2pst <source> out.pst` |
| `pst/sample_standard.pst` | **PST** | **Synthetic fixture, not a valid Outlook PST.** See `pst/README.md`. | `Outlook Personal Folders (PST)` | `mail2pst <source> out.pst` |
| `pst/sample_encrypted_password.pst` | **PST** | **Synthetic fixture, not a valid Outlook PST.** See `pst/README.md`. | `Outlook Personal Folders (PST)` | `mail2pst <source> out.pst` |
| `mbox/sample_mailbox.mbox` | **MBOX** | Standard RFC 4155 mboxrd with `>From` escaping | `Standard MBOX Archive` | `mail2mbox <source> out.mbox` |
| `mbox/sample_mailbox.mbox.gz` | **GZ/MBOX** | Gzip compressed MBOX stream (`0x1F8B` magic) | `Gzip Compressed Stream (GZ)` | `mail2mbox <source> out.mbox` |
| `eml/standalone_message.eml` | **EML** | RFC 5322 MIME message with text/html and attachments | `RFC 5322 EML Message` | `mail2pst eml out.pst` |
| `eml/thread/` | **EML** | Multi-message conversation thread linked by `Message-ID` & `In-Reply-To` | `RFC 5322 EML Message` | `mail2pst eml/thread out.pst` |
| `emlx/sample_apple_mail.emlx` | **EMLX** | Apple Mail message with byte-length header & `<plist>` metadata flags | `Apple Mail EMLX` | `mail2eml <source> out/` |
| `maildir/` | **Maildir** | Full `cur/`, `new/`, `tmp/` hierarchy with status flag suffixes (`:2,S`, `:2,2FS`) | `Maildir Mailbox Tree` | `mail2pst maildir out.pst` |
| `contacts/sample_contacts.vcf` | **vCard** | RFC 6350 / RFC 2426 electronic business cards (ingests to Outlook `Contacts`) | `vCard Address Book (VCF)` | `mail2pst contacts/ out.pst` |
| `calendar/sample_calendar.ics` | **iCalendar** | RFC 5545 appointments & meetings (ingests to Outlook `Calendar`) | `iCalendar Events & Meetings (ICS)` | `mail2pst calendar/ out.pst` |
| `msg/sample_outlook.msg` | **MSG** | Microsoft Outlook Compound File Binary (CFB) item (`\xD0\xCF\x11\xE0...`) | `Outlook Item (MSG)` | `mail2eml msg/ out/` |
| `olm/sample_mac_archive.olm` | **OLM** | Outlook for Mac compressed ZIP archive | `Outlook for Mac Archive (OLM)` | `mail2eml olm/ out/` |
| `takeout/sample_google_takeout.zip` | **Takeout** | Google Takeout archive (`Takeout/Mail/All mail...`) | `Compressed Archive (ZIP / Takeout)` | `mail2pst takeout/ out.pst` |
| `mh/` | **MH** | MH / Sylpheed / Claws Mail numbered mail files with `.mh_sequences` | `MH / Sylpheed / Claws Mail` | `mail2pst mh/ out.pst` |
| `mmdf/sample_mmdf.txt` | **MMDF** | Multichannel Memorandum Distribution Facility (`\x01\x01\x01\x01\n` delimiters) | `MMDF Mailbox` | `mail2eml mmdf/ out/` |
| `babyl/sample_babyl.txt` | **Babyl** | Emacs RMAIL Babyl format (`BABYL OPTIONS:` and `\x1F\x0C` delimiters) | `Emacs Babyl` | `mail2eml babyl/ out/` |
| `forensics/sample_manifest.json` | **JSON** | Forensic eDiscovery manifest with RFC 6234 Merkle root & Bates range | `JSON Forensic Manifest` | `cat forensics/sample_manifest.json` |
| `forensics/sample_manifest.csv` | **CSV** | Chain-of-custody CSV ledger with dual SHA-256 and MD5 hashes | `CSV Forensic Manifest` | `cat forensics/sample_manifest.csv` |
| `sync/sample_checkpoint.json` | **JSON** | ACID sync checkpoint tracking migrated message IDs and folder watermarks | `Sync Checkpoint` | `cat sync/sample_checkpoint.json` |

## About the PST fixtures

`mail2mail` reads and writes PST files through Outlook's MAPI interface (`pywin32`), so it does not
need to parse the PST binary format itself. The three `.pst` files in `pst/` are small synthetic
containers produced by an early prototype of the sibling `mail2rust` writer. They carry the correct
`!BDN` signature and are identified as PST by header sniffing, but they are **not readable by
Microsoft Outlook, `scanpst.exe` or libpff**, and they must not be used to validate PST
compatibility.

To exercise the PST code path, create a real PST on a Windows machine with Outlook installed:

```powershell
mail2pst samples\eml out.pst
```

The resulting file is produced by Outlook itself and is therefore conformant by construction.

---

## Verifying the samples

```powershell
# Every non-PST sample can be round-tripped without Outlook:
mail2eml samples\mbox\sample_mailbox.mbox out\eml
mail2mbox samples\eml out\mailbox.mbox
```
