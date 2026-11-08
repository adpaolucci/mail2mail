# Mail2Rust Multi-Format Sample & Evaluation Suite

This directory contains verified, production-compliant sample data for all major email formats supported by **Mail2Rust**.

## Included Sample Datasets

| Directory | Format | Description | Sniffer Identification | CLI Test Command |
| :--- | :--- | :--- | :--- | :--- |
| `pst/sample_unicode.pst` | **PST** | 64-bit Outlook Unicode PST with `mpbbR` compressible encryption, folders, & attachments | `Outlook Personal Folders (PST)` | `mail2rust validate-pst pst/sample_unicode.pst` |
| `pst/sample_encrypted_password.pst` | **PST** | High-Encryption (`NDB_CRYPT_CYCLIC`) PST with password (`SecretPass123!`) | `Outlook Personal Folders (PST)` | `mail2rust validate-pst pst/sample_encrypted_password.pst --password SecretPass123!` |
| `mbox/sample_mailbox.mbox` | **MBOX** | Standard RFC 4155 mboxrd with `>From` escaping | `Standard MBOX Archive` | `mail2rust sniff mbox/sample_mailbox.mbox` |
| `mbox/sample_mailbox.mbox.gz` | **GZ/MBOX** | Gzip compressed MBOX stream (`0x1F8B` magic) | `Gzip Compressed Stream (GZ)` | `mail2rust pst mbox/sample_mailbox.mbox.gz out.pst` |
| `eml/standalone_message.eml` | **EML** | RFC 5322 MIME message with text/html and attachments | `RFC 5322 EML Message` | `mail2rust pst eml/standalone_message.eml out.pst` |
| `eml/thread/` | **EML** | Multi-message conversation thread linked by `Message-ID` & `In-Reply-To` | `RFC 5322 EML Message` | `mail2rust pst eml/thread out.pst` |
| `emlx/sample_apple_mail.emlx` | **EMLX** | Apple Mail message with byte-length header & `<plist>` metadata flags | `Apple Mail EMLX` | `mail2rust sniff emlx/sample_apple_mail.emlx` |
| `maildir/` | **Maildir** | Full `cur/`, `new/`, `tmp/` hierarchy with status flag suffixes (`:2,S`, `:2,2FS`) | `Maildir Mailbox Tree` | `mail2rust pst maildir out.pst` |
| `contacts/sample_contacts.vcf` | **vCard** | RFC 6350 / RFC 2426 electronic business cards (ingests to Outlook `Contacts`) | `vCard Address Book (VCF)` | `mail2rust pst contacts/sample_contacts.vcf out.pst` |
| `calendar/sample_calendar.ics` | **iCalendar** | RFC 5545 appointments & meetings (ingests to Outlook `Calendar`) | `iCalendar Events & Meetings (ICS)` | `mail2rust pst calendar/sample_calendar.ics out.pst` |
| `msg/sample_outlook.msg` | **MSG** | Microsoft Outlook Compound File Binary (CFB) item (`\xD0\xCF\x11\xE0...`) | `Outlook Item (MSG)` | `mail2rust sniff msg/sample_outlook.msg` |
| `olm/sample_mac_archive.olm` | **OLM** | Outlook for Mac compressed ZIP archive | `Outlook for Mac Archive (OLM)` | `mail2rust sniff olm/sample_mac_archive.olm` |
| `takeout/sample_google_takeout.zip` | **Takeout** | Google Takeout archive (`Takeout/Mail/All mail...`) | `Compressed Archive (ZIP / Takeout)` | `mail2rust pst takeout/sample_google_takeout.zip out.pst` |
| `mh/` | **MH** | MH / Sylpheed numbered mail files with `.mh_sequences` | `MH / Sylpheed / Claws Mail` | `mail2rust pst mh out.pst` |
| `mmdf/sample_mmdf.txt` | **MMDF** | Multichannel Memorandum Distribution Facility (`\x01\x01\x01\x01\n` delimiters) | `MMDF Mailbox` | `mail2rust sniff mmdf/sample_mmdf.txt` |
| `babyl/sample_babyl.txt` | **Babyl** | Emacs RMAIL Babyl format (`BABYL OPTIONS:` and `\x1F\x0C` delimiters) | `Emacs Babyl` | `mail2rust sniff babyl/sample_babyl.txt` |
| `forensics/sample_manifest.json` | **JSON** | Forensic eDiscovery manifest with RFC 6234 Merkle root & Bates range | `JSON Forensic Manifest` | `cat forensics/sample_manifest.json` |
| `forensics/sample_manifest.csv` | **CSV** | Chain-of-custody CSV ledger with dual SHA-256 and MD5 hashes | `CSV Forensic Manifest` | `cat forensics/sample_manifest.csv` |
| `sync/sample_checkpoint.json` | **JSON** | ACID sync checkpoint tracking migrated message IDs and folder watermarks | `Sync Checkpoint` | `cat sync/sample_checkpoint.json` |
| `jmap/sample_jmap_export.json` | **JMAP** | RFC 8620 / RFC 8621 JMAP Email batch entities with keywords and thread tracking | `JMAP RFC 8620/8621 JSON` | `python mail2mail.py eml samples/jmap/sample_jmap_export.json out_eml` |

---
