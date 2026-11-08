# Forensic Auditing & Chain of Custody

In eDiscovery, compliance archiving, and legal investigations, email migrations must satisfy rigorous **chain-of-custody** standards. Any silent alteration of headers, loss of attachment metadata, or message truncation can invalidate evidence in court.

`mail2rust` and `mail2mail` incorporate an enterprise forensic auditing framework that produces cryptographic audit logs and ensures tamper-evident migrations.

---

## 1. Cryptographic Hash Verification

Every message ingested and converted is tracked by:
1. **Source MIME Hash:** $\text{SHA-256}(\text{RawMessageBytes})$
2. **Normalized Header Hash:** $\text{SHA-256}(\text{CanonicalHeaders})$ (stripping transit-added headers like `X-Spam-Status`, `Received`, `Return-Path` to verify body integrity across different relay servers).
3. **Payload Attachment Hashes:** Individual SHA-256 checksums for every embedded attachment stream.

---

## 2. Audit Manifest Schema (`--audit-log`)

Running a migration with `--audit-log <path.csv>` generates a tamper-evident record of every item processed:

```bash
mail2rust convert source.mbox archive.pst \
  --audit-log migration_audit.csv \
  --dedup sha256
```

### Manifest CSV Column Definitions:
| Column | Type | Description |
| :--- | :--- | :--- |
| `timestamp` | ISO-8601 UTC | Exact time the item was processed |
| `source_file` | String | Originating mailbox file or server host |
| `source_folder` | String | Folder path in source store (e.g., `Inbox/Legal`) |
| `target_folder` | String | Folder path in destination store |
| `message_id` | String | RFC 5322 `Message-ID` header value |
| `sha256` | Hex String | 64-character SHA-256 checksum of raw MIME payload |
| `size_bytes` | Integer | Total byte size of the message and attachments |
| `subject` | String | Decoded message subject |
| `sender` | String | Standardized sender address (`user@domain.com`) |
| `received_date`| ISO-8601 UTC | Original server delivery timestamp |
| `flags` | String | Flags transferred (e.g., `\Seen`, `\Flagged`, `Replied`) |
| `status` | String | `MIGRATED`, `DUPLICATE_SKIPPED`, `FILTERED`, `FAILED` |

### Sample Audit Log Record:
```csv
timestamp,source_file,source_folder,target_folder,message_id,sha256,size_bytes,subject,sender,received_date,flags,status
2026-11-07T18:42:01Z,source.mbox,Inbox,Inbox,<contract-992@law.corp>,e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855,142850,"Confidential NDA Review",counsel@law.corp,2026-11-05T09:15:22Z,"\Seen",MIGRATED
```

---

## 3. MIME Normalization & Bit-Exact Reproducibility

When comparing messages between formats (e.g. MBOX $\leftrightarrow$ PST $\leftrightarrow$ EML), differences in line endings or header folding can cause naive hash comparisons to fail.

The forensic engine applies deterministic RFC 5322 normalization:
* **Line Ending Canonicalization:** Converts mixed `\r\n` and bare `\n` into standard network `\r\n` (CRLF) for MIME and `\n` for MBOX.
* **Mboxrd `>From ` Un-escaping:** In MBOX files, lines beginning with `From ` are escaped with a prepended chevron (`>From `). The engine cleanly un-escapes these lines during conversion so the resulting MIME bytes match the original transit hash.
* **Non-ASCII Header Decoding (RFC 2047):** Safely decodes `ISO-8859-1`, `Windows-1252`, and `UTF-8` encoded words (`=?utf-8?B?...?=`) into normalized Unicode strings without data loss.

---

## 4. Forensic Inspection Tools

Both toolchains offer dedicated CLI commands to inspect, verify, and validate mail stores without modifying them:

```bash
# Verify integrity of a PST file against an audit manifest
mail2rust forensics verify archive.pst --manifest migration_audit.csv

# Inspect MAPI property streams of an individual .msg file
mail2rust inspect message.msg

# Dump raw MAPI property tags from an Outlook .ost cache
python mail2mail.py inspect cache.ost
```
