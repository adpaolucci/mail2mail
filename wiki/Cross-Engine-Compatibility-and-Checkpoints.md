# Cross-Engine Compatibility & Checkpoint State Architecture

A core architectural strength of the project is the decoupled, interoperable relationship between **`mail2mail`** (Python) and **`mail2rust`** (Rust). 

Without sharing a single line of compile-time code, both engines adhere to a unified CLI contract, shared SQLite WAL / JSON checkpoint formats, and deterministic message hashing. This allows either engine to pause, resume, hand off, or audit migrations initiated by the other.

---

## 1. Shared CLI Specification

Both engines implement standardized flags and exit behaviors:

### Standardized Exit Codes
* **`0`**: Migration completed successfully with zero unhandled errors.
* **`1`**: Fatal error (invalid input path, unrecoverable I/O failure, permission denied).
* **`130`**: Gracefully halted by user interrupt (`SIGINT` / Ctrl+C, `SIGTERM`, or interactive `q` keypress), saving current progress to the checkpoint file before terminating.

### Shared Filter & Execution Flags
| Flag | Semantic Expected Behavior |
| :--- | :--- |
| `--since <date>` / `--after <date>` | Inclusive lower timestamp bound. Parses ISO 8601 (`YYYY-MM-DD`, `YYYY-MM-DDTHH:MM:SS`), RFC 2822, or 4-digit year. Evaluated against UTC. |
| `--until <date>` / `--before <date>` | Inclusive upper timestamp bound. Date-only strings evaluate to `23:59:59.999999` UTC of that day. |
| `--include-folder <glob>` | Case-insensitive glob applied to folder hierarchy (repeatable). |
| `--exclude-folder <glob>` | Case-insensitive glob applied to folder hierarchy (repeatable). |
| `--split-size <size>` | Dynamic capacity boundary (e.g. `20GB`, `500MB`, `15G`, `4g`, `1024K`, raw bytes). |
| `--state-file <path>` | Path to shared SQLite (`.sqlite`, `.db`) or JSON Lines (`.jsonl`) checkpoint cache. |
| `--resume <path>` | Explicit checkpoint file to resume an interrupted transfer from. |
| `--dedup <mode>` | Deduplication mode: `none`, `auto`, `message-id`, `sha256`. |

---

## 2. Checkpoint Serialization Contracts

### 2.1 JSON Checkpoint Format (`.json`)
Used for lightweight, human-readable state hand-off between Python and Rust:

```json
{
  "id": "chk_1712345678901",
  "timestamp": "2026-11-07T16:00:00Z",
  "source_id": "source_mailbox.mbox",
  "target_id": "destination_archive.pst",
  "migrated_message_ids": [
    "<msg001@example.org>",
    "<msg002@example.org>",
    "sha256_8f4a1b2c3d4e5f6a..."
  ],
  "folder_high_watermarks": {
    "Inbox": 14205,
    "Sent Items": 4810
  },
  "total_bytes_migrated": 1048576000,
  "total_messages_migrated": 19015,
  "status": "in_progress"
}
```

#### Cross-Engine Handoff Workflows:
* **Python $\to$ Rust:**
  1. `mail2mail` starts an export: `python mail2mail.py export mailbox.mbox out.pst --state-file checkpoint.json`.
  2. The process is interrupted at 50,000 messages.
  3. `mail2rust` resumes the job: `mail2rust convert mailbox.mbox out.pst --resume checkpoint.json`.
  4. Rust skips the 50,000 processed IDs and streams the remaining messages without duplicates.
* **Rust $\to$ Python:**
  1. `mail2rust` runs with `--state-file checkpoint.json`.
  2. `mail2mail` resumes using `--resume checkpoint.json` seamlessly.

---

### 2.2 SQLite WAL High-Concurrency State Cache (`.sqlite`)
For multi-million-item migrations, in-memory arrays and JSON files can exhaust system memory. In these scenarios, both tools support SQLite with Write-Ahead Logging:

```sql
PRAGMA journal_mode=WAL;
PRAGMA synchronous=NORMAL;

CREATE TABLE IF NOT EXISTS completed_records (
    key TEXT PRIMARY KEY,
    message_id TEXT,
    sha256 TEXT,
    source_file TEXT,
    folder TEXT,
    timestamp TEXT
);

CREATE INDEX IF NOT EXISTS idx_mid ON completed_records(message_id);
CREATE INDEX IF NOT EXISTS idx_sha ON completed_records(sha256);
```

#### Deduplication Resolution Logic:
$$\text{LookupKey} = \begin{cases} 
\text{NormalizedMessageID} & \text{if Header present and valid} \\
\text{"sha256:"} + \text{SHA256}(\text{RawMimeBytes}) & \text{if Message-ID is missing or malformed}
\end{cases}$$

---

## 3. Black-Box Test Harness Parity

To prevent regressions, the repositories share identical synthetic datasets covering edge cases:
* **Unicode & Modified UTF-7 Folders:** `&J5Yw-` (Japanese), `&AOQ-` (ä), nested sub-folders (`Inbox/Work/Projects/2026`).
* **Corrupted Boundaries:** Missing `From ` separators, malformed dates (`Date: -0000`), folded headers with mixed tabs and spaces.
* **Multipart Edge Cases:** Mixed alternative HTML/Plain, inline CID attachments, nested `.eml` attachments (`message/rfc822`), non-ASCII attachment filenames (RFC 2231 / RFC 2047).
