# `mail2mail` & `mail2rust` Cross-Compatibility Specification

This specification defines the shared interoperability contract between **`mail2mail`** (Python) and **`mail2rust`** (Rust). By adhering to this contract, both applications can independently test each other, hand off incremental delta migration states, and maintain 1:1 format parity without sharing a single line of codebase code.

---

## 1. Core Principles

1. **Strict Decoupling**: Neither project depends on the other as a submodule, library, or compile-time dependency.
2. **Black-Box Testing**: Testing is conducted via standard CLI invocation, standard email file formats (`.eml`, `.mbox`), and shared audit manifests.
3. **Format Parity**: Both engines parse and emit RFC 5322 MIME, RFC 4155 / `mboxrd`, and RFC 3501 Modified UTF-7 with identical semantic meaning and flag mapping.
4. **State Interoperability**: Both tools read and write the exact same SQLite Write-Ahead Logging (WAL) state database and CSV audit manifest format.

---

## 2. CLI Contract & Flags

Both `mail2mail` and `mail2rust` implement the following standardized command-line interface:

### 2.1 Subcommands
* `eml`: Convert mail sources into `.eml` directory trees.
* `mbox`: Convert mail sources into `.mbox` (or compressed `.mbox.gz`) files.
* `pst`: Convert mail sources into Microsoft Outlook `.pst` files (Windows MAPI).

### 2.2 Standard Exit Codes
* **`0`**: Migration completed successfully.
* **`1`**: Fatal error (e.g., source not found, unrecoverable I/O error).
* **`130`**: Gracefully halted by user request (`SIGINT` / Ctrl+C, `SIGTERM`, or interactive `q` stop).

### 2.3 Shared Filter & Capacity Flags
| Flag | Description | Expected Behavior |
| :--- | :--- | :--- |
| `--since <date>` / `--after <date>` | Lower date bound (inclusive) | Parse ISO 8601 (`YYYY-MM-DD`, `YYYY-MM-DDTHH:MM:SS`), RFC 2822, or year `YYYY`. UTC comparison. |
| `--until <date>` / `--before <date>` | Upper date bound (inclusive) | Same parsing; if date-only, defaults to end of day (23:59:59.999999 UTC). |
| `--include-folder <glob>` | Folder inclusion rule | Case-insensitive glob matching full path or segment. Repeatable. |
| `--exclude-folder <glob>` | Folder exclusion rule | Case-insensitive glob matching full path or segment. Repeatable. |
| `--split-size <size>` | Dynamic capacity splitting | Accepts human strings (`20GB`, `500MB`, `15G`, `4g`, `1024K`, raw bytes). |
| `--state-file <path>` | Incremental delta state cache | Path to shared SQLite (`.sqlite`, `.db`) or JSON Lines (`.jsonl`) database. |
| `--audit-log <path>` | Forensic manifest | Path to output `.csv` or `.jsonl` audit log. |
| `--dedup <mode>` | Deduplication mode | `none`, `auto`, `message-id`, `sha256`. |

---

## 3. Incremental Delta State & Checkpoint Contract (`--state-file`, `--resume`)

Both applications support seamless, bidirectional state hand-off across runs. Two storage formats are supported:

### 3.1 JSON Checkpoint Format (`.json`)
For direct bidirectional delta synchronization between Python and Rust engines, checkpoints are serialized as standard JSON:

```json
{
  "id": "chk_1712345678901",
  "timestamp": "2026-11-07T16:00:00Z",
  "source_id": "source_mailbox",
  "target_id": "destination_target",
  "migrated_message_ids": [
    "<msg1@domain.test>",
    "<msg2@domain.test>",
    "sha256_hex_digest_if_no_message_id"
  ],
  "folder_high_watermarks": {
    "Inbox": 1420
  },
  "total_bytes_migrated": 10485760,
  "total_messages_migrated": 2,
  "status": "completed"
}
```

* **Pass 1 in Python $\to$ Pass 2 in Rust**:
  - Python writes `state_file="checkpoint.json"`.
  - Rust resumes via `mail2rust eml <source> <dest> --resume checkpoint.json`.
  - Rust skips all messages in `migrated_message_ids` and migrates only delta messages.
* **Pass 1 in Rust $\to$ Pass 2 in Python**:
  - Rust writes `mail2rust eml <source> <dest> --state-file checkpoint.json`.
  - Python resumes via `mail2eml.export_to_eml(..., resume="checkpoint.json")`.
  - Python loads the JSON checkpoint and skips all previously migrated items.

### 3.2 SQLite Delta State Cache Contract (`.sqlite`, `.db`)

When `--state-file <path.sqlite>` is provided in Python or SQLite-enabled environments:

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

### Key Generation Rules:
1. If `Message-ID` header exists and is non-empty:
   - Primary key: `mid:<Message-ID>`
2. If `Message-ID` is missing or empty:
   - Primary key: `sha256:<SHA-256 hex digest of raw message bytes>`
3. Lookup check during migration:
   - A message is skipped if `mid:<Message-ID>` OR `sha256:<hash>` exists in `completed_records`.

---

## 4. Structured Audit Manifest Contract (`--audit-log`)

When `--audit-log <manifest.csv>` is specified, both tools write RFC 4180 compliant CSV files with the exact header row:

```csv
timestamp,status,source_file,source_folder,target_location,message_id,subject,date,sender,recipients,size_bytes,attachment_count,sha256,notes
```

### Valid Status Values:
* `SUCCESS`: Message was successfully written/imported.
* `SKIPPED_DUPLICATE`: Message was identified as a duplicate via `--dedup`.
* `SKIPPED_DATE`: Message fell outside `--since` / `--until` date range.
* `SKIPPED_DELETED`: Message was marked as deleted/expunged and `--include-deleted` was not set.
* `SKIPPED_FOLDER`: Message was filtered by `--include-folder` or `--exclude-folder`.
* `SKIPPED_STATE_CACHE`: Message was already processed in a prior pass (`--state-file`).
* `FAILED`: Message extraction or write failed (details in `notes`).

---

## 5. Automated Cross-Testing Execution

In `mail2mail`'s test suite, [`tests/test_cross_compatibility.py`](tests/test_cross_compatibility.py) automates:
1. **Differential Output Test**: Runs both tools against identical synthetic mailboxes and asserts identical output counts, folder structures, and SHA-256 hashes.
2. **Round-Trip Test**: `mail2mail` (MBOX $\to$ EML) $\to$ `mail2rust` (EML $\to$ MBOX) $\to$ `mail2mail` (MBOX $\to$ EML). Asserts byte and header fidelity.
3. **State Cache Hand-Off Test**: Runs Pass 1 in Python, adds new emails, runs Pass 2 in Rust, and asserts Rust exports only the new emails. Runs the reverse sequence as well.

To run the cross-compatibility suite, point `MAIL2RUST_BIN` to the compiled Rust binary:
```bash
# In mail2mail repository:
export MAIL2RUST_BIN="../mail2rust/target/release/mail2rust"
python -m unittest tests.test_cross_compatibility -v
```
If `MAIL2RUST_BIN` is not set and no binary is found in sibling directories, the suite skips gracefully without failing standard CI runs.
