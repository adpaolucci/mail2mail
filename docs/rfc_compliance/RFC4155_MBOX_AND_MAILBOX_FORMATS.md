# RFC 4155 (`mbox` / `mboxrd`), RFC 3501 (IMAP mUTF-7), & Mailbox Storage Formats

This document specifies how `mail2mail` reads and writes `mbox`/`mboxrd` streams, decodes/encodes IMAP4rev1 Modified UTF-7 folder names, and maps status flags across Thunderbird, Maildir/Maildir++, Apple Mail `.emlx`, MH, MMDF, and Emacs Babyl formats.

---

## 1. RFC 4155 (`application/mbox`) & `mboxrd` Body Quoting

### 1.1 Envelope `From ` Line Format (RFC 4155 Appendix A)
- **Standard Requirement**: Each message in an mbox database is preceded by a separator line:
  ```text
  From <sender-addr> <24-char-UTC-asctime>\n
  ```
  Example: `From alice@example.com Fri Nov  6 20:30:00 2020\n`
- **Reader ([`is_mbox_from_line`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L1528-L1534) & [`MBOX_FROM_RE`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L88-L94))**:
  - Matches canonical RFC 4155 24-character `asctime` lines (`From <addr> <Day> <Mon> <dd> <hh:mm:ss> [TZ] <yyyy>`) as well as Thunderbird's `From - <Day> <Mon> ...` placeholder sender format.
  - Handles back-to-back messages (including header-only messages with no body) without requiring a preceding blank line when a valid `From ` envelope line is encountered.
- **Writer ([`build_mbox_from_line`](file:///c:/Users/andre/Desktop/all2pst/mail2mbox.py#L64-L94))**:
  - Extracts the sender address from `Return-Path` or `From` (defaulting to `MAILER-DAEMON` if absent or whitespace-bearing) and formats the `Date` header in UTC using `time.asctime(dt.utctimetuple())`.

### 1.2 `mboxrd` Reversible `>From ` Escaping
- **Why `mboxrd` over `mboxo`**: In legacy `mboxo`, only bare `From ` lines in the body are prefixed with `>`, which cannot be unambiguously reversed if a line originally began with `>From `. In **`mboxrd`** (used by Thunderbird, qmail, and modern Unix tools), *any* body line matching `^>*From ` has one `>` prepended on write (`From ` $\to$ `>From `, `>From ` $\to$ `>>From `) and one `>` stripped on read.
- **Writer ([`format_mboxrd_entry`](file:///c:/Users/andre/Desktop/all2pst/mail2mbox.py#L97-L135))**:
  - Applies `re.sub(rb"^(>*From )", rb">\1", body_block, flags=re.MULTILINE)` exclusively to the message body (after the `\n\n` header separator) and appends a trailing blank line.
- **Reader ([`iter_mbox_stream`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L1561-L1620))**:
  - While in the body state (`in_headers == False`), any non-boundary line where `line.startswith(b">")` and `line.lstrip(b">").startswith(b"From ")` has its leading `>` stripped (`line = line[1:]`), restoring exact byte-for-byte body fidelity across round-trips.

---

## 2. RFC 3501 §5.1.3 — IMAP4rev1 Modified UTF-7 Mailbox Naming

- **Standard Specification**:
  1. Printable US-ASCII characters (`0x20..0x7E`), except `&` (`0x26`), represent themselves.
  2. Literal `&` is represented by the two-character sequence `&-`.
  3. All other characters (`0x00..0x1F`, `0x7F..0xFFFF`, and supplementary planes via UTF-16 surrogate pairs) are encoded in **UTF-16BE**, converted to **Modified Base64** (where `/` is replaced by `,` so `/` cannot conflict with filesystem path separators, and trailing `=` padding is omitted), and bracketed between `&` and `-`.
- **Examples**:
  - `Notes & Invoices` $\leftrightarrow$ `Notes &- Invoices`
  - `Café` $\leftrightarrow$ `Caf&AOk-`
  - `日本語/台北` $\leftrightarrow$ `&ZeVnLIqe-/&U,BTFw-`
- **Implementation**:
  - [`decode_imap_utf7`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L130-L160) & [`encode_imap_utf7`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L163-L193) in [`mail_formats.py`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py).

---

## 3. Message Status Flag Specifications Across Formats

`mail2mail` normalizes message status flags across all formats into a canonical 5-flag model (`is_read`, `is_replied`, `is_flagged`, `is_forwarded`, `is_deleted`):

| Flag | `X-Mozilla-Status` (Hex Bitmask) | BSD `Status` / UW-IMAP `X-Status` | Maildir `:2,` / `;2,` Info | Apple Mail `.emlx` `<plist>` `flags` | Outlook MAPI Property |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Read (Seen)** | `0x0001` (`MSG_FLAG_READ`) | `Status: R` (`RO` vs `O`) | `S` in `cur/` (absent in `new/`) | Bit 0 (`1 << 0 = 1`) | `MailItem.UnRead = False` |
| **Replied** | `0x0002` (`MSG_FLAG_REPLIED`) | `X-Status: A` | `R` | Bit 2 (`1 << 2 = 4`) | `PR_LAST_VERB_EXECUTED = 102`, `PR_ICON_INDEX = 261` |
| **Flagged (Marked)** | `0x0004` (`MSG_FLAG_MARKED`) | `X-Status: F` / `X-Message-Flag` | `F` | Bit 4 (`1 << 4 = 16`) | `MailItem.FlagStatus = 2` (`olFlagMarked`) |
| **Deleted (Expunged)** | `0x0008` (`MSG_FLAG_EXPUNGED`) | `Status: D` / `X-Status: D` | `T` (Trashed) | Bit 1 (`1 << 1 = 2`) | Skipped unless `--include-deleted` |
| **Forwarded** | `0x1000` (`MSG_FLAG_FORWARDED`) | `X-Status: W` | `P` (Passed) | Bit 8 (`1 << 8 = 256`) | `PR_LAST_VERB_EXECUTED = 104`, `PR_ICON_INDEX = 262` |

---

## 4. Maildir, Apple Mail `.emlx`, MH, MMDF, and Babyl Specifications

- **Maildir & Maildir++ ([`parse_maildir_flags`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L469-L511), [`normalize_maildir_folder_parts`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L514-L539))**:
  - Supports both POSIX `:2,<flags>` and Windows-compatible `;2,<flags>` (used by Thunderbird Maildir on Windows and offlineimap/mbsync).
  - Ignores `tmp/` directories and converts Maildir++ dot-prefixed folders (`.Work.Projects.2024`) into nested folder paths (`Work/Projects/2024`).
- **Apple Mail `.emlx` & `*.mbox` Bundles ([`parse_emlx_bytes`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L569-L619))**:
  - Reads the ASCII decimal byte-count line at the start of `.emlx` / `.partial.emlx` files, slices the exact RFC 5322 byte range, and parses the trailing XML `<plist>` `<key>flags</key><integer>...</integer>` bitmask.
  - Collapses internal `*.mbox/<UUID>/Data/0..9/Messages` shard directories into clean user-facing folder paths.
- **RAND MH ([`parse_mh_sequences`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L1470-L1496))**:
  - Parses `.mh_sequences` sequence ranges (`unseen: 1-5 8`, `flagged`, `replied`, `forwarded`, `deleted`) and maps numbered message files (`1`, `2`, …).
- **MMDF ([`iter_mmdf_stream`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L1622-L1653))**:
  - Splits messages delimited by 4-SOH lines (`\x01\x01\x01\x01\n`).
- **Emacs Rmail Babyl ([`iter_babyl_stream`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L1655-L1702))**:
  - Splits `\x1f`-delimited Babyl records, parses the attribute line (`1, unseen, answered, flagged, forwarded, deleted,,`), and extracts the full header block around `*** EOOH ***`.
