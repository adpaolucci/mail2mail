# Security Policy

## Local-Only Processing Guarantee

`mail2mail` processes all email messages, attachments, and mailbox archives **100% locally** on your machine:
- Zero network requests or external telemetry are ever made by any module (`mail_formats.py`, `mail2eml.py`, `mail2mbox.py`, `mail2pst.py`, `mail2mail.py`).
- `.gitignore` is pre-configured to ignore `.pst`, `.ost`, `.olm`, `.msg`, `.mbox`, `.eml`, and `.emlx` files to prevent accidental commits of personal or confidential email archives.

## Built-In Defensive Protections

1. **Path Traversal & Filename Sanitization**:
   - Folder names and attachment filenames extracted from untrusted messages or archives are sanitized via [`sanitize_name`](mail_formats.py), stripping path separators (`/`, `\`), control characters (`0x00..0x1F`, `0x7F`), and Windows reserved device names (`CON`, `PRN`, `AUX`, `NUL`, `COM1..9`, `LPT1..9`).
2. **OLE2 (`[MS-CFB]`) & MH Sequence Bounds Guards**:
   - [`MiniOleReader`](mail_formats.py) enforces strict sector-chain cycle detection (`visited` set and maximum sector count bounds) so corrupted or maliciously crafted `.msg` files cannot trigger infinite loops or memory exhaustion.
   - [`parse_mh_sequences`](mail_formats.py) caps `.mh_sequences` numeric range expansions (`start-end`) to prevent range-bomb memory exhaustion.
3. **Malformed Header Isolation**:
   - [`safe_get_header`](mail_formats.py) and [`safe_get_all_headers`](mail_formats.py) isolate Python `HeaderParseError` exceptions when parsing malformed RFC 5322 headers.

## Supported Versions

| Version | Supported |
| :--- | :--- |
| `2.2.x` | :white_check_mark: |
| `2.1.x` | :white_check_mark: |
| `< 2.1.0` | :x: |

## Reporting a Vulnerability

If you discover a security vulnerability (such as an archive extraction bypass or parser denial-of-service), please open a private security advisory on GitHub rather than posting a public issue containing exploit payloads.
