# Security Policy

## Local-Only Processing Guarantee

`mail2mail` processes all local email messages, attachments, and mailbox archives **100% locally** on your machine:
- Local conversion and auditing modules (`mail_formats.py`, `mail_audit.py`, `mail_filter.py`, `mail_shutdown.py`, `mail2eml.py`, `mail2mbox.py`, `mail2pst.py`, `mail2mail.py`, `mail2mcp.py`, `file_picker.py`) make zero network requests or external telemetry calls.
- When explicitly invoked, [`mail2jmap.py`](mail2jmap.py) (`mail2mail jmap`) connects exclusively to the user-configured `https://` or `http://` JMAP server endpoint. Non-HTTP(S) URL schemes (such as `file://`) are rejected.
- `.gitignore` is pre-configured to ignore `.pst`, `.ost`, `.olm`, `.msg`, `.mbox`, `.eml`, and `.emlx` files to prevent accidental commits of personal or confidential email archives.

## Built-In Defensive Protections

1. **Path Traversal, Archive & Filename Sanitization**:
   - Folder names and attachment filenames extracted from untrusted messages or archives are sanitized via [`sanitize_name`](mail_formats.py), stripping path separators (`/`, `\`), control characters (`0x00..0x1F`, `0x7F`), and Windows reserved device names (`CON`, `PRN`, `AUX`, `NUL`, `COM1..9`, `LPT1..9`).
   - ZIP, TAR, and OLM archive extractors strip `.` and `..` path traversal segments before constructing target folder hierarchies, and reject OLM XML payloads containing `<!ENTITY` declarations.
2. **OLE2 (`[MS-CFB]`) & MH Sequence Bounds Guards**:
   - [`MiniOleReader`](mail_formats.py) enforces strict sector-chain cycle detection (`visited` set and maximum sector count bounds) so corrupted or maliciously crafted `.msg` files cannot trigger infinite loops or memory exhaustion.
   - [`parse_mh_sequences`](mail_formats.py) caps `.mh_sequences` numeric range expansions (`start-end`) to prevent range-bomb memory exhaustion.
3. **Malformed Header Isolation, CSV Formula Neutralization & Shell Escaping**:
   - [`safe_get_header`](mail_formats.py) and [`safe_get_all_headers`](mail_formats.py) isolate Python `HeaderParseError` exceptions when parsing malformed RFC 5322 headers.
   - [`csv_safe_cell`](mail_audit.py) neutralizes spreadsheet formula-injection prefixes (`=`, `+`, `-`, `@`, `\t`, `\r`) when exporting CSV audit manifests while preserving lossless `--resume` round-tripping.
   - [`file_picker.py`](file_picker.py) sanitizes and single-quote-escapes dialog parameters passed to Windows PowerShell fallback pickers.

## Supported Versions

| Version | Supported |
| :--- | :--- |
| `2.3.x` | :white_check_mark: |
| `2.2.x` | :white_check_mark: |
| `2.1.x` | :white_check_mark: |
| `< 2.1.0` | :x: |

## Reporting a Vulnerability

If you discover a security vulnerability (such as an archive extraction bypass or parser denial-of-service), please open a private security advisory on GitHub rather than posting a public issue containing exploit payloads:
`https://github.com/adpaolucci/mail2mail/security/advisories/new`
