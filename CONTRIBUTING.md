# Contributing to `mail2mail`

Thank you for your interest in contributing to **`mail2mail`**!

---

## Development Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/adpaolucci/mail2mail.git
   cd mail2mail
   ```
2. Install in editable mode:
   ```bash
   pip install -e .
   ```
   *(Note: `mail2eml`, `mail2mbox`, and `mail_formats` use only the Python 3.9+ standard library; `pywin32` and Classic Microsoft Outlook on Windows are only required for `.pst` input/output).*

---

## Running the Test Suite

Before submitting a pull request, run the full automated test suite from the repository root:

```bash
python -m unittest discover -s tests -v
```

All tests in `tests/` run on Windows, macOS, and Linux (using automatic COM/MAPI stubs when `pywin32` is not installed).

---

## Guidelines for Pull Requests

1. **RFC & Specification Compliance**:
   - Check [`docs/rfc_compliance/`](docs/rfc_compliance/README.md) when modifying header parsing, MIME body/attachment handling, `mboxrd` quoting, IMAP Modified UTF-7, or MAPI properties.
   - Real-world mail archives often contain malformed headers or non-standard encodings; always use [`safe_get_header`](mail_formats.py) / [`safe_get_all_headers`](mail_formats.py) and guard against parser exceptions so a single corrupted message never aborts a multi-gigabyte migration.
2. **Streaming & Memory Efficiency**:
   - Keep mailbox readers streaming ($O(1)$ memory per message) so multi-gigabyte `.mbox`, `.mbox.gz`, and `.pst` archives can be processed without loading the entire archive into RAM.
3. **Tests & Documentation**:
   - Add unit tests in `tests/` for any new format or bug fix, and update `README.md` / `docs/rfc_compliance/` as appropriate.
