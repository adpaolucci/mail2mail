# Contributors, Attribution & Sources

`mail2mail` is a universal, RFC-compliant email and mailbox migration toolkit developed through a collaborative synthesis of community open-source contributions, real-world forensic migration practices, and advanced agentic AI engineering.

Every contributor, source repository, community fork author, and AI model that contributed to the architecture, hardening, and resilience of this project is gratefully acknowledged and credited below.

---

## 1. Project Maintainer & Lead

- **mail2mail Maintainers ([@adpaolucci](https://github.com/adpaolucci))**  
  Project creator, architect, and lead maintainer. Directed the unification of disparate email migration tools into a single, high-fidelity, cross-platform CLI suite (`mail2mail`, `mail2eml`, `mail2mbox`, `mail2pst`), established the zero-data-loss standard, and guided the complete RFC compliance and production hardening roadmap.

---

## 2. AI Engineering Collaborators

### **Argon 4** (Anthropic)
- **Role**: *Predecessor AI Engineering Assistant*
- **Contributions**:
  - Authored the initial architectural foundation and modular design of `mail2mail`.
  - Implemented the core multi-format streaming engine in `mail_formats.py` and the exporter modules (`mail2eml.py`, `mail2mbox.py`).
  - Drafted the extensive RFC compliance documentation suite in `docs/rfc_compliance/`, covering RFC 5322, RFC 2045–2047, RFC 2183/2231/2392, RFC 4155 (`mboxrd`), RFC 3501 (IMAP Modified UTF-7), and Microsoft Open Specifications ([MS-OXMSG], [MS-CFB], [MS-OXPROPS]).
  - Built the baseline automated test suite across unit testing modules.

### **Antigravity** (Google DeepMind)
- **Role**: *Current Agentic AI Engineering Assistant*
- **Contributions**:
  - Conducted the exhaustive code audit and multi-repository comparative analysis of all upstream forks and related open-source migration engines.
  - Implemented MAPI draft flag resolution (`PR_MESSAGE_FLAGS` clearing `MSGFLAG_UNSENT` & `MSGFLAG_SUBMIT`), curing Outlook's draft composer display and restoring true message delivery timestamps in folder list views.
  - Designed and implemented the **Zero-Data-Loss Fallback Preservation** engine, ensuring corrupted or unrenderable MIME payloads are preserved intact as raw `.eml` attachments rather than dropped.
  - Developed the Google Takeout `X-Gmail-Labels` folder hierarchy and status parser, in-flight deduplication (`--dedup`), and crash-resilient interrupted migration resuming (`--resume`).
  - Implemented RFC 2047 encoded-word attachment parameter decoding, filesystem `mtime` date fallback, Outlook attachment limit detection and registry unblocking, and 32-bit Outlook COM memory protection (`CoFreeUnusedLibraries`, message pumping).
  - Expanded and maintained the CI/CD test matrix across all 15 environments on Windows, macOS, and Linux (Python 3.9–3.13).

---

## 3. Original Upstream Source

- **Rick Madson ([@madsonrick](https://github.com/madsonrick))**  
  Author of [`madsonrick/mbox-to-pst-toolkit`](https://github.com/madsonrick/mbox-to-pst-toolkit).  
  The original source project that demonstrated driving Microsoft Outlook's local MAPI COM interface from Python for `mbox` to `pst` extraction and conversion.

---

## 4. Upstream Fork Authors & Community Contributors

We conducted an in-depth audit of all forks in the `madsonrick/mbox-to-pst-toolkit` network. The following authors contributed crucial real-world bugfixes and feature insights:

- **Guillaume Bordier ([@gbordier](https://github.com/gbordier))** — [PR #4 (`pr/fix-unsent-sent-date`)](https://github.com/madsonrick/mbox-to-pst-toolkit/pull/4)  
  *Insight*: Pinpointed that Outlook COM `Items.Add(0)` sets `MSGFLAG_UNSENT` (0x08), causing messages to display the script execution date instead of the email's true received date and opening in compose mode. Provided the MAPI property bitmask fix to clear unsent flags and introduced filesystem modification date fallback.

- **HTCGS ([@HTCGS](https://github.com/HTCGS))** — [`master` branch](https://github.com/HTCGS/mbox-to-pst-toolkit)  
  *Insight*: Identified the necessity of stamping `PR_CREATION_TIME` (`0x30070040`) and `PR_LAST_MODIFICATION_TIME` (`0x30080040`) for Outlook column sorting, along with timezone-stripping for Outlook COM timestamp marshaling.

- **`enthusedamateur` ([@enthusedamateur](https://github.com/enthusedamateur))** — [`bugfixes` branch](https://github.com/enthusedamateur/mbox-to-pst-toolkit)  
  *Insight*: Identified a critical silent data loss vulnerability during periodic PST flushes (where detaching and re-attaching PSTs left the router holding detached COM pointers), eliminated redundant `_part1` suffixes on single/initial PST files, and resolved `email.header.Header` object crashes in filename sanitization (co-authored with Claude Sonnet).

- **`dianlight` ([@dianlight](https://github.com/dianlight))** — [PR #1 (`google_keepout_extensions`)](https://github.com/madsonrick/mbox-to-pst-toolkit/pull/1)  
  *Insight*: Introduced support for Google Takeout `.zip` archives, directory-level batch MBOX scanning, and read/unread status extraction.

- **`icmaa` ([@icmaa](https://github.com/icmaa))** — [`main` branch](https://github.com/icmaa/mbox-to-pst-toolkit)  
  *Insight*: Demonstrated recursive nested folder hierarchy generation from source directory paths during PST import.

- **`Erzmaster` ([@Erzmaster](https://github.com/Erzmaster))** — [PR #2](https://github.com/madsonrick/mbox-to-pst-toolkit/pull/2)  
  *Insight*: Reported and patched the `TypeError: expected string or bytes-like object, got 'Header'` crash when regex filename sanitization encountered encoded Subject headers.

- **ITKKDF ([@ITKKDF](https://github.com/ITKKDF))** — [`mbox-to-pst-toolkit-label`](https://github.com/ITKKDF/mbox-to-pst-toolkit-label)  
  *Insight*: Enhanced documentation, command-line parameter clarity, and workflow guides for export and import operations.

- **Community Fork Mirrors**:  
  Acknowledgment to [@jozefkun](https://github.com/jozefkun), [@apersonintech](https://github.com/apersonintech), and [@ommaldon341](https://github.com/ommaldon341) for maintaining mirrors and contributing to ecosystem awareness.

---

## 5. Related Open-Source Projects & Field-Tested Prior Art

Architectural concepts and resilience safeguards were informed by inspecting prominent open-source mail conversion utilities:

- **[`3en5/EML-to-PST` (WLM2PST)](https://github.com/3en5/EML-to-PST)**  
  *Developed by KALIT / Natanor IT Services.*  
  Field-tested on a 26-year archive (301,118 unique messages / 67 GB). Contributed the **Zero-Data-Loss Fallback Preservation** concept (preserving corrupt MIME as attached fallback items), extended-length Windows paths (`\\?\`), and diagnostic procedures for Click-to-Run MAPI converter bugs.

- **[`SSDECT/eml-to-pst`](https://github.com/ssdect/eml-to-pst)**  
  *Developed by ssdect.*  
  Field-tested on an enterprise archive (110,000 messages / 117 GB). Informed our Outlook COM session recycling patterns, offline mode protection against background mailbox sync throttling, and safe 50 GB PST size ceiling checks.

- **`ordekalo/mbox_to_pst_converter` ([@ordekalo](https://github.com/ordekalo))**  
  Early reference for combining Python's `mailbox` module with Outlook automation and `pypff`.

- **[`MimeKit`](https://github.com/jstedfast/MimeKit) & [`MsgKit`](https://github.com/Sicos1977/MsgKit)**  
  Created by Jeffrey Stedfast and Ramon de Klein. The definitive reference implementations for robust MIME parser edge cases and Microsoft Outlook `.msg` Compound File Binary structures.

---

## 6. Standards & Specifications

`mail2mail` adheres strictly to standard specifications published by the Internet Engineering Task Force (IETF) and Microsoft Open Specifications:
- **RFC 5322**: Internet Message Format
- **RFC 2045, 2046, 2047, 2231**: Multipurpose Internet Mail Extensions (MIME)
- **RFC 2183**: Content-Disposition Header Field
- **RFC 2392**: Content-ID and Message-ID Uniform Resource Locators (`cid:`)
- **RFC 4155**: The `application/mbox` Media Type (`mboxrd` standard)
- **RFC 3501 §5.1.3**: IMAP Modified UTF-7 Mailbox Naming
- **[MS-OXMSG]**: Outlook Item (.msg) Binary File Format
- **[MS-CFB]**: Compound File Binary File Format (OLE2)
- **[MS-OXPROPS]**: Exchange Server / MAPI Protocol Properties
