# RFC 5322, RFC 2045–2047, RFC 2183/2231/2392, & S/MIME / PGP Compliance

This document specifies how `mail2mail` parses, normalizes, and generates Internet Message Format (RFC 5322) headers and Multipurpose Internet Mail Extensions (MIME) bodies across [`mail_formats.py`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py), [`mail2eml.py`](file:///c:/Users/andre/Desktop/all2pst/mail2eml.py), [`mail2mbox.py`](file:///c:/Users/andre/Desktop/all2pst/mail2mbox.py), and [`mail2pst.py`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py).

---

## 1. RFC 5322 — Internet Message Format (Obsoletes RFC 2822 & RFC 822)

### 1.1 Header Syntax, Folding, and Malformed Header Resilience (§2.2, §2.2.3)
- **Standard Requirement**: Header fields consist of a field name followed by a colon (`:`), a field body, and `CRLF`, with optional folding white space (`FWS` — `CRLF` followed by `SP` or `HTAB`).
- **Real-World Challenge**: Historical archives (1990s–2020s) frequently contain non-RFC-compliant headers (e.g., `Date: 2020-99-99`, unquoted special characters in address headers, or `Content-Type` parameters with unescaped quotes) that raise `email.errors.HeaderParseError`, `IndexError`, or `ValueError` when accessed via `msg.get(name)` under Python's `email.policy.default`.
- **`mail2mail` Implementation**:
  - [`safe_get_header`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L246-L264) and [`safe_get_all_headers`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L267-L283) attempt structured `policy.default` access first (`msg.get` / `msg.get_all`), and if a header defect triggers an exception, fall back to iterating `msg.raw_items()` and decoding raw header values via [`decode_mime_words`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L236-L243).
  - [`HEADER_START_RE`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L96) (`^[A-Za-z0-9!#$%&'*+\-/^_`{|}~]+:\s*`) validates RFC 5322 §3.6.8 `ftext` field-name tokens (`33..57`, `59..126`) when sniffing header-only or `From `-less messages.

### 1.2 Date & Time Specification (§3.3)
- **Standard Requirement**: `Date` headers use `[day-of-week ","] day month year hour ":" minute [":" second] zone` (e.g., `Fri, 06 Nov 2020 15:30:00 -0500`).
- **`mail2mail` Implementation**:
  - [`pick_year_month`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L286-L300) uses `email.utils.parsedate_to_datetime`, clamping out-of-bounds years outside `[1900, 2100]` to `(1970, 1)` so corrupted spam headers do not create garbage year folders or invalid MAPI `FILETIME` structures.
  - [`to_mapi_utc`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py#L171-L186) converts timezone-aware `datetime` values to UTC (`dt.astimezone(timezone.utc)`) and strips `tzinfo` before passing to `pywin32` `PropertyAccessor.SetProperty` (`PR_CLIENT_SUBMIT_TIME` / `PR_MESSAGE_DELIVERY_TIME`), preventing `pywin32`'s `VT_DATE` converter from shifting UTC timestamps into local system time.

### 1.3 Address Specification & Multiple Recipient Headers (§3.4, §3.6.3)
- **Standard Requirement**: `From`, `Sender`, `Reply-To`, `To`, `Cc`, and `Bcc` contain `address-list` values (`display-name <addr-spec>` or `addr-spec`). Although RFC 5322 §3.6 specifies `0..1` occurrences of `To`/`Cc`/`Bcc`, legacy MTAs occasionally emit multiple `To:` or `Cc:` header lines.
- **`mail2mail` Implementation**:
  - [`safe_get_all_headers`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L267-L283) collects all occurrences of `From`, `To`, `Cc`, and `Bcc` and passes them to `email.utils.getaddresses`.
  - [`addresses_to_str`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py#L134-L156) formats recipient lists as `; `-separated entries for Outlook's Object Model while replacing any `;` inside a recipient's `display-name` (e.g., `"Last; First" <user@example.com>`) with `,` so Outlook does not split a single recipient into two broken addresses.

### 1.4 Message-ID & Transport Headers (§3.6.4)
- **Standard Requirement**: `Message-ID` uniquely identifies a message (`<id-left@id-right>`).
- **`mail2mail` Implementation**:
  - [`populate_mailitem`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py#L453-L456) writes `Message-ID` into `PR_INTERNET_MESSAGE_ID_W` (`0x1035001F`, with ANSI fallback `0x1035001E`) and preserves the entire original RFC 5322 header block in `PR_TRANSPORT_MESSAGE_HEADERS_W` (`0x007D001F`) via [`build_headers_text`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py#L348-L361) (supporting both `\r\n\r\n` and `\n\n` header-body separators).
  - All MAPI property writes occur **after** setting `mail_item.HTMLBody` / `mail_item.Body` so Outlook's MHTML body synchronization does not overwrite `PR_TRANSPORT_MESSAGE_HEADERS` or `PR_SENDER_*`.

---

## 2. RFC 2047 — MIME Part 3: Encoded-Word Extensions for Non-ASCII Headers

- **Standard Requirement**: Non-ASCII characters in `Subject`, `From`, `To`, and attachment names are encoded as `=?charset?encoding?encoded-text?=` where `encoding` is `B` (Base64) or `Q` (Quoted-Printable with `_` representing `0x20`).
- **`mail2mail` Implementation**:
  - [`decode_mime_words`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L236-L243) uses `email.header.decode_header` + `email.header.make_header` with graceful string fallback on malformed charsets (`unknown-8bit`, `x-user-defined`, truncated Base64).

---

## 3. RFC 2045, RFC 2046, RFC 2183, RFC 2231, & RFC 2392 — MIME Bodies & Attachments

### 3.1 Multipart Body Selection (`multipart/alternative`, `multipart/mixed`, `multipart/related`)
- [`pick_body`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py#L207-L243) walks the MIME tree, selects the primary inline `text/html` and `text/plain` parts (ignoring parts with `Content-Disposition: attachment` or an explicit `filename`), decodes payloads using the declared charset (falling back to `utf-8` with `errors="replace"` on unknown charsets), and records the Python `id(part)` of body parts in `body_parts: Set[int]`.
- [`add_attachments`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py#L246-L345) skips any part whose `id(part)` is in `body_parts`, preventing inline `text/plain` or `text/html` bodies from being duplicated as `.txt`/`.html` file attachments while preserving legitimate inline `.csv`/`.txt` attachments that carry a `filename`.

### 3.2 Inline MHTML Assets (`Content-ID` / `cid:` — RFC 2392)
- When a MIME part carries a `Content-ID` header (e.g., `Content-ID: <img001@example.com>`) referenced by `<img src="cid:img001@example.com">` in `HTMLBody`:
  - [`add_attachments`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py#L329-L337) strips angle brackets (`<>`), sets `PR_ATTACH_CONTENT_ID_W` (`0x3712001F`), and sets `PR_ATTACH_FLAGS` (`0x37140003`) to `4` (`ATT_MHTML_REF`) **before** `mail_item.HTMLBody` is assigned, ensuring Outlook links and renders inline images inside the message body.

### 3.3 Nested `message/rfc822`, S/MIME (RFC 8551), and PGP/MIME (RFC 3156)
- **Forwarded Messages (`message/rfc822`)**: Extracted as `.eml` attachments (`forwarded_message_<N>.eml`) with their complete inner RFC 5322 headers and MIME structure intact.
- **Detached Signatures (`application/pkcs7-signature`, `application/pgp-signature`)**: Preserved as `smime.p7s` / `signature.asc` attachments.
- **Opaque S/MIME (`application/pkcs7-mime` / `smime.p7m`)**: When the top-level message is a single-part opaque S/MIME envelope, `mail2pst` attaches the binary `smime.p7m` payload and populates a descriptive placeholder body (`[S/MIME Encrypted or Opaque-Signed Message - See attached smime.p7m]`) so nothing is lost.
