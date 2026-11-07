# `mail2mail` RFC & Format Specifications Compliance Reference

This directory documents the IETF RFCs, Microsoft Open Specifications, and de-facto vendor mailbox specifications implemented and verified by the `mail2mail` suite ([`mail_formats.py`](../../mail_formats.py), [`mail2eml.py`](../../mail2eml.py), [`mail2mbox.py`](../../mail2mbox.py), [`mail2pst.py`](../../mail2pst.py), and [`mail2mail.py`](../../mail2mail.py)).

---

## Reference Documents in This Folder

1. **[`RFC5322_RFC2045_MIME_AND_HEADERS.md`](../../docs/rfc_compliance/RFC5322_RFC2045_MIME_AND_HEADERS.md)**
   - **RFC 5322** (Internet Message Format; obsoletes RFC 2822 / RFC 822)
   - **RFC 2045 / RFC 2046** (MIME Part 1 & Part 2: Media Types, Multipart Boundaries, Content-Transfer-Encoding)
   - **RFC 2047** (MIME Part 3: Message Header Extensions for Non-ASCII Text / Encoded-Words)
   - **RFC 2183 / RFC 2231** (`Content-Disposition` Header Field & Parameter Value Continuations/Character Sets)
   - **RFC 2392** (`Content-ID` and `cid:` URLs for MHTML Inline Multipart/Related Assets)
   - **RFC 6532** (Internationalized Email Headers / UTF-8)
   - **RFC 8551 / RFC 3156** (S/MIME v4.0 `application/pkcs7-mime`, `application/pkcs7-signature`, and PGP/MIME `application/pgp-signature`)

2. **[`RFC4155_MBOX_AND_MAILBOX_FORMATS.md`](../../docs/rfc_compliance/RFC4155_MBOX_AND_MAILBOX_FORMATS.md)**
   - **RFC 4155** (`application/mbox` Media Type & `From <sender> <ctime>` Envelope Framing)
   - **`mboxrd` vs. `mboxo`** (`>From ` / `>>From ` Reversible Body Escaping)
   - **RFC 3501 §5.1.3** (IMAP4rev1 Modified UTF-7 Mailbox Naming Convention)
   - **Mozilla / Thunderbird `ImapMail` & `Mail`** (`X-Mozilla-Status`, `X-Mozilla-Status2`, `.sbd` Subfolder Containers)
   - **BSD `Status` & UW-IMAP `X-Status`** Header Flag Conventions
   - **D. J. Bernstein Maildir & Courier/Dovecot Maildir++** (`cur/`, `new/`, `tmp/`, `:2,FRSPT` / `;2,FRSPT`, `.Folder.Subfolder`)
   - **Apple Mail `.emlx` / `.partial.emlx` & `*.mbox` Bundles** (Byte-Count Framing, XML `<plist>` 32-bit Flags Bitmask, `Data/0..9/Messages`)
   - **RAND MH (`.mh_sequences`), MMDF (`\x01\x01\x01\x01`), and Emacs Rmail Babyl (`BABYL OPTIONS:`)**

3. **[`MS_OXMSG_CFBF_MAPI_AND_OLM.md`](../../docs/rfc_compliance/MS_OXMSG_CFBF_MAPI_AND_OLM.md)**
   - **[MS-CFB]** (Compound File Binary File Format — v3 512-byte & v4 4096-byte Sectors, DIFAT, FAT, MiniFAT, Red-Black Directory Tree)
   - **[MS-OXMSG]** (Outlook Item `.msg` File Format — Property Stream, Named/Substg Streams, Recipient & Attachment Storages)
   - **[MS-OXPROPS] & [MS-OXCMSG]** (MAPI Property Tags, DASL Schemas, `PT_SYSTIME` UTC `VT_DATE`, Verb & Icon Indices, MHTML `PR_ATTACH_FLAGS`)
   - **Outlook for Mac `.olm`** (`com.microsoft.__Messages` XML Schema & ZIP Archive Structure)

---

## Master Standards Compliance Matrix

| Standard / Specification | Scope | Primary Implementation | Automated Test Coverage |
| :--- | :--- | :--- | :--- |
| **RFC 5322** | Internet Message Format headers, dates, addresses, folding | [`safe_get_header`](../../mail_formats.py#L246-L264), [`safe_get_all_headers`](../../mail_formats.py#L267-L283), [`addresses_to_str`](../../mail2pst.py#L134-L156) | [`tests/test_rfc_compliance.py`](../../tests/test_rfc_compliance.py) |
| **RFC 2045 / 2046** | MIME multipart bodies, `base64` / `quoted-printable` decoding | [`pick_body`](../../mail2pst.py#L207-L243), [`add_attachments`](../../mail2pst.py#L246-L345) | [`tests/test_rfc_compliance.py`](../../tests/test_rfc_compliance.py) |
| **RFC 2047** | Encoded-word headers (`=?charset?B/Q?...?=`) | [`decode_mime_words`](../../mail_formats.py#L236-L243) | [`tests/test_rfc_compliance.py`](../../tests/test_rfc_compliance.py) |
| **RFC 2183 / 2231** | `Content-Disposition` (`inline` / `attachment`) & filenames | [`add_attachments`](../../mail2pst.py#L246-L345) | [`tests/test_rfc_compliance.py`](../../tests/test_rfc_compliance.py) |
| **RFC 2392** | `Content-ID` (`cid:`) inline HTML images (`PR_ATTACH_CONTENT_ID`) | [`add_attachments`](../../mail2pst.py#L329-L337), [`parse_msg_bytes`](../../mail_formats.py#L1028-L1037) | [`tests/test_rfc_compliance.py`](../../tests/test_rfc_compliance.py) |
| **RFC 3501 §5.1.3** | IMAP4rev1 Modified UTF-7 folder name encoding/decoding | [`decode_imap_utf7`](../../mail_formats.py#L130-L160), [`encode_imap_utf7`](../../mail_formats.py#L163-L193) | [`tests/test_rfc_compliance.py`](../../tests/test_rfc_compliance.py) |
| **RFC 4155 & `mboxrd`** | `From <sender> <asctime>` framing & `>From ` body quoting | [`iter_mbox_stream`](../../mail_formats.py#L1561-L1620), [`format_mboxrd_entry`](../../mail2mbox.py#L97-L135) | [`tests/test_rfc_compliance.py`](../../tests/test_rfc_compliance.py), [`tests/test_exporters_and_cli.py`](../../tests/test_exporters_and_cli.py) |
| **RFC 8551 / 3156** | S/MIME (`pkcs7-mime`, `pkcs7-signature`) & PGP (`pgp-signature`) | [`populate_mailitem`](../../mail2pst.py#L414-L421), [`add_attachments`](../../mail2pst.py#L277-L279) | [`tests/test_rfc_compliance.py`](../../tests/test_rfc_compliance.py) |
| **[MS-CFB] & [MS-OXMSG]** | OLE2 Compound File Binary & Outlook `.msg` streams/storages | [`MiniOleReader`](../../mail_formats.py#L649-L837), [`parse_msg_bytes`](../../mail_formats.py#L878-L1046) | [`tests/test_mail_formats.py`](../../tests/test_mail_formats.py) |
| **[MS-OXPROPS]** | MAPI Property Tags, naive UTC `VT_DATE`, sender/verb/icon props | [`populate_mailitem`](../../mail2pst.py#L393-L479), [`iter_pst_records`](../../mail_formats.py#L1246-L1466) | [`tests/test_exporters_and_cli.py`](../../tests/test_exporters_and_cli.py) |
| **Maildir / Maildir++** | `cur/`, `new/`, `tmp/`, `:2,FRSPT` / `;2,FRSPT`, `.Folder.Sub` | [`parse_maildir_flags`](../../mail_formats.py#L469-L511), [`normalize_maildir_folder_parts`](../../mail_formats.py#L514-L539) | [`tests/test_mail_formats.py`](../../tests/test_mail_formats.py) |
| **Apple Mail `.emlx`** | Byte-count prefix, XML `<plist>` 32-bit flags, `.mbox` bundles | [`parse_emlx_bytes`](../../mail_formats.py#L569-L619), [`normalize_apple_mail_folder_parts`](../../mail_formats.py#L622-L644) | [`tests/test_mail_formats.py`](../../tests/test_mail_formats.py) |
| **Outlook for Mac `.olm`** | `com.microsoft.__Messages` XML & ZIP attachments | [`iter_olm_records`](../../mail_formats.py#L1078-L1241) | [`tests/test_mail_formats.py`](../../tests/test_mail_formats.py) |
| **MH / MMDF / Babyl** | `.mh_sequences`, `\x01\x01\x01\x01`, `BABYL OPTIONS:` | [`parse_mh_sequences`](../../mail_formats.py#L1470-L1496), [`iter_mmdf_stream`](../../mail_formats.py#L1622-L1653), [`iter_babyl_stream`](../../mail_formats.py#L1655-L1702) | [`tests/test_mail_formats.py`](../../tests/test_mail_formats.py) |
