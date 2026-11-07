# Microsoft Open Specifications ([MS-CFB], [MS-OXMSG], [MS-OXPROPS]) & Outlook for Mac `.olm`

This document details how `mail2mail` implements Microsoft's Compound File Binary Format ([MS-CFB]), Outlook `.msg` Item Format ([MS-OXMSG]), MAPI Property Tags ([MS-OXPROPS]), and Outlook for Mac `.olm` archives.

---

## 1. [MS-CFB] — Compound File Binary File Format (OLE2)

[`MiniOleReader`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L649-L837) in [`mail_formats.py`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py) is a zero-dependency, pure-Python parser for Microsoft Compound File Binary (`.msg`) files:

| [MS-CFB] Section | Structure | Implementation Details |
| :--- | :--- | :--- |
| **§2.2 Header** | 512-byte CFBF Header | Validates signature `D0 CF 11 E0 A1 B1 1A E1` (`OLE2_MAGIC`), reads `SectorShift` (v3 = `9` $\to$ 512B sectors; v4 = `12` $\to$ 4096B sectors), `MiniSectorShift` (`6` $\to$ 64B mini-sectors), `MiniStreamCutoffSize` (`4096` bytes), `FirstDirSectorLocation`, `FirstMiniFATSectorLocation`, and the 109 header DIFAT entries. |
| **§2.3 DIFAT & FAT** | Sector Allocation Chain | Traverses header DIFAT entries (`offset 76..512`) and chained DIFAT sectors (`first_difat_sec`) to assemble the complete File Allocation Table (`self.fat`), guarding against circular chains (`ENDOFCHAIN = 0xFFFFFFFE`, `FREESECT = 0xFFFFFFFF`). |
| **§2.4 MiniFAT & MiniStream** | Small-Stream Allocation (`< 4096B`) | Reads the `Root Entry` stream (`obj_type == 5`) as the `mini_stream` container and resolves streams smaller than `mini_cutoff` (`4096B`) via `self.minifat` (64-byte slices). |
| **§2.6 Directory Entries** | 128-byte Red-Black Tree Nodes | Parses UTF-16LE entry names (`name_len` at offset `64`), `obj_type` (`1` = Storage, `2` = Stream, `5` = Root), `left_sibling` (`68`), `right_sibling` (`72`), `child_id` (`76`), `start_sec` (`116`), and 64-bit `size` (`120`, masked to 32-bit for v3 files per [MS-CFB] §2.6.4). |

---

## 2. [MS-OXMSG] — Outlook Item (`.msg`) File Format

[`parse_msg_bytes`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L878-L1046) reconstructs a full RFC 5322 `EmailMessage` (with HTML/plain body, recipients, attachments, inline CIDs, and status flags) from an OLE2 `.msg` structure:

### 2.1 Variable-Length Property Streams (`__substg1.0_XXXXYYYY`)
- **`XXXX`** = 16-bit MAPI Property ID in uppercase hex; **`YYYY`** = 16-bit Property Type:
  - `001F` (`PT_UNICODE`): Decoded as UTF-16LE (`rstrip("\x00")`).
  - `001E` (`PT_STRING8`): Decoded as `cp1252` / `utf-8`.
  - `0102` (`PT_BINARY`): Raw bytes.
- Key top-level streams extracted:
  - `0037` (`PidTagSubject`), `007D` (`PidTagTransportMessageHeaders`), `1035` (`PidTagInternetMessageId`), `1000` (`PidTagBody`), `1013` (`PidTagHtml` / `PR_HTML` binary stream), `0C1A` (`PidTagSenderName`), `5D01` / `0C1F` (`PidTagSenderSmtpAddress` / `PidTagSenderEmailAddress`), `0042` / `5D02` (`PidTagSentRepresentingName` / `PidTagSentRepresentingSmtpAddress`), `0E04` (`PidTagDisplayTo`), `0E03` (`PidTagDisplayCc`), `0E02` (`PidTagDisplayBcc`).

### 2.2 Fixed-Length Property Stream (`__properties_version1.0`)
- Header size ([MS-OXMSG] §2.4.1):
  - Top-level message root: **32 bytes** (`hdr_size = 32`)
  - Sub-storages (`__recip_version1.0_#XXXXXXXX`, `__attach_version1.0_#XXXXXXXX`): **24 bytes**
- Followed by 16-byte fixed property entries (`<IIQ`: `prop_tag`, `flags`, `value_8bytes`):
  - `0x0E070003` (`PidTagMessageFlags`): `0x01` = `MSGFLAG_READ`
  - `0x10900003` (`PidTagFlagStatus`): `2` = `olFlagMarked`
  - `0x10810003` (`PidTagLastVerbExecuted`): `102`/`103` = Replied/Reply-All, `104` = Forwarded
  - `0x00390040` (`PidTagClientSubmitTime`) / `0x0E060040` (`PidTagMessageDeliveryTime`): 64-bit Windows `FILETIME` (100-nanosecond intervals since `1601-01-01 00:00:00 UTC`), converted to RFC 5322 `Date` via [`_filetime_to_datetime`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L840-L848).

### 2.3 Recipient & Attachment Sub-Storages
- **`__recip_version1.0_#XXXXXXXX`**: Reads `PidTagRecipientType` (`0x0C150003`: `1` = `MAPI_TO`, `2` = `MAPI_CC`, `3` = `MAPI_BCC`), `PidTagDisplayName` (`3001`), and `PidTagSmtpAddress` (`39FE`) / `PidTagEmailAddress` (`3003`).
- **`__attach_version1.0_#XXXXXXXX`**: Reads `PidTagAttachDataBinary` (`3701`), `PidTagAttachLongFilename` (`3707`) / `PidTagAttachFilename` (`3704`), `PidTagAttachMimeTag` (`370E`), and `PidTagAttachContentId` (`3712`).

---

## 3. [MS-OXPROPS] & MAPI DASL Properties in `mail2pst.py`

When importing into an Outlook `.pst` store via `PropertyAccessor.SetProperty`, [`mail2pst.py`](file:///c:/Users/andre/Desktop/all2pst/mail2pst.py) sets the following DASL properties (`http://schemas.microsoft.com/mapi/proptag/0x...`):

| Constant in `mail2pst.py` | DASL Property Tag | MAPI Canonical Name | Purpose |
| :--- | :--- | :--- | :--- |
| `PR_TRANSPORT_MESSAGE_HEADERS_W` | `0x007D001F` (`001E` fallback) | `PidTagTransportMessageHeaders` | Preserves full original RFC 5322 headers (`View -> Message Options` in Outlook) |
| `PR_INTERNET_MESSAGE_ID_W` | `0x1035001F` (`001E` fallback) | `PidTagInternetMessageId` | Preserves original `Message-ID` |
| `PR_SENDER_NAME_W` | `0x0C1A001F` | `PidTagSenderName` | Populates `From` display name |
| `PR_SENDER_EMAIL_ADDRESS_W` | `0x0C1F001F` | `PidTagSenderEmailAddress` | Populates `From` SMTP address |
| `PR_SENDER_ADDRTYPE_W` | `0x0C1E001F` | `PidTagSenderAddressType` | Set to `"SMTP"` |
| `PR_SENT_REPRESENTING_NAME_W` | `0x0042001F` | `PidTagSentRepresentingName` | Populates sender column in folder views |
| `PR_SENT_REPRESENTING_EMAIL_W` | `0x0065001F` | `PidTagSentRepresentingEmailAddress` | Populates sender SMTP address |
| `PR_SENT_REPRESENTING_ADDRTYPE_W` | `0x0064001F` | `PidTagSentRepresentingAddressType` | Set to `"SMTP"` |
| `PR_CLIENT_SUBMIT_TIME` | `0x00390040` (`PT_SYSTIME`) | `PidTagClientSubmitTime` | Original `Sent` timestamp (naive UTC `datetime` for `VT_DATE`) |
| `PR_MESSAGE_DELIVERY_TIME` | `0x0E060040` (`PT_SYSTIME`) | `PidTagMessageDeliveryTime` | Original `Received` timestamp (naive UTC `datetime` for `VT_DATE`) |
| `PR_ATTACH_CONTENT_ID_W` | `0x3712001F` (`001E` fallback) | `PidTagAttachContentId` | Links inline `<img src="cid:...">` attachments |
| `PR_ATTACH_FLAGS` | `0x37140003` (`PT_LONG`) | `PidTagAttachFlags` | Set to `4` (`ATT_MHTML_REF`) for inline HTML images |
| `PR_LAST_VERB_EXECUTED` | `0x10810003` (`PT_LONG`) | `PidTagLastVerbExecuted` | `102` = Replied (`EXCHIVERB_REPLYTOSENDER`), `104` = Forwarded (`EXCHIVERB_FORWARD`) |
| `PR_ICON_INDEX` | `0x10800003` (`PT_LONG`) | `PidTagIconIndex` | `261` (`0x105`) = Replied icon, `262` (`0x106`) = Forwarded icon |

---

## 4. Outlook for Mac `.olm` Archive Format

An `.olm` file is a ZIP archive produced by Outlook for Mac containing XML message descriptors under `Accounts/<account>/com.microsoft.__Messages/<Folder>/<Subfolder>/message_XXXXX.xml` and binary attachments under `Attachments/`:
- [`iter_olm_records`](file:///c:/Users/andre/Desktop/all2pst/mail_formats.py#L1078-L1241) parses `OPFMessageCopySubject`, `OPFMessageCopySentTime` / `OPFMessageCopyReceivedTime` (ISO-8601), `OPFMessageCopyMessageID`, `OPFMessageGetIsRead`, `OPFMessageCopyFlagStatus`, `OPFMessageCopyFromAddresses`, `OPFMessageCopyToAddresses`, `OPFMessageCopyCCAddresses`, `OPFMessageCopyBCCAddresses`, `OPFMessageCopyBody`, `OPFMessageCopyHTMLBody`, and `OPFMessageCopyAttachmentList` (`OPFAttachmentURL`, `OPFAttachmentName`, `OPFAttachmentContentType`, `OPFAttachmentContentID`).
