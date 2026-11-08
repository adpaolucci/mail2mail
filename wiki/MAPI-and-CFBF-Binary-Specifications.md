# MAPI & CFBF Binary Specifications ([MS-OXMSG], [MS-CFB], [MS-OXPROPS])

Both `mail2rust` and `mail2mail` implement pure, zero-dependency binary parsers and writers for Microsoft's **Compound File Binary Format ([MS-CFB])**, **Outlook Item Format ([MS-OXMSG])**, and **MAPI Property Tags ([MS-OXPROPS])**.

This enables full forensic reading, parsing, and direct PST/MSG creation on **Linux, macOS, and Windows without requiring Microsoft Outlook or Windows MAPI DLLs installed**.

---

## 1. [MS-CFB] — Compound File Binary File Format (OLE2)

Microsoft Compound Files (used by `.msg` single-item files and internal MAPI storage) act as a complete FAT filesystem inside a single file.

### CFBF Header Layout (§2.2)
```text
Offset  Length  Field                  Description / Expected Value
--------------------------------------------------------------------------------
0x0000  8       Header Signature       D0 CF 11 E0 A1 B1 1A E1 (OLE2_MAGIC)
0x001E  2       Sector Shift           0x0009 (512B, v3) or 0x000C (4096B, v4)
0x0020  2       Mini Sector Shift      0x0006 (64B mini sectors)
0x002C  4       FAT Sectors Count      Number of sectors allocated for FAT
0x0030  4       First Dir Sector       Starting sector of Directory Stream
0x0038  4       Mini Stream Cutoff     0x00001000 (Streams < 4096B use MiniFAT)
0x003C  4       First MiniFAT Sector   Starting sector of MiniFAT chain
0x0044  4       First DIFAT Sector     Starting sector of Double-Indirect FAT
0x004C  436     Header DIFAT Array     First 109 FAT sector locations
```

### Key Parsing Implementation Details:
* **DIFAT & FAT Chaining:** The engine traverses the 109 header DIFAT entries (`offset 76..512`) and follows chained DIFAT sectors (`first_difat_sec`) to reconstruct the complete File Allocation Table (`FAT`), guarding against circular loops (`ENDOFCHAIN = 0xFFFFFFFE`, `FREESECT = 0xFFFFFFFF`).
* **MiniFAT & MiniStream Container:** For streams smaller than 4,096 bytes, standard FAT allocation is inefficient. The `Root Entry` directory node acts as the `MiniStream` container, and individual small streams are resolved in 64-byte slices via the `MiniFAT` chain.
* **Red-Black Directory Tree:** Directory entries are stored as 128-byte nodes containing UTF-16LE stream names, entry types (`1 = Storage`, `2 = Stream`, `5 = Root`), left/right child pointers, and stream length.

---

## 2. [MS-OXMSG] — Outlook `.msg` Structure & Property Streams

A valid `.msg` container stores message metadata across two primary structures: **variable-length sub-storages** and the **fixed-length property stream**.

### 2.1 Variable-Length Property Streams (`__substg1.0_XXXXYYYY`)
Streams named `__substg1.0_XXXXYYYY` encode specific MAPI properties:
* **`XXXX`** = 16-bit Property ID in uppercase hexadecimal.
* **`YYYY`** = 16-bit Property Type:
  * `001F` (`PT_UNICODE`): Decoded as UTF-16LE without null terminator.
  * `001E` (`PT_STRING8`): Decoded as `Windows-1252` or `UTF-8`.
  * `0102` (`PT_BINARY`): Raw binary payload (e.g. attachment data, RTF compressed stream).

#### Core Top-Level Property IDs:
| Property ID | Type | MAPI Tag Constant | Purpose |
| :--- | :--- | :--- | :--- |
| `0037` | `001F` | `PidTagSubject` | Message Subject |
| `007D` | `001F` | `PidTagTransportMessageHeaders` | Complete original RFC 5322 MIME headers |
| `1035` | `001F` | `PidTagInternetMessageId` | Standard `Message-ID` string |
| `1000` | `001F` | `PidTagBody` | Plain-text body |
| `1013` | `0102` | `PidTagHtml` | Full HTML body binary stream |
| `0C1A` | `001F` | `PidTagSenderName` | Sender display name |
| `0C1F` | `001F` | `PidTagSenderEmailAddress` | Sender email address |
| `0042` | `001F` | `PidTagSentRepresentingName` | "Sent on behalf of" display name |
| `0E04` | `001F` | `PidTagDisplayTo` | Formatted `To:` recipient list |

### 2.2 Fixed-Length Property Stream (`__properties_version1.0`)
Contains 16-byte fixed-size binary records:
* **Header Size:** 32 bytes for the root message; 24 bytes for sub-storages.
* **Structure:** `<PropTag: 4 bytes><Flags: 4 bytes><Value: 8 bytes>`
* **Fixed Properties Extracted:**
  * `0x0E070003` (`PidTagMessageFlags`): Bitmask `0x01 = MSGFLAG_READ`, `0x02 = MSGFLAG_UNMODIFIED`, `0x10 = MSGFLAG_HASATTACH`.
  * `0x10900003` (`PidTagFlagStatus`): Follow-up flag (`2 = olFlagMarked`).
  * `0x10810003` (`PidTagLastVerbExecuted`): `102 = Replied`, `104 = Forwarded`.
  * `0x00390040` (`PidTagClientSubmitTime`) & `0x0E060040` (`PidTagMessageDeliveryTime`): 64-bit Windows `FILETIME` (100-nanosecond intervals since January 1, 1601 UTC).

### 2.3 Recipient & Attachment Sub-Storages
* **Recipients (`__recip_version1.0_#XXXXXXXX`):** Contains `PidTagRecipientType` (`1 = To`, `2 = Cc`, `3 = Bcc`), `PidTagDisplayName` (`0x3001001F`), and `PidTagSmtpAddress` (`0x39FE001F`).
* **Attachments (`__attach_version1.0_#XXXXXXXX`):** Contains `PidTagAttachDataBinary` (`0x37010102`), filename streams (`0x3707001F`), MIME type (`0x370E001F`), and `PidTagAttachContentId` (`0x3712001F`) for inline `<img src="cid:...">` rendering.

---

## 3. MAPI DASL Property Reference Table

When populating MAPI stores (such as `.pst` or live Exchange via COM/DASL), the engine sets the following property tags:

```text
http://schemas.microsoft.com/mapi/proptag/0x<TagHex>
```

| DASL Tag | Type | Canonical Name | Target Field |
| :--- | :--- | :--- | :--- |
| `0x007D001F` | `PT_UNICODE` | `PidTagTransportMessageHeaders` | Original Internet MIME headers |
| `0x1035001F` | `PT_UNICODE` | `PidTagInternetMessageId` | RFC 5322 Message-ID |
| `0x0C1A001F` | `PT_UNICODE` | `PidTagSenderName` | Sender display name |
| `0x0C1F001F` | `PT_UNICODE` | `PidTagSenderEmailAddress` | Sender address |
| `0x0C1E001F` | `PT_UNICODE` | `PidTagSenderAddressType` | `"SMTP"` |
| `0x00390040` | `PT_SYSTIME` | `PidTagClientSubmitTime` | Client sent time (`FILETIME`) |
| `0x0E060040` | `PT_SYSTIME` | `PidTagMessageDeliveryTime` | Delivery time (`FILETIME`) |
| `0x3712001F` | `PT_UNICODE` | `PidTagAttachContentId` | Content-ID (`cid:...`) |
| `0x37140003` | `PT_LONG` | `PidTagAttachFlags` | `4` (`ATT_MHTML_REF` for inline images) |
| `0x10800003` | `PT_LONG` | `PidTagIconIndex` | `261` (Replied icon), `262` (Forwarded icon) |
