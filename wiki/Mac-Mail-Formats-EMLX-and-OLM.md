# Mac Mail Formats: Apple Mail (.emlx) & Outlook for Mac (.olm)

Forensic extraction and migration on macOS presents unique challenges due to Apple and Microsoft's proprietary macOS formats. 

Both `mail2rust` and `mail2mail` support parsing and extracting messages directly from **Apple Mail (`.emlx`)** and **Outlook for Mac (`.olm`)** archives.

---

## 1. Apple Mail `.emlx` & `.partial.emlx` Format Structure

Apple Mail stores individual email messages in `.emlx` files. An `.emlx` file is a hybrid container composed of three consecutive segments:

```text
+-----------------------------------------------------------+
| Segment 1: Byte Count (ASCII decimal string + newline)    | e.g. "15842\n"
+-----------------------------------------------------------+
| Segment 2: Raw RFC 5322 MIME Message Stream               | Length exactly matches
|            (Headers, text/html body, base64 attachments)  | byte count in Segment 1
+-----------------------------------------------------------+
| Segment 3: Apple XML Property List (plist) Trailer        | Contains flags, read state,
|            <plist version="1.0">...</plist>               | conversation tokens, tags
+-----------------------------------------------------------+
```

### 1.1 Parsing Specification
1. **Segment 1 (Length Prefix):** Read the first line as an ASCII decimal string. Example: `20485\n`. This defines the exact length in bytes of the following MIME message.
2. **Segment 2 (MIME Stream):** Read exactly $N$ bytes starting immediately after the first newline. This slice is a 100% compliant RFC 5322 message stream and can be saved directly as a `.eml` file.
3. **Segment 3 (XML Plist Trailer):** All remaining bytes after $N$ belong to an Apple Property List dictionary. Key plist tags extracted:
   * `flags`: Integer bitmask encoding read/unread (`1 << 0`), flagged (`1 << 4`), replied (`1 << 2`), forwarded (`1 << 8`).
   * `conversation-id`: 64-bit integer tracking Apple Mail thread hierarchies.
   * `read`: Boolean flag indicating `\Seen` state.

### 1.2 Handling `.partial.emlx` Files
When macOS operates in low-storage or optimized-iCloud mode, Apple Mail downloads headers and the first few kilobytes of the body, renaming the file to `.partial.emlx`.
* The parser detects `.partial.emlx`, extracts available header metadata and body previews, and flags the item in the migration audit log as a partial download so users can trigger full offline downloads in Apple Mail if desired.

### 1.3 macOS Mail Directory Layout
```text
~/Library/Mail/V10/
  ├── <UUID-Account-A>/
  │   ├── INBOX.mbox/
  │   │   ├── <UUID>/
  │   │   │   └── Data/
  │   │   │       ├── 0/0/1/Messages/1001.emlx
  │   │   │       └── 0/0/1/Messages/1002.emlx
  │   └── Sent Messages.mbox/
```

---

## 2. Outlook for Mac `.olm` Archive Architecture

Microsoft Outlook for Mac does not use PST or OST files; instead, it exports mailboxes as an `.olm` package. An `.olm` file is standard **ZIP archive** containing XML message descriptors and attachment streams.

### 2.1 Internal Directory Layout
```text
archive.olm (ZIP Archive)
  ├── Accounts/
  │   └── Personal/
  │       └── com.microsoft.__Messages/
  │           └── Inbox/
  │               ├── message_00001.xml
  │               ├── message_00002.xml
  │               └── Subfolder/
  │                   └── message_00003.xml
  ├── Attachments/
  │   └── Att_00001_1_image.png
  └── Categories/
```

### 2.2 XML Message Descriptors
Each `message_XXXXX.xml` describes a complete email. Key elements parsed by `mail2mail` and `mail2rust`:

```xml
<email>
  <OPFMessageCopySubject>Project Roadmap Q1</OPFMessageCopySubject>
  <OPFMessageCopySentTime>2026-11-07T14:30:00Z</OPFMessageCopySentTime>
  <OPFMessageCopyReceivedTime>2026-11-07T14:30:15Z</OPFMessageCopyReceivedTime>
  <OPFMessageCopyMessageID>&lt;roadmap-2026@domain.test&gt;</OPFMessageCopyMessageID>
  <OPFMessageGetIsRead>1</OPFMessageGetIsRead>
  <OPFMessageCopyFromAddresses>
    <emailAddress OPFContactEmailAddressAddress="alice@domain.test" 
                 OPFContactEmailAddressName="Alice Smith"/>
  </OPFMessageCopyFromAddresses>
  <OPFMessageCopyToAddresses>
    <emailAddress OPFContactEmailAddressAddress="bob@domain.test" 
                 OPFContactEmailAddressName="Bob Jones"/>
  </OPFMessageCopyToAddresses>
  <OPFMessageCopyHTMLBody>&lt;html&gt;&lt;body&gt;Roadmap details...&lt;/body&gt;&lt;/html&gt;</OPFMessageCopyHTMLBody>
  <OPFMessageCopyAttachmentList>
    <messageAttachment OPFAttachmentURL="Attachments/Att_00001_1_image.png"
                       OPFAttachmentName="image.png"
                       OPFAttachmentContentType="image/png"
                       OPFAttachmentContentID="logo@domain.test"/>
  </OPFMessageCopyAttachmentList>
</email>
```

### 2.3 Reconstructing RFC 5322 / PST from `.olm`
1. The engine parses the XML tree and decodes entities in `OPFMessageCopyHTMLBody`.
2. Resolves binary attachment streams from the `Attachments/` path inside the ZIP container.
3. Synthesizes standard RFC 5322 MIME or MAPI PST nodes with correct CID references, timestamps, and flags.

---

## 3. Command Line Usage

```bash
# Convert Apple Mail store directly to PST
mail2rust convert ~/Library/Mail/V10/ mac_mail.pst

# Extract Outlook for Mac OLM into standard EML directory tree
mail2rust convert backup.olm ./extracted_emails/

# Ingest OLM via mail2mail
python mail2mail.py convert backup.olm out.mbox
```
