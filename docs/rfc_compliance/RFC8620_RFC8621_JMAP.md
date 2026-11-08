# RFC 8620 & RFC 8621: JMAP (JSON Meta Application Protocol) & JMAP for Mail Compliance

This document specifies how `mail2mail` implements and complies with **RFC 8620** (The JSON Meta Application Protocol / JMAP Core) and **RFC 8621** (JMAP for Mail) across [`mail2jmap.py`](../../mail2jmap.py), [`mail_formats.py`](../../mail_formats.py), [`mail2eml.py`](../../mail2eml.py), [`mail2mbox.py`](../../mail2mbox.py), [`mail2pst.py`](../../mail2pst.py), and [`mail2mail.py`](../../mail2mail.py).

---

## 1. Abstract & Architectural Background

**JMAP** is an open Internet standard developed by **Fastmail** and standardized by the IETF (Internet Engineering Task Force) to modernize and replace legacy email protocols:
- **RFC 8620 (JMAP Core)**: Defines the transport agnostic, JSON-over-HTTP RPC communication architecture, session discovery, binary object handling, and batch method invocation.
- **RFC 8621 (JMAP for Mail)**: Defines the data models and operations for modern email mailboxes, messages, threads, and search queries.

### Why JMAP Replaces Legacy IMAP (RFC 3501)
Traditional IMAP suffers from fundamental design limitations rooted in 1980s network architectures:
1. **Stateful Connection Drops**: IMAP requires maintaining stateful TCP connections with sequential line-oriented synchronization (`SELECT`, `UID SEARCH`, `UID FETCH`), leading to broken transfers, socket timeouts, and complex reconnection handshakes over modern mobile or WAN networks.
2. **Parsing Ambiguity**: IMAP responses use irregular, nested parentheses-delimited S-expressions with mixed literals (`{123}\r\n...`), necessitating complex state-machine parsers susceptible to buffer overflow or parser confusion.
3. **Redundant MIME Payload Overhead**: Fetching message bodies in IMAP frequently requires downloading entire multipart RFC 5322 MIME payloads including base64 wrappers and boundary markers, wasting significant bandwidth.
4. **Chatty Roundtrips**: IMAP requires multiple roundtrips to list mailboxes, inspect counts, query message ranges, and fetch headers.

**JMAP solves these challenges** by providing:
- **Stateless HTTP/HTTPS Transport**: Standard HTTPS requests with Bearer Token authentication (RFC 6750), resilient to connection drops.
- **Structured JSON Representations**: Pure RFC 8259 JSON objects with clean typing.
- **Client-Side Property Filtering**: Allows fetching only the required properties (e.g. `subject`, `receivedAt`, `bodyValues`) in a single roundtrip.
- **Atomic Batch Ingestion**: Up to 500 emails queried and fetched in a single request with backreferenced method calls.
- **State-Token Delta Syncing**: High-watermark synchronization using opaque server state tokens (`state` on `Mailbox/get` and `Email/changes`).

---

## 2. RFC 8620 — JMAP Core Protocol Implementation

### 2.1 JMAP Session Discovery (§2)
To communicate with a JMAP server, clients query the standard Session URL via `GET` with an `Authorization: Bearer <token>` header:
- **Fastmail Session URL**: `https://api.fastmail.com/jmap/session`
- **Standard Local/Enterprise Endpoints**: `https://<mail-host>/.well-known/jmap` or `https://<mail-host>/jmap/session`

The session response provides:
```json
{
  "capabilities": {
    "urn:ietf:params:jmap:core": {
      "maxSizeUpload": 50000000,
      "maxConcurrentUpload": 4,
      "maxSizeRequest": 10000000,
      "maxConcurrentRequests": 8,
      "maxCallsInRequest": 16,
      "maxObjectsInGet": 500,
      "maxObjectsInSet": 500
    },
    "urn:ietf:params:jmap:mail": {
      "maxMailboxesPerEmail": 50,
      "maxMailboxDepth": null,
      "maxSizeMailboxName": 255,
      "totalQuota": 53687091200
    }
  },
  "accounts": {
    "u12345678": {
      "name": "user@fastmail.com",
      "isPersonal": true,
      "isReadOnly": false,
      "accountCapabilities": {
        "urn:ietf:params:jmap:mail": {}
      }
    }
  },
  "primaryAccounts": {
    "urn:ietf:params:jmap:mail": "u12345678"
  },
  "apiUrl": "https://api.fastmail.com/jmap/api/",
  "downloadUrl": "https://api.fastmail.com/jmap/download/{accountId}/{blobId}/{name}",
  "uploadUrl": "https://api.fastmail.com/jmap/upload/{accountId}/"
}
```

`mail2mail`'s [`JmapClient.connect()`](../../mail2jmap.py) performs HTTP discovery, verifies mail capability (`urn:ietf:params:jmap:mail`), and resolves the primary mail `accountId`.

### 2.2 Method Invocations & Batching (§3.1, §3.2)
All JMAP API operations are sent as HTTP `POST` requests to `apiUrl` with `Content-Type: application/json`:
```json
{
  "using": [
    "urn:ietf:params:jmap:core",
    "urn:ietf:params:jmap:mail"
  ],
  "methodCalls": [
    ["Mailbox/get", { "accountId": "u12345678", "ids": null }, "m0"],
    ["Email/query", { "accountId": "u12345678", "filter": { "inMailbox": "mb_inbox" }, "limit": 200 }, "q0"]
  ]
}
```
The server returns a structured `methodResponses` array with matching tag IDs (`m0`, `q0`).

---

## 3. RFC 8621 — JMAP for Mail Implementation

### 3.1 Mailbox Object Model & Discovery (§2)
A JMAP `Mailbox` represents a folder in the user's mail store:
- `id`: Opaque immutable identifier.
- `name`: Display name of the folder (e.g. `Inbox`, `Archives/2026`).
- `parentId`: Parent mailbox ID, or `null` for top-level root folders.
- `role`: Standardized system role (`inbox`, `drafts`, `sent`, `trash`, `archive`, `junk`, etc.).
- `totalEmails` & `unreadEmails`: Message count telemetry.

`mail2mail` resolves full folder path hierarchies by recursively looking up `parentId` references, ensuring nested folder trees (e.g. `Projects/Titan/Q4`) are mirrored faithfully into Outlook PST, MBOX, and EML directory structures.

### 3.2 Email Object Model (§4)
A JMAP `Email` entity represents a message:
- **Identifiers**: `id` (JMAP entity ID), `blobId` (raw binary content handle), `threadId` (conversation grouping identifier).
- **Location**: `mailboxIds` (dictionary of `mailboxId: true` representing mailboxes containing the message).
- **Keywords / Status Flags**: RFC 8621 §4.1.1 system keywords and custom tags:
  - `$seen`: Read status.
  - `$flagged`: Starred / marked for followup.
  - `$answered`: Replied to.
  - `$draft`: Draft state.
  - `$deleted`: Marked for deletion / expunged.
- **Timestamps**:
  - `receivedAt`: ISO 8601 UTC timestamp when delivered to the server (corresponds to IMAP `INTERNALDATE`).
  - `sentAt`: ISO 8601 UTC timestamp from the RFC 5322 `Date` header.
- **Header Properties**:
  - `messageId`: Array of RFC 5322 `Message-ID` strings.
  - `inReplyTo`: Array of parent message IDs.
  - `references`: Array of ancestor conversation message IDs.
  - `from`, `to`, `cc`, `bcc`, `replyTo`: Arrays of `EmailAddress` objects (`{"name": "...", "email": "..."}`).
  - `subject`: Unfolded UTF-8 subject line.
- **Body Values & Attachments**:
  - `textBody`: Array of `EmailBodyPart` objects describing plain text parts.
  - `htmlBody`: Array of `EmailBodyPart` objects describing HTML parts.
  - `bodyValues`: Map of `partId` to `EmailBodyValue` objects containing `value` (unfolded UTF-8 text/HTML content), `isTruncated`, and `isEncodingProblem`.
  - `attachments`: Array of `EmailBodyPart` metadata (`blobId`, `name`, `type`, `size`, `cid`).

---

## 4. Master JMAP Mapping Reference

### 4.1 JMAP to RFC 5322 MIME & MAPI Properties

| JMAP Property (RFC 8621) | RFC 5322 Header / MIME Representation | Outlook MAPI Property (`mail2pst`) |
| :--- | :--- | :--- |
| `id` | `X-JMAP-ID: <id>` | Tagged in transport headers |
| `threadId` | `Conversation-ID` | `PR_CONVERSATION_TOPIC`, `PR_CONVERSATION_INDEX` |
| `messageId[0]` | `Message-ID: <...>` | `PR_INTERNET_MESSAGE_ID_W` (`0x1035001F`) |
| `inReplyTo[0]` | `In-Reply-To: <...>` | `PR_IN_REPLY_TO_ID_W` (`0x1042001F`) |
| `references` | `References: <...> <...>` | Transport headers preservation |
| `subject` | `Subject: ...` | `PR_SUBJECT_W` (`0x0037001F`) |
| `from[0]` | `From: Name <email@domain>` | `PR_SENDER_NAME_W`, `PR_SENDER_EMAIL_ADDRESS_W` |
| `to` | `To: a@b, c@d` | `mail_item.To`, `PR_DISPLAY_TO_W` |
| `cc` | `Cc: e@f` | `mail_item.CC`, `PR_DISPLAY_CC_W` |
| `bcc` | `Bcc: g@h` | `mail_item.BCC`, `PR_DISPLAY_BCC_W` |
| `receivedAt` | Delivery envelope / `Received:` | `PR_MESSAGE_DELIVERY_TIME` (UTC `VT_DATE`) |
| `sentAt` | `Date: ...` (RFC 5322 §3.3) | `PR_CLIENT_SUBMIT_TIME` (UTC `VT_DATE`) |
| `keywords.$seen` | IMAP `\Seen` / `Status: RO` | `MSG_FLAG_READ` (`0x0001`) on `PR_MESSAGE_FLAGS` |
| `keywords.$flagged` | IMAP `\Flagged` / `X-Status: F` | `MSG_FLAG_MARKED` (`0x0004`), follow-up flag |
| `keywords.$answered` | IMAP `\Answered` / `Status: A` | `MSG_FLAG_REPLIED` (`0x0002`), icon index |
| `keywords.$draft` | IMAP `\Draft` | `MSG_FLAG_DRAFT` (`0x0008`) |
| `bodyValues[pid].value` (`textBody`) | Inline `text/plain; charset=utf-8` | `mail_item.Body`, `PR_BODY_W` (`0x1000001F`) |
| `bodyValues[pid].value` (`htmlBody`) | Inline `text/html; charset=utf-8` | `mail_item.HTMLBody`, `PR_HTML` (`0x10130102`) |
| `attachments` | MIME `multipart/mixed` attachments | Outlook `Attachments.Add()`, `PR_ATTACH_FILENAME_W` |

---

## 5. File Formats & Static Ingestion

In addition to live HTTPS synchronization via `mail2jmap.py`, `mail2mail` natively ingests static JMAP exports and offline backup archives:
1. **`.jmap` and `.jmap.json` Files**: Single JSON files or arrays containing JMAP `Email` entities.
2. **Batch RPC Responses**: Raw JSON dumps of JMAP `Email/get` responses (`{"methodResponses": [["Email/get", {"list": [...]}, "tag"]]}`).
3. **Directory Collections**: Folders of `.jmap` files organized by mailbox names.
4. **Compressed Archives**: `.jmap.gz`, `.jmap.zip`, or `.tar.gz` containing JMAP entities.

The [`iter_jmap_records()`](../../mail_formats.py) reader sniffs JMAP JSON structure automatically and streams standardized `MailRecord` instances with zero temporary disk footprint.

---

## 6. Automated Test & Verification Matrix

| Verification Scope | Test File | Implemented Test |
| :--- | :--- | :--- |
| **RFC 8620 Session Discovery** | [`tests/test_jmap.py`](../../tests/test_jmap.py) | `test_jmap_session_discovery_and_account_resolution` |
| **RFC 8621 Mailbox Listing** | [`tests/test_jmap.py`](../../tests/test_jmap.py) | `test_jmap_mailbox_get_hierarchy` |
| **RFC 8621 Email Query & Batch Get** | [`tests/test_jmap.py`](../../tests/test_jmap.py) | `test_jmap_email_query_and_batch_get` |
| **Keywords & Status Flag Mapping** | [`tests/test_jmap.py`](../../tests/test_jmap.py) | `test_jmap_keyword_to_status_flags_mapping` |
| **Static JMAP JSON File Sniffing** | [`tests/test_jmap.py`](../../tests/test_jmap.py) | `test_jmap_file_format_sniffing_and_ingestion` |
| **End-to-End Live Sync to PST/MBOX/EML** | [`tests/test_jmap.py`](../../tests/test_jmap.py) | `test_jmap_mock_sync_end_to_end_to_all_formats` |
| **Resumable Delta Sync Checkpoint** | [`tests/test_jmap.py`](../../tests/test_jmap.py) | `test_jmap_checkpoint_delta_sync` |
