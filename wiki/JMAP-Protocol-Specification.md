# JMAP Protocol (RFC 8620 & RFC 8621) Reference

The **JSON Meta Application Protocol (JMAP)** is an open standard designed by the IETF to modernize and replace IMAP.

* **RFC 8620:** The JSON Meta Application Protocol (Core Architecture)
* **RFC 8621:** Using JMAP for Mail

For data extraction and high-speed migrations, JMAP eliminates almost all of IMAP's foundational flaws.

---

## 1. Why JMAP is Superior to IMAP for Data Migration

| Capability | Legacy IMAP (RFC 3501) | JMAP (RFC 8620 / 8621) |
| :--- | :--- | :--- |
| **Transport** | Stateful, long-lived TCP socket | **Stateless HTTPS (HTTP/2 or HTTP/3)** |
| **Payload Format** | Custom line-oriented text + binary literals | **Standard JSON over REST** |
| **Connection Drops** | Transient socket resets abort operations | **Stateless requests auto-retry safely** |
| **Roundtrips** | Chatty; requires separate capability, login, select, search, fetch calls | **Single roundtrip can query, batch-fetch, and resolve** |
| **Batch Retrieval** | Unwieldy UID ranges | **Structured batching (`Email/get` with 500 IDs)** |
| **Header Parsing** | Raw MIME parsing required on client | **Server parses headers, body text, HTML, and attachments** |

---

## 2. JMAP Session Discovery & Authentication

JMAP endpoints advertise their capabilities through a standard discovery path:

```http
GET /.well-known/jmap HTTP/1.1
Host: api.fastmail.com
Authorization: Bearer <API-TOKEN>
```

### Response Example:
```json
{
  "username": "user@fastmail.com",
  "apiUrl": "https://api.fastmail.com/jmap/api/",
  "downloadUrl": "https://api.fastmail.com/jmap/download/{accountId}/{blobId}/{name}?accept={type}",
  "uploadUrl": "https://api.fastmail.com/jmap/upload/{accountId}/",
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
    "urn:ietf:params:jmap:mail": {}
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
  }
}
```

---

## 3. High-Volume Email Extraction via `Email/query` and `Email/get`

Extracting emails over JMAP involves two core method calls:

### Step 1: Querying Message IDs (`Email/query`)
```json
{
  "using": ["urn:ietf:params:jmap:core", "urn:ietf:params:jmap:mail"],
  "methodCalls": [
    [
      "Email/query",
      {
        "accountId": "u12345678",
        "filter": {
          "inMailbox": "mbx-inbox-id"
        },
        "sort": [{ "property": "receivedAt", "isAscending": false }],
        "position": 0,
        "limit": 500
      },
      "call-query-1"
    ]
  ]
}
```

### Step 2: Batch Fetching Full Messages (`Email/get`)
```json
{
  "using": ["urn:ietf:params:jmap:core", "urn:ietf:params:jmap:mail"],
  "methodCalls": [
    [
      "Email/get",
      {
        "accountId": "u12345678",
        "ids": ["M1001", "M1002", "M1003"],
        "properties": [
          "id", "blobId", "threadId", "mailboxIds", 
          "from", "to", "subject", "receivedAt", 
          "textBody", "htmlBody", "attachments", "size"
        ]
      },
      "call-get-1"
    ]
  ]
}
```

---

## 4. JMAP Implementation in `mail2mail`

The repository includes a production-ready, pure-Python JMAP client (`mail2jmap.py`):

```bash
# Export entire Fastmail account to structured JMAP JSON
python mail2mail.py jmap \
  --token "fmu1-xxxx-xxxx-xxxx" \
  --out fastmail_dump.jmap.json

# Convert JMAP JSON into a standard Outlook PST
python mail2mail.py convert fastmail_dump.jmap.json fastmail.pst
```
