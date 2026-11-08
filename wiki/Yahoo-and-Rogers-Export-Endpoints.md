# Yahoo & Rogers Mail Dedicated Export Endpoints

Users and administrators attempting to back up or migrate mailboxes from **Yahoo Mail**, **Rogers Yahoo Mail**, or **AOL Mail** frequently encounter a frustrating barrier: **folders appear truncated at exactly 10,000 emails**. 

This article explains the origin of this limitation, how Yahoo's dedicated migration endpoint bypasses it, and how to configure tools like `mail2rust` and `mail2mail` to safely export large mailboxes from `@yahoo.com` and `@rogers.com` accounts.

---

## 1. The 10,000-Message Limitation on Standard IMAP

On the standard public endpoint:
* **Host:** `imap.mail.yahoo.com`
* **Port:** `993` (SSL/TLS)

Yahoo enforces an intentional architectural constraint: **only the 10,000 most recent messages in any given folder are exposed via IMAP**.
* If an Inbox contains 45,000 messages, standard IMAP clients (such as Microsoft Outlook, Apple Mail, or Thunderbird) will only sync the newest 10,000.
* Older messages are invisible to the client, even though they remain accessible in Yahoo webmail.
* Issuing `UID SEARCH` or `UID FETCH` commands for sequences beyond index 10,000 returns empty datasets or fails with:
  ```text
  NO [LIMIT] CAPABILITY Rate limit hit.
  ```

---

## 2. The Dedicated Export Endpoint: `export.imap.mail.yahoo.com`

To accommodate data export, compliance archiving, and migrations, Yahoo maintains dedicated bulk export endpoints:

| Service / Provider | Standard Sync Endpoint (10k Cap) | Dedicated Export Endpoint (No 10k Cap) | Port & Security |
| :--- | :--- | :--- | :--- |
| **Yahoo Mail** | `imap.mail.yahoo.com` | **`export.imap.mail.yahoo.com`** | Port `993` (SSL/TLS) |
| **Rogers Yahoo Mail** | `imap.mail.yahoo.com` | **`export.imap.mail.yahoo.com`** | Port `993` (SSL/TLS) |
| **AOL Mail** | `imap.aol.com` | **`export.imap.aol.com`** | Port `993` (SSL/TLS) |

### Key Properties of the Export Endpoint:
1. **Lifts the 10,000 Ceiling:** Reports the true message count and full UID space for all folders, allowing sequential download of mailboxes with 100,000+ messages.
2. **Export-Oriented Protocol Handling:** Optimized for sequential `UID FETCH (RFC822.PEEK[])` reading rather than two-way state synchronization.

---

## 3. Rogers Yahoo Mail Integration Details

### Architecture
Rogers Communications does not host its own independent email servers. All `@rogers.com` and `@nl.rogers.com` accounts are white-labeled instances running on **Yahoo's core email infrastructure**.
* Authentication, mailbox storage, indexing, and anti-abuse policies are managed directly by Yahoo.
* Consequently, Rogers accounts are subject to the same 10,000-message folder limit on `imap.mail.yahoo.com`.
* **Rogers accounts can seamlessly use `export.imap.mail.yahoo.com` to bypass this limit.**

### Generating an App Password for Rogers Mail
Standard webmail passwords cannot be used with external IMAP clients. An App Password is required:
1. Go to the **Rogers Member Centre** (`rogersmembercentre.com`) or log in via `mail.rogers.com`.
2. Navigate to **Account Security** $\rightarrow$ **Generate App Password** (or **External Connections**).
3. Select an application name (e.g., `Mail Migration`) and copy the generated 16-character string.
4. Use this string as the IMAP password.

---

## 4. Strict Operational Rules for the Export Endpoint

While `export.imap.mail.yahoo.com` removes the 10k message ceiling, it is guarded by strict anti-scraping and rate-limiting rules:

> [!WARNING]
> Violating these rules will result in immediate socket disconnections (`BYE IMAP4rev1 Server logging out`) and temporary IP blocks.

### Rule 1: Strictly 1 Concurrent Connection (`concurrency = 1`)
Do not run multi-threaded workers against `export.imap.mail.yahoo.com`.
* Standard IMAP allows up to 5 concurrent sockets.
* The `export` endpoint is strictly tuned for single-threaded sequential streaming. Opening $\ge 2$ connections triggers session aborts.

### Rule 2: Discrete UID Paging (50 to 100 Messages per Batch)
Never execute an unbounded fetch (such as `FETCH 1:* (RFC822)`). Query UIDs in batches:
```text
C: A001 UID FETCH 1001:1050 (UID RFC822.SIZE FLAGS)
C: A002 UID FETCH 1001:1025 (BODY.PEEK[])
```

### Rule 3: Command Pacing Delay (100–250 ms)
Introduce a brief programmatic pause between batch queries:
* Sleep **150 ms to 250 ms** after every batch response.
* This ensures requests stay below Yahoo's command burst rate threshold.

---

## 5. Usage Examples with `mail2rust` & `mail2mail`

### Example A: Exporting Rogers Yahoo Mail with `mail2rust`
```bash
mail2rust imap \
  --host export.imap.mail.yahoo.com \
  --port 993 \
  --user "john.doe@rogers.com" \
  --password "abcd-efgh-ijkl-mnop" \
  --concurrency 1 \
  --batch-size 50 \
  --pacing-ms 200 \
  --out rogers_export.pst
```

### Example B: Server-to-Server Transfer (Rogers $\rightarrow$ Fastmail)
```bash
mail2rust imap-sync \
  --src-host export.imap.mail.yahoo.com \
  --src-user "john.doe@rogers.com" \
  --src-pass "abcd-efgh-ijkl-mnop" \
  --src-concurrency 1 \
  --dst-host imap.fastmail.com \
  --dst-user "john.doe@fastmail.com" \
  --dst-pass "fmu1-xxxx-xxxx" \
  --dst-concurrency 4
```

### Example C: Ingestion with `mail2mail`
```bash
python mail2mail.py imap \
  --host export.imap.mail.yahoo.com \
  --user "j_smith@yahoo.com" \
  --password "abcd-efgh-ijkl-mnop" \
  --concurrency 1 \
  --out yahoo_dump.mbox
```

---

## 6. Alternative Workarounds for Yahoo / Rogers Mail

If network policies or transient server issues prevent connecting to `export.imap.mail.yahoo.com`:

1. **Yahoo Privacy Data Download (`privacy.yahoo.com`):**
   * Select **"Request a copy of your data"** $\rightarrow$ **"Yahoo Mail"**.
   * Yahoo compiles a complete, untruncated mailbox archive available for browser download.
2. **Webmail Folder Partitioning:**
   * Create year-based folders in Yahoo webmail (`Inbox_2023`, `Inbox_2022`, etc.) and move emails so each folder holds $<10,000$ messages.
   * Standard IMAP can then fetch all messages without hitting the per-folder ceiling.
