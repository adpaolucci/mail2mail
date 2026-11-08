# IMAP Rate Limits & Provider Quotas

Standard Internet Message Access Protocol (IMAP4rev1, RFC 3501) was designed for lightweight, interactive message header indexing and sporadic message fetching. When utilized for bulk mailbox migrations, automated backups, or forensic extraction, public mail service providers enforce strict rate limits, bandwidth quotas, and connection constraints to protect their infrastructure.

This guide provides technical specifications for throttling mechanisms and error signatures across major public email providers.

---

## 1. Provider Technical Limits Matrix

| Provider | Standard IMAP Endpoint | Port / Security | Concurrency Cap | Bandwidth / Daily Limits | Lockout Duration |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Google Gmail & Workspace** | `imap.gmail.com` | `993` SSL/TLS | 15 concurrent sockets | **2,500 MB (2.5 GB) / day download**<br>500 MB / day upload | **1 to 24 hours** |
| **Microsoft 365 / Outlook.com** | `outlook.office365.com` | `993` SSL/TLS | ~20 connections / mailbox | Dynamic throttling policy (~15–20 GB/day soft cap; max 100k items/folder) | Dynamic (minutes to hours) |
| **Yahoo Mail & AOL** | `imap.mail.yahoo.com` | `993` SSL/TLS | **5 concurrent connections** | **10,000 items / folder ceiling**;<br>Aggressive burst command pacing | Immediate socket teardown |
| **Apple iCloud Mail** | `imap.mail.me.com` | `993` SSL/TLS | ~2–3 connections safely | Heuristic bandwidth throttle (undocumented MB limits) | Temporary socket backoff |
| **Zoho Mail** | `imap.zoho.com` | `993` SSL/TLS | 100 connections max | Connection churn limit; POP capped at 5 GB/day | 1 hour temporary block |
| **Fastmail** | `imap.fastmail.com` | `993` SSL/TLS | Adaptive | 500 successful logins / 10 minutes | Dynamic throttle |
| **Proton Mail** | `127.0.0.1` *(Bridge)* | `1143` STARTTLS | Unlimited (Local TCP) | Local CPU / GPG decryption throughput | None (local crash if OOM) |

---

## 2. Deep Dive: Provider Throttling Behaviors

### 2.1 Google Gmail & Google Workspace
* **Daily Bandwidth Cliff:** Gmail enforces a rigid **2.5 GB per 24-hour rolling window** for IMAP downloads. A mailbox of 50 GB will take a minimum of **20 consecutive days** to export over standard IMAP without triggering lockouts.
* **Throttling Trigger:** The bandwidth cap is shared across all IMAP clients and all folders (including `[Gmail]/All Mail`).
* **Error Signature:**
  ```text
  NO [ALERT] Account exceeded bandwidth limits. (Failure)
  ```
  Once triggered, the server terminates active sessions and refuses new authentication handshakes for 1 to 24 hours.
* **Folder Structure Pitfall:** Gmail's label architecture means messages with multiple labels appear under multiple IMAP folders (`Inbox`, `Important`, `Work`, `[Gmail]/All Mail`). Syncing every folder naively downloads the same message bytes multiple times, depleting the 2.5 GB daily quota 3x–5x faster.

---

### 2.2 Microsoft 365 / Exchange Online / Outlook.com
* **Tenant-Level Throttling:** Microsoft 365 enforces policy-based throttling (`EWSMaxConcurrency`, `IMAPMaxConcurrency`) at the mailbox and tenant layer.
* **Large Folder Freezes:** Folders containing more than 100,000 items experience severe latency on `SELECT`, `EXAMINE`, and `SEARCH` commands.
* **Error Signatures:**
  ```text
  * BYE Connection closed.
  NO Server Unavailable. 15
  NO Maximum concurrent IMAP connections exceeded.
  ```
* **Connection Drop Behavior:** Unlike Gmail which emits a structured `[ALERT]` before locking, Exchange Online frequently resets TCP sockets unceremoniously when concurrency thresholds are exceeded.

---

### 2.3 Yahoo Mail & AOL Mail
* **The 10,000-Message Folder Wall:** On `imap.mail.yahoo.com`, folders containing more than 10,000 messages cannot be fully synced. The server caps folder listings and UID responses at 10,000 items. Requests for higher ranges return empty responses or errors.
* **Strict Concurrency Limit:** Yahoo allows at most **5 concurrent connections** per account. Attempting to open 6+ connections immediately causes the server to issue a `BYE` logoff command across sessions.
* **Error Signatures:**
  ```text
  * BYE IMAP4rev1 Server logging out
  NO [LIMIT] CAPABILITY Rate limit hit. Please try again later.
  NO [UNAVAILABLE] Service temporarily unavailable.
  ```
* **Dedicated Solution:** See [[Yahoo-and-Rogers-Export-Endpoints]] for the `export.imap.mail.yahoo.com` migration gateway.

---

### 2.4 Apple iCloud Mail
* **Heuristic Traffic Profiling:** Apple does not publish fixed MB quotas, but monitors traffic patterns for rapid, continuous full-body downloads (`FETCH RFC822`).
* **Session Drops:** When suspicious bulk activity is detected, iCloud begins returning generic transient failures:
  ```text
  NO Server Busy. Please try again later.
  * BYE 421 Connection timeout / rejected.
  ```
* **Security Revocation:** Extreme bulk fetching over standard IMAP can trigger Apple's automated fraud detection, revoking the user's generated App-Specific Password mid-transfer.

---

### 2.5 Zoho Mail
* **Connection Churn Rate:** While Zoho permits up to 100 concurrent sockets on high-tier business plans, rapidly opening and closing connections to poll different folders triggers abuse guards.
* **Error Signature:**
  ```text
  NO Your account is temporarily blocked for IMAP use. This may happen if you exceed the maximum number of simultaneous IMAP connections allowed.
  ```

---

## 3. Detecting and Handling IMAP Error Codes in Client Software

When engineering migration clients (such as `mail2rust` or `mail2mail`), software should inspect server responses and execute protocol-appropriate remediation:

| Server Response Substring | Classification | Action Required |
| :--- | :--- | :--- |
| `[ALERT] Account exceeded bandwidth limits` | Hard Quota Breach | Halt migration immediately. Do not retry for at least 6 hours (retrying resets the 24h timer). Switch to Google Takeout or REST API. |
| `NO [LIMIT]` / `Server Busy` | Transient Rate Throttling | Apply exponential backoff with jitter: sleep $2^n + \text{rand}(1, 5)$ seconds. Reduce batch size to 25. |
| `BYE Connection closed` / `421` | Socket Reset / Concurrency Drop | Reduce max concurrent worker threads to 1 or 2. Re-authenticate after a 10-second cool-down. |
| `NO [UNAVAILABLE]` / `Temporary failure` | Backend Load Shedding | Retry item after 5 seconds; log failed UID to retry queue. |

---

## 4. Next Steps
* Learn how to bypass the 10,000-message ceiling on Yahoo and Rogers accounts: [[Yahoo-and-Rogers-Export-Endpoints]]
* Explore modern API and bulk-archive alternatives that bypass IMAP limits: [[Alternative-Export-Pathways]]
* Review client-side tuning and backoff algorithms: [[Migration-Best-Practices-and-Tuning]]
