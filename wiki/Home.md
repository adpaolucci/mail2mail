# Mailbox Migration & Protocol Engineering Wiki

Welcome to the **Mailbox Migration & Protocol Engineering Wiki**. This knowledge base provides comprehensive documentation on email protocols, server-side throttling behaviors, public provider rate limits, dedicated export endpoints, and optimal migration strategies for tools like [`mail2rust`](https://github.com/adpaolucci/mail2rust) and [`mail2mail`](https://github.com/adpaolucci/mail2mail).

---

## 📚 Table of Contents

### 1. Provider Restrictions & Endpoints
* **[[IMAP-Rate-Limits-and-Provider-Quotas]]**: Exhaustive reference of concurrent socket caps, bandwidth limits, daily quotas, and error signatures across Google, Microsoft, Yahoo, Apple, Zoho, Fastmail, and Proton.
* **[[Yahoo-and-Rogers-Export-Endpoints]]**: Deep dive into Yahoo Mail and Rogers Yahoo Mail, the 10,000-message per folder IMAP restriction, and how to use the dedicated `export.imap.mail.yahoo.com` endpoint.
* **[[OAuth2-and-App-Specific-Passwords]]**: Authentication requirements, token exchange, and generation of App Passwords for legacy IMAP access post-Basic Auth deprecation.

### 2. Modern Protocols & Alternative Pathways
* **[[Alternative-Export-Pathways]]**: Bypassing IMAP bottlenecks completely via modern APIs (JMAP, Microsoft Graph, Gmail REST), official bulk archives (Google Takeout, Apple Privacy, Yahoo Privacy), and local client caches (`.ost`, `.emlx`).
* **[[JMAP-Protocol-Specification]]**: Overview of RFC 8620 (JMAP Core) and RFC 8621 (JMAP Mail) and why JMAP is vastly superior to IMAP for high-volume exports.

### 3. Engineering & Client Tuning
* **[[Migration-Best-Practices-and-Tuning]]**: Algorithmic mitigation strategies, adaptive exponential backoff, concurrency clamping, UID range paging, and non-destructive inspection (`BODY.PEEK[]`).
* **[[Mailbox-Format-Conversion-Matrix]]**: Format conversion, structure validation, and mapping between PST, OST, MBOX, EML, EMLX, MSG, Maildir, and JMAP.

---

## ⚡ Quick Reference: Provider Limits at a Glance

| Provider | Standard IMAP Host | Concurrency Limit | Daily / Folder Quota | Dedicated Export / Alternative Pathway |
| :--- | :--- | :--- | :--- | :--- |
| **Google Gmail** | `imap.gmail.com:993` | 15 sockets | **2.5 GB / day** download limit | **Google Takeout** (`takeout.google.com`) or **Gmail REST API** |
| **Microsoft 365** | `outlook.office365.com:993` | ~20 sockets / mailbox | Dynamic throttling (~15–20 GB/day) | **Microsoft Graph API** or **Desktop Outlook `.ost` cache** |
| **Yahoo / Rogers** | `imap.mail.yahoo.com:993` | **5 sockets** | **10,000 messages / folder** ceiling | **`export.imap.mail.yahoo.com:993`** (concurrency=1) |
| **Apple iCloud** | `imap.mail.me.com:993` | ~2 sockets safely | Heuristic bandwidth throttle | **Apple Privacy Export** or **macOS `~/Library/Mail/V10/` (`.emlx`)** |
| **Fastmail** | `imap.fastmail.com:993` | Adaptive | 500 logins / 10 minutes | **JMAP API** (`https://api.fastmail.com/jmap/api/`) |
| **Proton Mail** | `127.0.0.1:1143` *(Bridge)* | Unlimited (local) | Local GPG decryption CPU bound | **Proton Mail Export Tool** or Local Bridge with concurrency $\le 4$ |

---

## 🛠️ Deploying this Wiki to GitHub
To publish this documentation to your repository's official GitHub Wiki:
1. Navigate to your GitHub repository and click the **Wiki** tab.
2. Click **Create the first page** and save a dummy page to initialize the Git wiki repository.
3. Clone the wiki repository locally:
   ```bash
   git clone https://github.com/<owner>/<repo>.wiki.git
   ```
4. Copy all `.md` files from this `wiki/` directory into that clone and push:
   ```bash
   cp wiki/*.md <repo>.wiki/
   cd <repo>.wiki
   git add .
   git commit -m "Publish Mailbox Migration & Protocol Engineering Wiki"
   git push origin master
   ```
