# OAuth2 & App-Specific Passwords Guide

Across modern email infrastructure, **Basic Authentication (plain username and account password) has been universally deprecated** to protect against credential stuffing, phishing, and password reuse attacks.

To connect migration software (`mail2rust`, `mail2mail`, Thunderbird, Outlook) to modern public email services, you must use either **OAuth 2.0 (XOAUTH2 / Modern Auth)** or a generated **App-Specific Password**.

---

## 1. Authentication Requirements by Provider

| Provider | Basic Auth Status | Primary Authentication Method | Legacy Client Fallback |
| :--- | :--- | :--- | :--- |
| **Google Gmail / Workspace** | **Disabled** (May 2022) | **OAuth 2.0 (XOAUTH2)** | **App Password** (Requires 2FA) |
| **Microsoft 365 / Outlook.com** | **Disabled** (Jan 2023) | **Modern Auth (MSAL / OAuth2)** | App Password (Tenant policy dependent) |
| **Yahoo Mail & AOL** | **Disabled** | **OAuth 2.0 (XOAUTH2)** | **App Password** (Mandatory for IMAP) |
| **Rogers Yahoo Mail** | **Disabled** | Yahoo SSO / OAuth2 | **App Password** (Via Rogers Member Centre) |
| **Apple iCloud Mail** | **Disabled** | Apple ID 2FA | **App-Specific Password** (Mandatory for IMAP) |
| **Fastmail** | Permitted (Restricted) | **API Tokens** (JMAP/IMAP) | Custom App Password with scoped access |
| **Proton Mail** | N/A (Local Bridge) | Local Bridge generated password | Bridge local password + STARTTLS |

---

## 2. Setting Up App-Specific Passwords

### 2.1 Google Accounts (Gmail & Google Workspace)
> [!NOTE]
> App Passwords on Google require **2-Step Verification (2FA)** to be enabled first.

1. Navigate to your [Google Account Security Settings](https://myaccount.google.com/security).
2. Under **"How you sign in to Google"**, verify that **2-Step Verification** is turned ON.
3. Search for **"App passwords"** in the top search bar, or visit [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords).
4. Enter an application name (e.g., `Mail2Rust Migration`).
5. Click **Create**. Google will display a 16-character code (e.g., `abcd efgh ijkl mnop`).
6. Use this 16-character string (without spaces) as your IMAP password.

---

### 2.2 Yahoo Mail & AOL Mail
1. Log into your [Yahoo Account Security](https://login.yahoo.com/account/security) page.
2. Scroll to the bottom and click **"Generate app password"** or **"Manage app passwords"**.
3. Click **Get started**, enter an app name (e.g., `Mail Backup`), and click **Generate password**.
4. Copy the single-use 16-character password and provide it to the migration tool.

---

### 2.3 Rogers Yahoo Mail (`@rogers.com`)
Because Rogers uses Yahoo's backend with carrier single sign-on (SSO), app passwords must be generated via the Rogers Member Centre:
1. Log in to [mail.rogers.com](https://mail.rogers.com) or [rogersmembercentre.com](https://www.rogersmembercentre.com/).
2. Navigate to **Account Settings** $\rightarrow$ **Account Security**.
3. Under the **External Connections** or **App Passwords** section, click **Generate App Password**.
4. Enter an identifier (e.g., `Migration CLI`) and copy the code.
5. Use your full email address (`username@rogers.com`) with this code on `export.imap.mail.yahoo.com:993`.

---

### 2.4 Apple iCloud Mail (`@icloud.com`, `@me.com`, `@mac.com`)
1. Sign in to your [Apple ID Management Portal](https://appleid.apple.com/).
2. Under the **Sign-In and Security** section, select **App-Specific Passwords**.
3. Click **Generate an app-specific password** or the **+** icon.
4. Label the password (e.g., `IMAP Export`) and authenticate with your Apple device.
5. Copy the generated code (e.g., `xxxx-xxxx-xxxx-xxxx`) and enter it into the tool.

---

### 2.5 Fastmail API Tokens
1. In Fastmail, go to **Settings** $\rightarrow$ **Password & Security** $\rightarrow$ **Integrations**.
2. Click **New API Token**.
3. Grant permissions:
   * For IMAP: Check **Mail (read-only)** or **Mail (full access)**.
   * For JMAP: Check **JMAP Mail (read/write)**.
4. Copy the token string (prefixed with `fmu1-...`).

---

## 3. Implementing OAuth 2.0 (XOAUTH2) in Code

For tools supporting native OAuth2 token authentication, the client sends an initial SASL `AUTHENTICATE XOAUTH2` command formatted according to Google and Microsoft specifications:

### SASL XOAUTH2 String Construction (Base64)
$$\text{Payload} = \text{base64}(\text{"user="} + \text{email} + \text{"\x01auth=Bearer "} + \text{accessToken} + \text{"\x01\x01"})$$

### Python Implementation Example
```python
import base64
import imaplib

def generate_xoauth2_string(username: str, access_token: str) -> str:
    auth_string = f"user={username}\x01auth=Bearer {access_token}\x01\x01"
    return base64.b64encode(auth_string.encode('ascii')).decode('ascii')

def connect_xoauth2(host: str, port: int, username: str, access_token: str):
    imap = imaplib.IMAP4_SSL(host, port)
    xoauth2_str = generate_xoauth2_string(username, access_token)
    imap.authenticate('XOAUTH2', lambda x: xoauth2_str)
    return imap
```

### Rust Implementation Example
```rust
use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64;

pub fn build_xoauth2_token(user: &str, access_token: &str) -> String {
    let raw = format!("user={}\x01auth=Bearer {}\x01\x01", user, access_token);
    BASE64.encode(raw.as_bytes())
}
```
