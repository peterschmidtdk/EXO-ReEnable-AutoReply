# ReEnabled-ExO-AutoReply

Two PowerShell scripts that help you **re-enable (toggle)** Exchange Online Out of Office / Automatic Replies (OOF) for users **who already have OOF enabled**, using **app-only authentication**.

This is useful if you need to “refresh” OOF state (disable → re-enable) without accidentally enabling OOF for users who don’t use it.

---

## What’s included

### 1) `Install-ReEnabled-ExO-AutoReplyAppRegistration.ps1`
Creates the required **App Registration** and configures it for Exchange Online app-only access:

- Creates an **App Registration** + **Service Principal**
- Assigns Exchange Online **Application permission**: `Exchange.ManageAsApp`
- Grants **Admin consent**
- Creates a **self-signed certificate**, exports it as:
  - `.\Cert\ReEnabled-ExO-AutoReply.cer` (public)
  - `.\Cert\ReEnabled-ExO-AutoReply.pfx` (private key)
- Uploads the **public cert** to the App Registration (KeyCredentials)
- (Optional) Adds the Service Principal to an **Exchange Online role group** (default: `Organization Management`)

> ⚠️ `Organization Management` is highly privileged. Consider creating a custom EXO role group with only the roles needed.

---

### 2) `ReEnabled-ExO-AutoReply.ps1`
Connects to Exchange Online (app-only) and **re-enables OOF** safely:

- Supports two target modes:
  - **AllEnabled**: scans all mailboxes and processes only mailboxes where OOF is **Enabled** or **Scheduled**
  - **CsvEnabled**: reads users from a CSV and processes only those users where OOF is **Enabled** or **Scheduled**
- **Never enables OOF** if a mailbox is currently `Disabled`
- “Re-enable” logic:
  1) Snapshot current OOF config
  2) Set OOF to `Disabled`
  3) Restore original state (`Enabled` or `Scheduled`) and messages/schedule

---

## Repo structure (recommended)

```text
ReEnabled-ExO-AutoReply/
│
├─ ReEnabled-ExO-AutoReply.ps1
├─ Install-ReEnabled-ExO-AutoReplyAppRegistration.ps1
├─ README.md
│
├─ Cert/   (created by Install script)
└─ Logs/   (created by Main script)
