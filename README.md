<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-5391FE?style=for-the-badge&logo=powershell&logoColor=white" alt="PowerShell">
  <img src="https://img.shields.io/badge/Microsoft_Graph-0078D4?style=for-the-badge&logo=microsoft&logoColor=white" alt="Microsoft Graph">
  <img src="https://img.shields.io/badge/Exchange_Online-0078D4?style=for-the-badge&logo=microsoft-outlook&logoColor=white" alt="Exchange Online">
</p>

<h1 align="center">📧 Outlook Rules Manager</h1>

<p align="center">
  <strong>Automate your Outlook inbox organization with server-side rules and folders</strong>
</p>

<p align="center">
  <a href="#-features">Features</a> •
  <a href="#-quick-start">Quick Start</a> •
  <a href="#-demo-testing">Demo Testing</a> •
  <a href="#-configuration">Configuration</a> •
  <a href="#-security">Security</a>
</p>

---

## ✨ Features

| Feature | Description |
|:--------|:------------|
| 📁 **Smart Folders** | Creates organized inbox subfolders (Priority, Action Required, Metrics, Leadership, Alerts, Low Priority) |
| ⚡ **Server-Side Rules** | Rules work across all devices - desktop, mobile, web |
| 🎯 **Priority Handling** | VIP senders get priority treatment with stop-processing logic |
| 🏷️ **Auto-Categorization** | Keyword-based categorization and flagging |
| 🔇 **Noise Filtering** | Automatically archive or delete newsletters and marketing emails |
| 🔄 **Idempotent** | Safe to re-run - won't create duplicates |
| 📋 **Declarative Config** | Rules defined in JSON - version controllable and reviewable |

---

## 📋 Prerequisites

| Requirement | Details |
|:------------|:--------|
| 💻 **PowerShell** | 7+ (recommended) or Windows PowerShell 5.1 |
| 📧 **Microsoft 365** | Account with Exchange Online mailbox |
| 🔐 **Azure AD** | Tenant for app registration (can use personal tenant) |

---

## 🚀 Quick Start

### Step 1: Install Required Modules

```powershell
.\prereqs.ps1
```

### Step 2: Register Azure AD Application (One-Time)

```powershell
.\Register-OutlookRulesApp.ps1
```

> 📝 Creates an app registration with necessary permissions. No admin consent required for personal mailbox access.

### Step 3: Connect to Services

```powershell
# Default: Device code flow (SPACE compliant)
.\Connect-OutlookRulesApp.ps1
```

> 💡 You'll see a code to enter at [microsoft.com/devicelogin](https://microsoft.com/devicelogin)

### Step 4: Run the Setup

```powershell
.\Setup-OutlookRules.ps1
```

---

## 🧪 Demo Testing

Test the solution in a demo/dev tenant before deploying to production.

### Prerequisites for Demo Testing

| Item | Requirement |
|:-----|:------------|
| 🏢 **Demo Tenant** | Azure AD tenant with Exchange Online |
| 👤 **Test User** | User account with mailbox in demo tenant |
| 🔑 **App Registration** | Separate app registration in demo tenant |

### Step-by-Step Demo Setup

#### 1️⃣ Register App in Demo Tenant

```powershell
# Connect to your demo tenant
Connect-AzAccount -TenantId "<your-demo-tenant-id>"

# Register the app
.\Register-OutlookRulesApp.ps1 -AppName "Outlook Rules Manager - Demo"
```

**Or manually in Azure Portal:**

1. Go to **Azure Portal** → **Microsoft Entra ID** → **App registrations**
2. Click **New registration**
3. Configure:
   - **Name**: `Outlook Rules Manager - Demo`
   - **Supported account types**: `Accounts in this organizational directory only`
   - **Redirect URI**: Skip for now
4. Click **Register**

#### 2️⃣ Configure App Authentication

In the app's **Authentication** blade:

| Setting | Value |
|:--------|:------|
| **Platform** | Mobile and desktop applications |
| **Redirect URI** | `https://login.microsoftonline.com/common/oauth2/nativeclient` |
| **Allow public client flows** | ✅ **Yes** |

> ⚠️ **SPACE Compliance**: Do NOT add `http://localhost` - use device code flow instead

#### 3️⃣ Configure API Permissions

In the app's **API permissions** blade:

| API | Permission | Type |
|:----|:-----------|:-----|
| Microsoft Graph | `Mail.ReadWrite` | Delegated |
| Microsoft Graph | `User.Read` | Delegated |

#### 4️⃣ Create Demo Config File

Create `demo.env` in the project directory:

```powershell
# Demo Tenant Config
$ClientId = "<your-demo-app-client-id>"
$TenantId = "<your-demo-tenant-id>"
```

> 🔒 `demo.env` is gitignored - your credentials stay local

#### 5️⃣ Test Connection

```powershell
# Import the Graph module
Import-Module Microsoft.Graph.Authentication

# Connect with your demo app
Connect-MgGraph -ClientId "<client-id>" -TenantId "<tenant-id>" -Scopes "Mail.ReadWrite","User.Read" -UseDeviceCode

# Verify connection
Get-MgContext

# Test mailbox access
Get-MgUserMailFolder -UserId me -MailFolderId Inbox
```

#### 6️⃣ Run Full Test

```powershell
# Create demo-specific .env
Copy-Item demo.env .env

# Connect and test
.\Connect-OutlookRulesApp.ps1

# Deploy rules to demo mailbox
.\Manage-OutlookRules.ps1 -Operation Deploy

# Verify rules
.\Manage-OutlookRules.ps1 -Operation List

# Check folder creation
.\Manage-OutlookRules.ps1 -Operation Folders
```

### Demo Testing Checklist

- [ ] App registered in demo tenant
- [ ] Authentication settings configured (public client flow enabled)
- [ ] API permissions added (Mail.ReadWrite, User.Read)
- [ ] Device code authentication working
- [ ] Can access demo mailbox via Graph API
- [ ] Rules deploy successfully
- [ ] Folders created correctly
- [ ] Test email triggers correct rule

### Cleaning Up Demo Environment

```powershell
# Remove all rules from demo mailbox
.\Manage-OutlookRules.ps1 -Operation DeleteAll

# Restore production config
Copy-Item .env.backup .env
```

---

## 📁 Folder Structure

| Folder | Icon | Purpose |
|:-------|:----:|:--------|
| **Priority** | ⭐ | Messages from VIP senders (manager, skip-level, key collaborators) |
| **Action Required** | 🔴 | Time-sensitive items requiring response |
| **Metrics** | 📊 | Performance data, KPIs, Connect, QBR content |
| **Leadership** | 👔 | Executive communications and staff meeting notes |
| **Alerts** | 🔔 | System notifications and digests (auto-read) |
| **Low Priority** | 📭 | Newsletters, marketing, noise (archived or deleted) |

---

## 📜 Rules Summary

| # | Rule | Trigger | Actions |
|:-:|:-----|:--------|:--------|
| 01 | **Priority Senders** | From VIP list | 📁 Move to Priority, ⚡ High importance, 🛑 Stop processing |
| 02 | **Action Required** | Subject contains action keywords | 🏷️ Category, ⚡ High importance, 🚩 Flag, 📁 Move |
| 03 | **Connect & Metrics** | Subject/Body contains metrics keywords | 🏷️ Category, 🚩 Flag, 📁 Move |
| 04 | **Leadership & Exec** | Subject contains leadership keywords | ⚡ High importance, 📁 Move |
| 05 | **Alerts & Notifications** | Subject contains alert keywords | ✅ Mark read, 📁 Move |
| 99 | **Noise Filter** | From noise domains | 📦 Archive or 🗑️ Delete |

---

## ⚙️ Configuration

All rules are defined in `rules-config.json` - version controllable and reviewable.

### 📧 Priority Senders

```json
"senderLists": {
  "priority": {
    "description": "VIP senders who get priority treatment",
    "addresses": [
      "manager@company.com",
      "skip.level@company.com",
      "key.collaborator@company.com"
    ]
  }
}
```

### 🔤 Keywords

```json
"keywordLists": {
  "action": {
    "keywords": ["Action", "Approval", "Response Needed", "Due", "Deadline"]
  },
  "leadership": {
    "keywords": ["Leadership", "Executive", "LT", "Review", "Staff Meeting"]
  },
  "metrics": {
    "keywords": ["Connect", "ACR", "Performance", "Impact", "KPI", "QBR"]
  }
}
```

### 🔇 Noise Handling

```json
"settings": {
  "noiseAction": "Archive"  // or "Delete"
}
```

---

## 📂 File Reference

| File | Description |
|:-----|:------------|
| 📜 `prereqs.ps1` | Installs required PowerShell modules |
| 📜 `Register-OutlookRulesApp.ps1` | Creates Azure AD app registration |
| 📜 `Connect-OutlookRulesApp.ps1` | Authenticates to Graph and Exchange Online |
| 📜 `Setup-OutlookRules.ps1` | Creates folders and inbox rules (one-shot setup) |
| 📜 `Manage-OutlookRules.ps1` | Full rules management CLI |
| 📄 `rules-config.json` | Declarative rule definitions |
| 🔒 `.env` | Azure AD credentials (gitignored) |
| 🔒 `app-config.json` | Backup config (gitignored) |
| 📖 `docs/SDL.md` | SDL compliance documentation |
| 📖 `docs/SECURITY-QUESTIONNAIRE.md` | Security questionnaire |

---

## 🛠️ Management Operations

### Quick Reference

| Operation | Command | Description |
|:----------|:--------|:------------|
| 📋 List | `-Operation List` | View all current inbox rules |
| 🔍 Show | `-Operation Show -RuleName "..."` | Show details of a specific rule |
| 📤 Export | `-Operation Export` | Export rules to JSON file |
| 💾 Backup | `-Operation Backup` | Create timestamped backup |
| 📥 Import | `-Operation Import -ExportPath "..."` | Restore from backup |
| 🔄 Compare | `-Operation Compare` | Compare deployed vs config |
| 🚀 Deploy | `-Operation Deploy` | Deploy rules from config |
| ⬇️ Pull | `-Operation Pull` | Pull deployed rules to config |
| ✅ Enable | `-Operation Enable -RuleName "..."` | Enable a rule |
| ⏸️ Disable | `-Operation Disable -RuleName "..."` | Disable a rule |
| 🗑️ Delete | `-Operation Delete -RuleName "..."` | Delete a rule |
| 📊 Stats | `-Operation Stats` | Show mailbox statistics |
| ✔️ Validate | `-Operation Validate` | Check rules for issues |

### Examples

```powershell
# List all rules
.\Manage-OutlookRules.ps1 -Operation List

# Compare before deploying
.\Manage-OutlookRules.ps1 -Operation Compare

# Deploy changes
.\Manage-OutlookRules.ps1 -Operation Deploy

# Create backup
.\Manage-OutlookRules.ps1 -Operation Backup

# Show statistics
.\Manage-OutlookRules.ps1 -Operation Stats
```

---

## 🔐 Security

### Permissions Model

This solution uses **delegated permissions** only - no admin consent required for your own mailbox:

| Permission | Purpose | Admin Consent |
|:-----------|:--------|:-------------:|
| `Mail.ReadWrite` | Create mail folders under Inbox | ❌ |
| `User.Read` | Basic profile for authentication | ❌ |
| Exchange Online | Manage inbox rules | ❌ |

### Security Features

| Feature | Status | Description |
|:--------|:------:|:------------|
| 🔑 Device Code Flow | ✅ | SPACE-compliant authentication |
| 🚫 No Localhost | ✅ | No localhost redirect URIs |
| 🔒 No Secrets | ✅ | Public client, no client secrets |
| 📋 Delegated Only | ✅ | Cannot access other users' mailboxes |
| 🏠 Single Tenant | ✅ | App works only in registered tenant |
| 🔐 Gitignored Creds | ✅ | Sensitive files excluded from repo |

### SDL Compliance

For Azure AD admin consent requests, this application follows the **Shadow Org SDL Self-Attestation** process.

| Document | Description |
|:---------|:------------|
| 📄 [docs/SDL.md](docs/SDL.md) | SDL compliance, security controls, threat model |
| 📄 [docs/SECURITY-QUESTIONNAIRE.md](docs/SECURITY-QUESTIONNAIRE.md) | Highly Confidential permissions questionnaire |

---

## 🔧 Troubleshooting

<details>
<summary><strong>❌ "Admin approval required" error</strong></summary>

Your tenant has disabled user consent. Options:
1. Use your personal Azure tenant for the app registration
2. Request IT to consent to `Mail.ReadWrite` for Microsoft Graph PowerShell
3. Request an app registration from IT

</details>

<details>
<summary><strong>❌ "Cannot bind parameter" on New-InboxRule</strong></summary>

Ensure you're connected to Exchange Online:
```powershell
Get-ConnectionInformation
# If not connected:
Connect-ExchangeOnline
```

</details>

<details>
<summary><strong>❌ Rules not applying</strong></summary>

1. Verify rules are enabled: `Get-InboxRule | Select-Object Name, Enabled`
2. Check rule priority order: `Get-InboxRule | Select-Object Name, Priority | Sort-Object Priority`
3. Remember Priority Senders rule stops processing - VIP mail won't hit other rules

</details>

<details>
<summary><strong>❌ Device code authentication fails</strong></summary>

1. Verify app has **Allow public client flows** = **Yes**
2. Check redirect URI is set to `https://login.microsoftonline.com/common/oauth2/nativeclient`
3. Ensure user has access to the tenant

</details>

---

## 📝 License

MIT License - Use freely, modify as needed.

---

<p align="center">
  <strong>Made with ❤️ for inbox sanity</strong>
</p>
