<p align="center">
  <img src="https://img.shields.io/badge/Classification-Highly_Confidential-red?style=for-the-badge" alt="Highly Confidential">
  <img src="https://img.shields.io/badge/Status-Review_Required-yellow?style=for-the-badge" alt="Review Required">
</p>

<h1 align="center">🔐 Highly Confidential Permissions - Security Questionnaire</h1>

<p align="center">
  <strong>Security questionnaire for Azure AD admin consent requests</strong>
</p>

---

## 📋 Application Overview

| Field | Value |
|:------|:------|
| 🏷️ **Application Name** | Outlook Rules Manager |
| 💻 **Application Type** | PowerShell Local Scripts |
| 🔑 **Client ID** | `[YOUR_CLIENT_ID]` |
| 🏢 **Tenant ID** | `[YOUR_TENANT_ID]` |
| 🔒 **Data Classification** | High Confidential |

---

## 🔍 Critical Security Questions

### 1️⃣ Service Tree Entry

| Item | Status | Details |
|:-----|:------:|:--------|
| 📋 Service Tree ID | ☐ | `[INSERT SERVICE TREE ID]` |
| 🔗 Service Tree URL | ☐ | `https://servicetree.msftcloudes.com/[SERVICE-ID]` |
| 🔐 Data Classification | ✅ | High Confidential |
| ☁️ Azure Subscriptions Linked | ✅ N/A | Local scripts, no Azure resources |
| 📦 Code Repository Registered | ☐ | [Repository URL] |

---

### 2️⃣ API Calls Documentation

#### Microsoft Graph API Calls

| Endpoint | Method | Permission | Purpose |
|:---------|:------:|:-----------|:--------|
| `/me` | GET | User.Read | Get authenticated user profile |
| `/me/mailFolders/Inbox` | GET | Mail.ReadWrite | Access Inbox folder |
| `/me/mailFolders/{id}/childFolders` | GET | Mail.ReadWrite | List inbox subfolders |
| `/me/mailFolders/{id}/childFolders` | POST | Mail.ReadWrite | Create new mail folders |
| `/me/mailFolders/{id}` | DELETE | Mail.ReadWrite | Remove mail folders |

#### Exchange Online PowerShell Cmdlets

| Cmdlet | Purpose |
|:-------|:--------|
| 📋 `Get-InboxRule` | List existing inbox rules |
| ➕ `New-InboxRule` | Create new inbox rules |
| ✏️ `Set-InboxRule` | Modify existing rules |
| 🗑️ `Remove-InboxRule` | Delete inbox rules |

> 📝 **Note**: All operations are delegated (user context only). No application permissions are used.

---

### 3️⃣ Code Analysis / Security Scanning

#### Scanning Products

| Product | Status |
|:--------|:------:|
| 🔍 Microsoft Security DevOps (MSDO) | ☐ |
| 🔍 CodeQL / GitHub Advanced Security | ☐ |
| 🔍 Defender for DevOps | ☐ |
| 🔍 CredScan | ☐ |
| 🔍 PSScriptAnalyzer | ✅ |
| 🔍 Gitleaks | ✅ |

#### Scan Results

```
[INSERT LINK TO SCAN RESULTS]
Example: https://github.com/[repo]/actions
```

#### Manual Security Review

| Check | Status | Notes |
|:------|:------:|:------|
| 🔐 No hardcoded credentials | ✅ Pass | OAuth interactive flow only |
| 🔒 No secrets in source code | ✅ Pass | `.env` and config files gitignored |
| ✅ Input validation | ✅ Pass | Config file structure validated |
| 💉 No SQL injection vectors | ✅ N/A | No database operations |
| 🛡️ No command injection | ✅ Pass | No user input executed as commands |
| 📁 Secure file operations | ✅ Pass | Paths validated, no arbitrary file access |

---

### 4️⃣ SDL Task Link

| Item | Value |
|:-----|:------|
| 📄 SDL Documentation | [docs/SDL.md](SDL.md) |
| 🔗 SDL Task URL | `[INSERT SDL TASK URL]` |
| ✅ SDL Compliance Status | Shadow Org Self-Attestation |

**SDL URL Options:**
- S360 SDL Metric URL (CSEO org): `https://s360.microsoft.com/[path]`
- 1CS Work Item URL: `https://1cs.microsoft.com/[task-id]`
- GitHub/ADO URL: `https://[repo]/docs/SDL.md`

---

### 5️⃣ Azure Subscription Health Scanning

| Product | Status |
|:--------|:------:|
| ☁️ Microsoft Defender for Cloud | ☐ |
| ☁️ Azure Security Center | ☐ |
| ☁️ Azure Policy | ☐ |
| ☁️ S360 Azure Health | ☐ |
| ✅ **N/A - Local PowerShell Scripts** | ✅ |

**Justification for N/A:**

This application runs as local PowerShell scripts on the user's workstation:

| Resource | Present |
|:---------|:-------:|
| ☁️ Azure App Service | ❌ |
| ⚡ Azure Functions | ❌ |
| 💾 Azure Storage | ❌ |
| 🖥️ Azure VMs | ❌ |
| 🏗️ Cloud infrastructure | ❌ |

> The application connects to existing Microsoft 365 services (Exchange Online, Microsoft Graph) using delegated user permissions only.

---

### 6️⃣ Component Governance

| Status | Details |
|:-------|:--------|
| ☐ Onboarded to Component Governance | |
| ✅ **Not Applicable** | |

**Justification for N/A:**

This project does not use package management tools in its build environment:

| Package Type | Used |
|:-------------|:----:|
| 📦 npm/yarn packages | ❌ |
| 📦 NuGet packages | ❌ |
| 📦 pip/Python packages | ❌ |

**PowerShell modules used** (Microsoft-provided, not third-party):

| Module | Source | Purpose |
|:-------|:-------|:--------|
| 🔐 Microsoft.Graph.Authentication | PSGallery (Microsoft) | Graph API authentication |
| 📧 Microsoft.Graph.Mail | PSGallery (Microsoft) | Mail folder operations |
| 📬 ExchangeOnlineManagement | PSGallery (Microsoft) | Exchange Online rules |
| ☁️ Az.Accounts | PSGallery (Microsoft) | Azure authentication |
| 🏗️ Az.Resources | PSGallery (Microsoft) | App registration |

---

## 🛡️ Additional Security Controls

### 🔐 Authentication Security

| Control | Implementation |
|:--------|:---------------|
| 🔑 Authentication Method | OAuth 2.0 Device Code Flow |
| 💾 Token Storage | MSAL managed (not stored by application) |
| ⏱️ Session Persistence | None (interactive auth each session) |
| 🔒 Multi-factor Authentication | Enforced by Azure AD policies |

### 📦 Data Handling

| Control | Implementation |
|:--------|:---------------|
| 💾 Data at Rest | No data stored by application |
| 🔐 Data in Transit | TLS 1.2+ (Microsoft Graph/Exchange) |
| 🔑 Credential Storage | None (no secrets stored) |
| 📋 Logging | Console output only (no persistent logs) |

### 🔏 Access Controls

| Control | Implementation |
|:--------|:---------------|
| ⚖️ Least Privilege | Delegated permissions only |
| 🎯 Scope Limitation | User's own mailbox only |
| 👑 Admin Consent | Not required for delegated permissions |
| 🚪 Conditional Access | Inherited from Azure AD policies |

---

## ✅ Checklist Summary

| Requirement | Status | Evidence |
|:------------|:------:|:---------|
| 🏷️ Service Tree marked High Confidential | ☐ | [Link] |
| ☁️ Azure subscriptions linked | ✅ N/A | Local scripts |
| 📦 Code repository registered | ☐ | [Link] |
| 📋 API documentation | ✅ | This document |
| 🔍 Code scanning active | ☐ | [Scan results link] |
| 📄 SDL task link | ☐ | [SDL link] |
| ☁️ Azure subscription health | ✅ N/A | Local scripts |
| 📦 Component Governance | ✅ N/A | No package dependencies |

---

## 📞 Contact Information

| Role | Contact |
|:-----|:--------|
| 👤 Application Owner | `[YOUR ALIAS]@microsoft.com` |
| 🛡️ Security Contact | `[SECURITY CONTACT]@microsoft.com` |
| 👔 Manager | `[MANAGER ALIAS]@microsoft.com` |

---

## 📅 Version History

| Date | Version | Changes |
|:-----|:--------|:--------|
| 2026-01-12 | 1.0.0 | Initial questionnaire |
| 2026-01-14 | 1.1.0 | Updated with visual formatting |

---

<p align="center">
  <em>This document serves as evidence for the Highly Confidential permissions security review.</em>
</p>
