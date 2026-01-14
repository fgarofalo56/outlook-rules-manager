<p align="center">
  <img src="https://img.shields.io/badge/SDL-Compliant-green?style=for-the-badge" alt="SDL Compliant">
  <img src="https://img.shields.io/badge/Security-Self--Attestation-blue?style=for-the-badge" alt="Self-Attestation">
</p>

<h1 align="center">🔒 Secure Development Lifecycle (SDL) Compliance</h1>

<p align="center">
  <strong>SDL compliance documentation for the Outlook Rules Manager application</strong>
</p>

---

## 📋 Overview

All Microsoft applications requiring Azure AD admin consent must complete the Secure Development Lifecycle (SDL) process per Microsoft Policy.

> This application follows the **Shadow Org SDL Self-Attestation** process.

---

## 🎯 SDL Process for Shadow Orgs

| Step | Description | Status |
|:----:|:------------|:------:|
| 1️⃣ | Self-Approval Process | ✅ |
| 2️⃣ | Self-Attestation Workbook | 📝 |
| 3️⃣ | Documentation | ✅ |

### 📚 Resources

| Resource | Link |
|:---------|:-----|
| 📄 Shadow Org SDL Self Attestation Guide | [Guide.docx](https://microsoft.sharepoint.com/:w:/t/ShadowSDLSupport) |
| 📊 SDL Self-Attestation Workbook | Download from SharePoint and save to OneDrive |
| 💬 SDL Support | Contact @ShadowSDLSupport in Teams |
| 🔗 Microsoft Digital SDL Process | [CSEO SDL Steps](https://aka.ms/cseosdl) |
| 🔗 Other Org SDL Processes | [Org SDL Directory](https://aka.ms/orgsdl) |

---

## 🛡️ Security Controls Assessment

### 🔐 Authentication & Authorization

| Control | Status | Notes |
|:--------|:------:|:------|
| 🔑 OAuth 2.0 Device Code Flow | ✅ | SPACE-compliant, no localhost redirect URIs |
| 📋 Delegated permissions only | ✅ | No application permissions |
| 🏠 Single-tenant app | ✅ | Registered in work tenant |
| 🔒 No secrets stored | ✅ | Interactive auth only, no client secrets |
| 🚫 No localhost redirect URIs | ✅ | Uses device code flow for SPACE/ACE compliance |

### 🗄️ Data Protection

| Control | Status | Notes |
|:--------|:------:|:------|
| 🔐 No credentials in code | ✅ | OAuth interactive flow |
| 📁 Sensitive files gitignored | ✅ | `.env`, `app-config.json`, backups |
| 👤 User mailbox scope only | ✅ | Delegated permissions limit access to signed-in user |
| 🚫 No data exfiltration | ✅ | Rules managed locally, no external transmission |

### 🔏 Permissions (Principle of Least Privilege)

| Permission | Type | Justification |
|:-----------|:----:|:--------------|
| `Mail.ReadWrite` | Delegated | Required to create mail folders under Inbox |
| `User.Read` | Delegated | Required for basic authentication profile |

> 📝 **Note**: No admin consent required for delegated permissions on user's own mailbox.

### 💻 Code Security

| Control | Status | Notes |
|:--------|:------:|:------|
| ✅ Input validation | Implemented | Config file parsing validates structure |
| ❌ No SQL/injection vectors | N/A | PowerShell scripts, no database |
| ❌ No web interface | N/A | CLI tool only |
| 📝 Signed scripts | Optional | Can enable with `Set-ExecutionPolicy` |

### ⚙️ Operational Security

| Control | Status | Notes |
|:--------|:------:|:------|
| 💾 Backup before destructive ops | ✅ | `DeleteAll` creates automatic backup |
| ⚠️ Confirmation prompts | ✅ | Destructive operations require confirmation |
| 📋 Audit trail | ✅ | Console output logs all operations |
| 🔑 No persistent tokens | ✅ | Tokens managed by MSAL, not stored by scripts |

---

## 🎯 Threat Model Summary

### 📦 Assets

| Asset | Description |
|:------|:------------|
| 📧 User's Exchange Online mailbox | Primary target for management |
| 📜 Inbox rules configuration | Rule definitions and settings |
| 📁 Mail folder structure | Organizational hierarchy |

### 👤 Threat Actors

| Actor | Risk Level |
|:------|:----------:|
| 🦠 Malicious scripts attempting to access mailbox | 🟡 Medium |
| 👤 Unauthorized users with access to workstation | 🟡 Medium |

### 🛡️ Mitigations

| # | Mitigation | Description |
|:-:|:-----------|:------------|
| 1 | 🔑 **Device Code Flow authentication** | SPACE-compliant, no localhost redirect URIs |
| 2 | 📋 **Delegated permissions only** | Cannot access other users' mailboxes |
| 3 | 🏠 **Single-tenant registration** | App only works in registered tenant |
| 4 | 🔓 **Public client flow** | No client secrets to leak |
| 5 | 🔐 **Gitignored sensitive files** | Credentials not committed to source control |
| 6 | 🚫 **No sensitive redirect URIs** | Uses device code flow (microsoft.com/devicelogin) |

---

## ✅ SDL Compliance Checklist

| Item | Status |
|:-----|:------:|
| 📖 Reviewed Shadow Org SDL Self Attestation Guide | ☐ |
| 📊 Completed SDL Self-Attestation Workbook | ☐ |
| 👔 FTE Manager approval obtained | ☐ |
| 👥 Team agreement documented | ☐ |
| 🛡️ Security controls verified (see tables above) | ☐ |
| 🎯 Threat model reviewed | ☐ |

---

## 🔗 SDL URL for SPACE Request

For the SDL URL field on the Security Portal for ACE (SPACE) request form, provide one of:

| Option | URL Format |
|:-------|:-----------|
| 1️⃣ S360 SDL Metric URL | Microsoft Digital/CSEO org only |
| 2️⃣ 1CS Work Item URL | If tracked in 1CS |
| 3️⃣ This document URL | GitHub/ADO link to this SDL.md file |

**Example**: `https://github.com/[your-repo]/blob/main/docs/SDL.md`

---

## 📞 Contact

| Resource | Contact |
|:---------|:--------|
| 💬 SDL Support | @ShadowSDLSupport (Teams) |
| 🏢 LinkedIn Org | Contact LinkedIn Security team |
| 🏢 Other Orgs | Contact your organization's security team |

---

## 📚 Related Documents

| Document | Description |
|:---------|:------------|
| 📄 [SECURITY-QUESTIONNAIRE.md](SECURITY-QUESTIONNAIRE.md) | Highly Confidential permissions questionnaire |
| 📖 [rules-cheatsheet.md](rules-cheatsheet.md) | Quick reference for rule management |

---

## 📅 Version History

| Date | Version | Changes |
|:-----|:--------|:--------|
| 2026-01-12 | 1.0.0 | Initial SDL documentation |
| 2026-01-14 | 1.1.0 | Updated for SPACE compliance (device code flow) |

---

<p align="center">
  <em>This document serves as SDL compliance evidence for Azure AD admin consent requests.</em>
</p>
