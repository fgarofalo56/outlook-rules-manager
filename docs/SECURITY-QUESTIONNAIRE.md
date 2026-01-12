# Highly Confidential Permissions - Security Questionnaire

This document provides responses to the security questionnaire required for Azure AD admin consent requests involving Highly Confidential permissions.

---

## Application Overview

| Field | Value |
|-------|-------|
| **Application Name** | Outlook Rules Manager (frgarofa-graphapp) |
| **Application Type** | PowerShell Local Scripts |
| **Client ID** | `5940cf1d-31a0-4fcc-b2c4-18b5af54528f` |
| **Tenant ID** | `72f988bf-86f1-41af-91ab-2d7cd011db47` |
| **Data Classification** | High Confidential |

---

## Critical Security Questions

### 1. Service Tree Entry *

**Status**: ☐ Pending / ☑ Complete

- **Service Tree ID**: `[INSERT SERVICE TREE ID]`
- **Service Tree URL**: `https://servicetree.msftcloudes.com/[SERVICE-ID]`
- **Data Classification**: High Confidential ✓
- **Azure Subscriptions Linked**: ☐ Yes / ☑ N/A (Local scripts, no Azure resources)
- **Code Repository Registered**: ☐ Yes (provide link below)

**Code Repository URL**:
```
[INSERT ADO/GITHUB REPOSITORY URL]
Example: https://dev.azure.com/[org]/[project]/_git/outlook-rules-manager
```

---

### 2. API Calls Documentation *

**All Microsoft Graph API calls made by this application:**

| Endpoint | Method | Permission | Purpose |
|----------|--------|------------|---------|
| `/me` | GET | User.Read | Get authenticated user profile |
| `/me/mailFolders/Inbox` | GET | Mail.ReadWrite | Access Inbox folder |
| `/me/mailFolders/{id}/childFolders` | GET | Mail.ReadWrite | List inbox subfolders |
| `/me/mailFolders/{id}/childFolders` | POST | Mail.ReadWrite | Create new mail folders |
| `/me/mailFolders/{id}` | DELETE | Mail.ReadWrite | Remove mail folders |

**Exchange Online PowerShell cmdlets used:**

| Cmdlet | Purpose |
|--------|---------|
| `Get-InboxRule` | List existing inbox rules |
| `New-InboxRule` | Create new inbox rules |
| `Set-InboxRule` | Modify existing rules |
| `Remove-InboxRule` | Delete inbox rules |

**Note**: All operations are delegated (user context only). No application permissions are used.

---

### 3. Code Analysis / Security Scanning *

**Code scanning product**: `[SELECT ONE]`

- ☐ Microsoft Security DevOps (MSDO)
- ☐ CodeQL / GitHub Advanced Security
- ☐ Defender for DevOps
- ☐ CredScan
- ☐ Other: ________________

**Scan Results Task Link** *:
```
[INSERT LINK TO SCAN RESULTS]
Example: https://dev.azure.com/[org]/[project]/_build/results?buildId=[ID]
```

**Manual Security Review Notes**:

For PowerShell local scripts, the following security controls are in place:

| Check | Status | Notes |
|-------|--------|-------|
| No hardcoded credentials | ✓ Pass | OAuth interactive flow only |
| No secrets in source code | ✓ Pass | `.env` and config files gitignored |
| Input validation | ✓ Pass | Config file structure validated |
| No SQL injection vectors | ✓ N/A | No database operations |
| No command injection | ✓ Pass | No user input executed as commands |
| Secure file operations | ✓ Pass | Paths validated, no arbitrary file access |

---

### 4. SDL Task Link *

**SDL Documentation**: [docs/SDL.md](SDL.md)

**SDL Task URL**:
```
[INSERT SDL TASK URL]
Options:
- S360 SDL Metric URL (CSEO org): https://s360.microsoft.com/[path]
- 1CS Work Item URL: https://1cs.microsoft.com/[task-id]
- GitHub/ADO URL: https://[repo]/docs/SDL.md
```

**SDL Compliance Status**: Shadow Org Self-Attestation (See [SDL.md](SDL.md))

---

### 5. Azure Subscription Health Scanning *

**Azure subscription health product**: `[SELECT ONE]`

- ☐ Microsoft Defender for Cloud
- ☐ Azure Security Center
- ☐ Azure Policy
- ☐ S360 Azure Health
- ☑ **N/A - Local PowerShell Scripts**

**Justification for N/A**:
This application runs as local PowerShell scripts on the user's workstation. There are no Azure resources deployed:
- No Azure App Service
- No Azure Functions
- No Azure Storage
- No Azure VMs
- No cloud infrastructure

The application connects to existing Microsoft 365 services (Exchange Online, Microsoft Graph) using delegated user permissions only.

---

### 6. Component Governance *

**Component Governance Status**: `[SELECT ONE]`

- ☐ Onboarded to Component Governance
- ☑ **Not Applicable**

**Component Governance Link**:
```
[INSERT LINK IF APPLICABLE]
Example: https://s360.microsoft.com/[path]/componentgovernance
Example: https://dev.azure.com/[org]/[project]/_componentGovernance
```

**Justification for N/A**:
This project does not use package management tools in its build environment:
- No npm/yarn packages
- No NuGet packages
- No pip/Python packages
- Pure PowerShell scripts with no external dependencies beyond Microsoft-provided modules

**PowerShell modules used** (Microsoft-provided, not third-party):
| Module | Source | Purpose |
|--------|--------|---------|
| Microsoft.Graph.Authentication | PSGallery (Microsoft) | Graph API authentication |
| Microsoft.Graph.Mail | PSGallery (Microsoft) | Mail folder operations |
| ExchangeOnlineManagement | PSGallery (Microsoft) | Exchange Online rules |
| Az.Accounts | PSGallery (Microsoft) | Azure authentication |
| Az.Resources | PSGallery (Microsoft) | App registration |

---

## Additional Security Controls

### Authentication Security

| Control | Implementation |
|---------|----------------|
| Authentication Method | OAuth 2.0 with PKCE |
| Token Storage | MSAL managed (not stored by application) |
| Session Persistence | None (interactive auth each session) |
| Multi-factor Authentication | Enforced by Azure AD policies |

### Data Handling

| Control | Implementation |
|---------|----------------|
| Data at Rest | No data stored by application |
| Data in Transit | TLS 1.2+ (Microsoft Graph/Exchange) |
| Credential Storage | None (no secrets stored) |
| Logging | Console output only (no persistent logs) |

### Access Controls

| Control | Implementation |
|---------|----------------|
| Least Privilege | Delegated permissions only |
| Scope Limitation | User's own mailbox only |
| Admin Consent | Not required for delegated permissions |
| Conditional Access | Inherited from Azure AD policies |

---

## Checklist Summary

| Requirement | Status | Evidence |
|-------------|--------|----------|
| Service Tree marked High Confidential | ☐ | [Link] |
| Azure subscriptions linked | ☑ N/A | Local scripts |
| Code repository registered | ☐ | [Link] |
| API documentation | ✓ | This document |
| Code scanning active | ☐ | [Scan results link] |
| SDL task link | ☐ | [SDL link] |
| Azure subscription health | ☑ N/A | Local scripts |
| Component Governance | ☑ N/A | No package dependencies |

---

## Contact Information

| Role | Contact |
|------|---------|
| Application Owner | `[YOUR ALIAS]@microsoft.com` |
| Security Contact | `[SECURITY CONTACT]@microsoft.com` |
| Manager | `[MANAGER ALIAS]@microsoft.com` |

---

## Version History

| Date | Version | Changes |
|------|---------|---------|
| 2026-01-12 | 1.0.0 | Initial questionnaire |

---

*This document serves as evidence for the Highly Confidential permissions security review.*
