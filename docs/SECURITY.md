# Security Policy

## Reporting Security Vulnerabilities

If you discover a security vulnerability in this project, please report it responsibly:

1. **Do NOT** create a public GitHub issue
2. Contact the maintainer directly via email
3. Provide detailed information about the vulnerability
4. Allow reasonable time for a fix before public disclosure

## Security Model

### Authentication

This tool uses **delegated OAuth 2.0 authentication** with the following security properties:

| Property | Implementation |
|----------|---------------|
| Flow Type | Device Code Flow (SPACE/ACE compliant) |
| Token Storage | Session memory only (no persistence) |
| Token Lifetime | Session-scoped (discarded on exit) |
| MFA Support | Respected via Azure AD policies |
| Consent | User consent required (no admin consent) |

### Permissions (Principle of Least Privilege)

| Permission | Purpose | Scope |
|------------|---------|-------|
| `Mail.ReadWrite` | Create/manage mail folders | User's mailbox only |
| `User.Read` | Basic profile for authentication | User's profile only |

**Cannot access:**
- Other users' mailboxes
- Tenant-wide settings
- Admin operations
- Calendar, contacts, or other data

### Multi-Tier Authorization Model

This application implements a **defense-in-depth authorization model** that goes beyond basic permission consent:

```
┌─────────────────────────────────────────────────────────────────┐
│                    Authorization Layers                          │
├─────────────────────────────────────────────────────────────────┤
│  Layer 1: Azure AD Authentication                               │
│  └── User must authenticate via OAuth 2.0 device code flow     │
│                                                                 │
│  Layer 2: User Assignment Required (Azure AD)                   │
│  └── User must be explicitly assigned to the application       │
│  └── Unapproved users blocked at sign-in                       │
│                                                                 │
│  Layer 3: App Role Assignment (Azure AD)                        │
│  └── User must have OutlookRules.Admin or OutlookRules.User    │
│  └── Role determines capabilities within the application       │
│                                                                 │
│  Layer 4: Script-Level Authorization (PowerShell)               │
│  └── Connect script validates role claims                      │
│  └── Defense-in-depth against Azure AD bypass                  │
└─────────────────────────────────────────────────────────────────┘
```

#### Authorization Tiers

| Tier | Role | Entra ID Assignment | Capabilities |
|------|------|---------------------|--------------|
| **Owner** | Service Principal Owner | Azure AD built-in | Manage admins, full Azure AD control |
| **Admin** | OutlookRules.Admin | App Role Assignment | Add/remove authorized users, all app operations |
| **User** | OutlookRules.User | App Role Assignment | Manage own mailbox only |

#### App Role Definitions

| Role Value | Role ID | Description |
|------------|---------|-------------|
| `OutlookRules.Admin` | `f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f` | Full application access + user management |
| `OutlookRules.User` | `a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d` | Self-service mailbox management only |

#### Authorization Flow

```
User attempts sign-in
        │
        ▼
┌───────────────────┐    NO     ┌─────────────────┐
│ Valid Azure AD    │──────────►│ ACCESS DENIED   │
│ credentials?      │           │ (Auth failure)  │
└────────┬──────────┘           └─────────────────┘
         │ YES
         ▼
┌───────────────────┐    NO     ┌─────────────────┐
│ User assigned to  │──────────►│ ACCESS DENIED   │
│ application?      │           │ (Not assigned)  │
└────────┬──────────┘           └─────────────────┘
         │ YES
         ▼
┌───────────────────┐    NO     ┌─────────────────┐
│ Has valid app     │──────────►│ ACCESS DENIED   │
│ role?             │           │ (No role)       │
└────────┬──────────┘           └─────────────────┘
         │ YES
         ▼
┌───────────────────┐
│ ACCESS GRANTED    │
│ (Role determines  │
│  capabilities)    │
└───────────────────┘
```

#### Security Benefits

| Benefit | Description |
|---------|-------------|
| **Defense-in-Depth** | Multiple authorization checks at different layers |
| **Centralized Control** | IT admins manage access via Azure AD, not app config |
| **Audit Trail** | All role assignments logged in Azure AD audit logs |
| **Zero Trust Alignment** | Explicit trust verification for all access |
| **Least Privilege** | Users only receive permissions needed for their role |
| **Separation of Duties** | Admins manage access, users manage mailboxes |

#### Enabling User Assignment Required

This setting **MUST** be enabled for secure operation:

1. Azure Portal → Microsoft Entra ID → Enterprise Applications
2. Select "Outlook Rules Manager"
3. Properties → Set "Assignment required?" to **Yes**
4. Save

Or via PowerShell:
```powershell
Connect-MgGraph -Scopes "Application.ReadWrite.All"
$sp = Get-MgServicePrincipal -Filter "DisplayName eq 'Outlook Rules Manager'"
Update-MgServicePrincipal -ServicePrincipalId $sp.Id -AppRoleAssignmentRequired:$true
```

**WARNING:** If "User Assignment Required" is disabled, any authenticated user can access the application!

### Data Classification

| Data Type | Classification | Protection |
|-----------|---------------|------------|
| ClientId | Public | Can be shared (like a username) |
| TenantId | Public | Can be shared (like a domain name) |
| Email addresses | PII | Gitignored, never committed |
| Rule conditions | Contains PII | Gitignored, never committed |
| Exported rules | Contains PII | Gitignored, never committed |

## Threat Model

### In-Scope Threats

| Threat | Mitigation | Residual Risk |
|--------|------------|---------------|
| Credential leak to repository | Gitleaks, pre-commit hooks, .gitignore | Low |
| PII exposure (emails) | Gitignore, pre-commit checks | Low |
| Unauthorized mailbox access | Delegated OAuth, MFA, User Assignment Required | Low |
| Unauthorized app usage | User Assignment Required, App Roles | Low |
| Privilege escalation | Role-based access control, admin-only user management | Low |
| Supply chain attack | Microsoft-signed modules only | Low |
| Configuration tampering | File permissions (user responsibility) | Medium |
| Path traversal | Input validation (implemented) | Low |
| Forwarding/redirect abuse | Email validation, SecurityHelpers module | Low |
| OOO message injection | HTML sanitization, ConvertTo-SafeText | Low |

### Out-of-Scope Threats

| Threat | Reason |
|--------|--------|
| Compromised workstation | Beyond application scope |
| Stolen OAuth tokens | Handled by Microsoft identity platform |
| Azure AD compromise | Beyond application scope |
| Exchange Online vulnerabilities | Microsoft's responsibility |

## Security Controls

### Pre-Commit Controls

1. **Gitleaks** - Secret pattern detection
2. **Custom hooks** - Block sensitive files (.env, rules-config.json)
3. **PowerShell script** - `Check-BeforeCommit.ps1`

### CI/CD Controls

1. **GitHub Actions** - Automated security scanning on push
2. **Branch protection** - Required PR reviews and status checks
3. **PSScriptAnalyzer** - PowerShell code quality

### Runtime Controls

1. **Path validation** - Prevents file access outside script directory
2. **Input validation** - Email and domain format checking
3. **Schema validation** - Configuration structure verification
4. **Destructive operation protection** - Explicit confirmation required

## Secure Configuration

### .env File Security

The `.env` file contains:
```powershell
$ClientId = "..."  # Public - safe to share
$TenantId = "..."  # Public - safe to share
```

**Why this is safe:**
- ClientId is like a username - identifies the app but cannot authenticate alone
- TenantId is like a domain name - identifies the organization
- No client secrets or tokens are stored
- Authentication requires interactive user login

### rules-config.json Security

This file contains PII (email addresses) and should:
- Never be committed to version control
- Be protected with file permissions (read/write for owner only)
- Use example.com addresses in shared documentation

### Recommended File Permissions

```powershell
# Windows - restrict to current user
$acl = Get-Acl ".env"
$acl.SetAccessRuleProtection($true, $false)
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $env:USERNAME, "FullControl", "Allow")
$acl.SetAccessRule($rule)
Set-Acl ".env" $acl
```

## Known Limitations

1. **PowerShell History** - Commands containing email addresses may be saved in PSReadline history
2. **Console Output** - Email addresses displayed during rule listing (expected behavior)
3. **Memory** - Configuration loaded into memory during execution
4. **Audit Logging** - No persistent audit trail (enhancement planned)

## Compliance Notes

### Zero Trust Architecture

This tool aligns with Zero Trust principles:

| Principle | Implementation |
|-----------|---------------|
| Verify explicitly | Interactive OAuth authentication |
| Least privilege | Minimal scopes (Mail.ReadWrite, User.Read) |
| Assume breach | No persistent credentials, session-only tokens |

### Data Protection

- **PII Handling**: Email addresses treated as PII, excluded from version control
- **Data Minimization**: Only rule conditions stored, not email content
- **Right to Access**: Export operation available
- **Right to Delete**: DeleteAll operation available

## Security Checklist

### Initial Setup Security

- [ ] Run `.\Register-OutlookRulesApp.ps1` to create app with app roles
- [ ] Enable "User Assignment Required" in Enterprise Applications
- [ ] Run `.\Manage-AppAuthorization.ps1 -Operation Setup` to configure authorization
- [ ] Verify only authorized users are assigned app roles

### Pre-Commit Security

- [ ] Verify `.env` is in `.gitignore`
- [ ] Verify `rules-config.json` is in `.gitignore`
- [ ] Run `.\scripts\Check-BeforeCommit.ps1` before any commit
- [ ] Run `gitleaks detect --source .` to scan for secrets
- [ ] Run `Invoke-ScriptAnalyzer -Path . -Recurse` for code quality

### Runtime Security

- [ ] Use example.com addresses in any shared configurations
- [ ] Review rule conditions before deploying (especially ForwardTo/RedirectTo)
- [ ] Protect configuration files with appropriate file permissions
- [ ] Verify authorization status: `.\Manage-AppAuthorization.ps1 -Operation Status`

### Ongoing Security

- [ ] Periodically review authorized users: `.\Manage-AppAuthorization.ps1 -Operation List`
- [ ] Remove unused user assignments
- [ ] Monitor Azure AD audit logs for role assignment changes
- [ ] Review forwarding rules for potential data exfiltration

## Updates

This security policy is reviewed and updated with each major release.

**Last Updated:** 2026-01-14
**Version:** 2.0.0

### Change Log

- **v2.0.0** (2026-01-14): Added multi-tier authorization model with App Roles
- **v1.0.0** (2024-01-14): Initial security policy
