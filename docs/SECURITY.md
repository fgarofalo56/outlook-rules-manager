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
| Unauthorized mailbox access | Delegated OAuth, MFA | Low |
| Supply chain attack | Microsoft-signed modules only | Low |
| Configuration tampering | File permissions (user responsibility) | Medium |
| Path traversal | Input validation (implemented) | Low |

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

Before using this tool:

- [ ] Verify `.env` is in `.gitignore`
- [ ] Verify `rules-config.json` is in `.gitignore`
- [ ] Run `.\scripts\Check-BeforeCommit.ps1` before any commit
- [ ] Use example.com addresses in any shared configurations
- [ ] Review rule conditions before deploying (especially ForwardTo/RedirectTo)
- [ ] Protect configuration files with appropriate file permissions

## Updates

This security policy is reviewed and updated with each major release.

**Last Updated:** 2024-01-14
**Version:** 1.0.0
