# Secure Development Lifecycle (SDL) Compliance

This document provides SDL compliance information for the Outlook Rules Manager application.

## Overview

All Microsoft applications requiring Azure AD admin consent must complete the Secure Development Lifecycle (SDL) process per Microsoft Policy.

## SDL Process for Shadow Orgs (Self-Attestation)

This application follows the **Shadow Org SDL Self-Attestation** process:

1. **Self-Approval Process**: You, your team, and your FTE manager agree the service meets all applicable security controls
2. **Self-Attestation Workbook**: Complete the SDL self-attestation workbook
3. **Documentation**: Maintain this document as evidence of SDL compliance

### Resources

| Resource | Link |
|----------|------|
| Shadow Org SDL Self Attestation Guide | [Guide.docx](https://microsoft.sharepoint.com/:w:/t/ShadowSDLSupport) |
| SDL Self-Attestation Workbook | Download from SharePoint and save to OneDrive |
| SDL Support | Contact @ShadowSDLSupport in Teams |
| Microsoft Digital SDL Process | [CSEO SDL Steps](https://aka.ms/cseosdl) |
| Other Org SDL Processes | [Org SDL Directory](https://aka.ms/orgsdl) |

## Security Controls Assessment

### Authentication & Authorization

| Control | Status | Notes |
|---------|--------|-------|
| OAuth 2.0 with PKCE | Implemented | Public client flow with MSAL |
| Delegated permissions only | Implemented | No application permissions |
| Single-tenant app | Implemented | Registered in work tenant |
| No secrets stored | Implemented | Interactive auth only |

### Data Protection

| Control | Status | Notes |
|---------|--------|-------|
| No credentials in code | Implemented | OAuth interactive flow |
| Sensitive files gitignored | Implemented | `.env`, `app-config.json`, backups |
| User mailbox scope only | Implemented | Delegated permissions limit access to signed-in user |
| No data exfiltration | Implemented | Rules managed locally, no external transmission |

### Permissions (Principle of Least Privilege)

| Permission | Type | Justification |
|------------|------|---------------|
| `Mail.ReadWrite` | Delegated | Required to create mail folders under Inbox |
| `User.Read` | Delegated | Required for basic authentication profile |

**Note**: No admin consent required for delegated permissions on user's own mailbox.

### Code Security

| Control | Status | Notes |
|---------|--------|-------|
| Input validation | Implemented | Config file parsing validates structure |
| No SQL/injection vectors | N/A | PowerShell scripts, no database |
| No web interface | N/A | CLI tool only |
| Signed scripts | Optional | Can enable with `Set-ExecutionPolicy` |

### Operational Security

| Control | Status | Notes |
|---------|--------|-------|
| Backup before destructive ops | Implemented | `DeleteAll` creates automatic backup |
| Confirmation prompts | Implemented | Destructive operations require confirmation |
| Audit trail | Implemented | Console output logs all operations |
| No persistent tokens | Implemented | Tokens managed by MSAL, not stored by scripts |

## Threat Model Summary

### Assets
- User's Exchange Online mailbox
- Inbox rules configuration
- Mail folder structure

### Threat Actors
- Malicious scripts attempting to access mailbox
- Unauthorized users with access to workstation

### Mitigations
1. **Interactive authentication required** - No stored credentials
2. **Delegated permissions only** - Cannot access other users' mailboxes
3. **Single-tenant registration** - App only works in registered tenant
4. **Public client flow** - No client secrets to leak
5. **Gitignored sensitive files** - Credentials not committed to source control

## SDL Compliance Checklist

- [ ] Reviewed Shadow Org SDL Self Attestation Guide
- [ ] Completed SDL Self-Attestation Workbook
- [ ] FTE Manager approval obtained
- [ ] Team agreement documented
- [ ] Security controls verified (see tables above)
- [ ] Threat model reviewed

## SDL URL for SPACE Request

For the SDL URL field on the Security Portal for ACE (SPACE) request form, provide one of:

1. **S360 SDL Metric URL** (Microsoft Digital/CSEO org only)
2. **1CS Work Item URL** (if tracked in 1CS)
3. **This document URL** (GitHub/ADO link to this SDL.md file)

Example: `https://github.com/[your-repo]/blob/main/docs/SDL.md`

## Contact

- **SDL Support**: @ShadowSDLSupport (Teams)
- **LinkedIn Org**: Contact LinkedIn Security team
- **Other Orgs**: Contact your organization's security team

## Related Documents

- [SECURITY-QUESTIONNAIRE.md](SECURITY-QUESTIONNAIRE.md) - Highly Confidential permissions questionnaire
- [rules-cheatsheet.md](rules-cheatsheet.md) - Quick reference for rule management

## Version History

| Date | Version | Changes |
|------|---------|---------|
| 2026-01-12 | 1.0.0 | Initial SDL documentation |

---

*This document serves as SDL compliance evidence for Azure AD admin consent requests.*
