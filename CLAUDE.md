# Outlook Rules Manager - Claude Code Instructions

## Project Overview

PowerShell-based automation for managing Outlook inbox organization via Microsoft Graph and Exchange Online. Creates server-side rules and folders to automatically triage email.

---

## CRITICAL: Security-by-Design Development Rules

### Rule 1: Pre-Commit Security Validation (MANDATORY)

**BEFORE committing ANY code changes, Claude MUST run security scans locally:**

```powershell
# 1. Run the pre-commit security check script
.\scripts\Check-BeforeCommit.ps1

# 2. Run Gitleaks locally to detect secrets
gitleaks detect --source . --verbose

# 3. Run PSScriptAnalyzer for PowerShell security rules
Invoke-ScriptAnalyzer -Path . -Recurse -Severity Error,Warning

# 4. Verify no sensitive files are staged
git diff --cached --name-only | Select-String -Pattern '\.env|rules-config\.json|app-config\.json'
```

**If ANY scan fails or detects issues:**
- DO NOT commit
- Fix all issues first
- Re-run all scans until clean
- Document any false positives in `.gitleaksignore`

### Rule 2: Never Commit Sensitive Data

**ABSOLUTELY FORBIDDEN to commit:**
- `.env` files (any profile)
- `rules-config.json` or `rules-config.*.json`
- `app-config.json`
- `exported-rules.json`
- Any file containing: Client IDs, Tenant IDs, email addresses, GUIDs

**Before every commit, verify:**
```powershell
# Check what's staged
git status

# Verify gitignore is working
git check-ignore -v .env rules-config.json
```

### Rule 3: Input Validation on All User Data

**All user-provided data MUST be validated:**
- Email addresses: Use `Test-EmailAddress` from SecurityHelpers
- Domains: Use `Test-DomainName` from SecurityHelpers
- HTML content (OOO messages): Use `ConvertTo-SafeText` sanitization
- File paths: Validate against path traversal

### Rule 4: Security Module Required

**The SecurityHelpers module MUST be loaded for:**
- Any operation accepting email addresses
- Any operation with forwarding/redirect
- Any operation writing OOO messages
- Audit logging operations

### Rule 5: Authorization Layer Enforcement

**This application implements multi-tier access control:**

| Tier | Role | Capabilities |
|------|------|--------------|
| Owner | Service Principal Owner | Manage admins, full control |
| Admin | `OutlookRules.Admin` role | Add/remove authorized users, use app |
| User | `OutlookRules.User` role | Use app for own mailbox only |

**Authorization is enforced at TWO levels:**
1. **Azure AD Level**: "User Assignment Required" blocks unapproved users at sign-in
2. **Script Level**: Role claims validated before allowing operations

**NEVER bypass authorization checks. NEVER disable "User Assignment Required".**

---

## Architecture

```
outlook-rules-manager/
├── src/                           # Core PowerShell scripts
│   ├── Manage-OutlookRules.ps1        # Full management CLI
│   ├── Manage-AppAuthorization.ps1    # User authorization management
│   ├── Connect-OutlookRulesApp.ps1    # Authentication + authorization validation
│   ├── Register-OutlookRulesApp.ps1   # Azure AD app registration (with app roles)
│   ├── Install-Prerequisites.ps1      # Module installation script
│   ├── Setup-OutlookRules.ps1         # One-shot rules/folders creation (legacy)
│   └── modules/
│       └── SecurityHelpers.psm1       # Security helper module
├── scripts/                       # Utility scripts
│   └── Check-BeforeCommit.ps1         # Pre-commit security check
├── tests/                         # Pester unit tests
│   ├── SecurityHelpers.Tests.ps1
│   ├── ConfigParsing.Tests.ps1
│   └── Run-Tests.ps1
├── docs/                          # Documentation
│   ├── QUICKSTART.md
│   ├── USER-GUIDE.md
│   ├── TESTING-GUIDE.md
│   ├── SECURITY.md
│   ├── SDL.md
│   ├── SECURITY-QUESTIONNAIRE.md
│   └── rules-cheatsheet.md
├── examples/                      # Example configuration files
│   ├── .env.example
│   └── rules-config.example.json
├── .github/workflows/             # CI/CD and security scanning
├── rules-config.json              # Declarative rule definitions (gitignored)
├── .env                           # Azure AD credentials (gitignored)
├── .gitignore                     # Protects sensitive files
├── .gitleaks.toml                 # Secret scanning configuration
├── .pre-commit-config.yaml        # Pre-commit hook configuration
├── LICENSE                        # MIT License
└── README.md                      # Project documentation
```

## Key Files

| File | Purpose | When to Modify |
|------|---------|----------------|
| `rules-config.json` | All rule definitions | Adding/changing rules |
| `src/Manage-OutlookRules.ps1` | Management operations | Adding new operations |
| `src/Manage-AppAuthorization.ps1` | User authorization | Adding/removing users |
| `src/modules/SecurityHelpers.psm1` | Security validation | Adding security functions |

## Key APIs & Modules

| Module | Purpose |
|--------|---------|
| `Microsoft.Graph.Authentication` | OAuth connection to Graph API |
| `Microsoft.Graph.Mail` | Mail folder operations (`Get-MgUserMailFolder`, `New-MgUserMailFolderChildFolder`) |
| `ExchangeOnlineManagement` | Inbox rules (`New-InboxRule`, `Set-InboxRule`, `Get-InboxRule`) |
| `Az.Accounts`, `Az.Resources` | App registration management |

## Configuration Files

| File | Purpose | Gitignored? |
|------|---------|-------------|
| `.env` | Azure AD ClientId/TenantId (primary) | Yes |
| `.env.{profile}` | Profile-specific Azure AD config (e.g., `.env.personal`) | Yes |
| `app-config.json` | Azure AD config backup (JSON format) | Yes |
| `rules-config.json` | Rule definitions, sender lists, keywords | Yes |
| `rules-config.{profile}.json` | Profile-specific rules (e.g., `rules-config.work.json`) | Yes |
| `exported-rules.json` | Backup of deployed rules | Yes |
| `examples/.env.example` | Example .env template | No |
| `examples/rules-config.example.json` | Example rules config template | No |

## Common Tasks

### Add a new priority sender
Edit `rules-config.json` > `senderLists.priority.addresses`:
```json
"addresses": [
  "existing@email.com",
  "new.sender@email.com"
]
```
Then: `.\src\Manage-OutlookRules.ps1 -Operation Deploy`

### Add new keywords
Edit `rules-config.json` > `keywordLists.<list>.keywords`:
```json
"action": {
  "keywords": ["Action", "Approval", "New Keyword"]
}
```

### Add a new rule
1. Add folder to `folders` array (if needed)
2. Add rule to `rules` array with unique id, name, priority
3. Run: `.\src\Manage-OutlookRules.ps1 -Operation Compare` (review)
4. Run: `.\src\Manage-OutlookRules.ps1 -Operation Deploy` (apply)

### Change noise handling
Edit `rules-config.json` > `settings.noiseAction`:
```json
"noiseAction": "Archive"   // or "Delete"
```

### Backup current rules
```powershell
# Automatic timestamped backup
.\src\Manage-OutlookRules.ps1 -Operation Backup

# Or export to specific path
.\src\Manage-OutlookRules.ps1 -Operation Export -ExportPath ".\my-rules-backup.json"
```

## Management Operations Reference

### Rule Operations
| Operation | Command | Description |
|-----------|---------|-------------|
| List | `-Operation List` | Show all deployed rules |
| Show | `-Operation Show -RuleName "..."` | Show single rule details |
| Export | `-Operation Export` | Export rules to JSON file |
| Backup | `-Operation Backup` | Create timestamped backup in `./backups/` |
| Import | `-Operation Import -ExportPath "..."` | Restore rules from backup file |
| Compare | `-Operation Compare` | Diff deployed vs config |
| Deploy | `-Operation Deploy` | Apply config to Exchange |
| Pull | `-Operation Pull` | Pull deployed rules into config file |
| Enable | `-Operation Enable -RuleName "..."` | Enable a rule |
| Disable | `-Operation Disable -RuleName "..."` | Disable a rule |
| EnableAll | `-Operation EnableAll` | Enable all rules |
| DisableAll | `-Operation DisableAll` | Disable all rules |
| Delete | `-Operation Delete -RuleName "..."` | Remove a rule |
| DeleteAll | `-Operation DeleteAll` | Delete ALL rules (creates backup first) |

### Folder Operations
| Operation | Command | Description |
|-----------|---------|-------------|
| Folders | `-Operation Folders` | List inbox subfolders with counts |
| Stats | `-Operation Stats` | Show mailbox statistics |

### Mailbox Settings Operations
| Operation | Command | Description |
|-----------|---------|-------------|
| OutOfOffice | `-Operation OutOfOffice` | View/set Out-of-Office auto-reply |
| Forwarding | `-Operation Forwarding` | View/set mailbox forwarding |
| JunkMail | `-Operation JunkMail` | View/set safe/blocked sender lists |

### Utility Operations
| Operation | Command | Description |
|-----------|---------|-------------|
| Validate | `-Operation Validate` | Check rules for potential issues |
| Categories | `-Operation Categories` | Show category overview, sync status, and management guide |
| AuditLog | `-Operation AuditLog` | View audit logs (use with -EnableAuditLog on other ops) |

## rules-config.json Structure

```
{
  "settings": { noiseAction, categories }
  "folders": [ { name, description, parent } ]
  "senderLists": { listName: { addresses: [] } }
  "keywordLists": { listName: { keywords: [] } }
  "rules": [ { id, name, priority, conditions, actions } ]
}
```

### Reference Syntax
Use `@` to reference lists:
- `"from": "@senderLists.priority"` → expands to addresses array
- `"subjectContainsWords": "@keywordLists.action"` → expands to keywords array

## Permission Scopes Required

| Service | Permission | Type | Admin Consent |
|---------|------------|------|---------------|
| Microsoft Graph | `Mail.ReadWrite` | Delegated | No |
| Microsoft Graph | `User.Read` | Delegated | No |
| Exchange Online | Implicit via user auth | Delegated | No |

## Testing Commands

```powershell
# Verify connections
Get-MgContext | Select-Object Account, Scopes
Get-ConnectionInformation

# Quick rule check
.\src\Manage-OutlookRules.ps1 -Operation List

# Full rule details
.\src\Manage-OutlookRules.ps1 -Operation Show -RuleName "01 - Priority Senders"

# Compare before deploy
.\src\Manage-OutlookRules.ps1 -Operation Compare
```

## Error Handling Notes

- `Manage-OutlookRules.ps1`: Validates connections before operations
- Deploy operation: Creates folders first, then rules
- Compare operation: Shows what will be created/updated
- Force flag (`-Force`): Skips confirmation prompts

## Do NOT Modify

- Rule priority numbers (01-05, 99) - ordering is intentional
- `stopProcessingRules: true` on Priority Senders rule - prevents cascade
- Rule naming convention (`XX - Name`) - used for pattern matching

## Workflow Patterns

### Initial Setup
```powershell
.\src\Install-Prerequisites.ps1             # Install modules
.\src\Register-OutlookRulesApp.ps1          # Create Azure AD app
.\src\Connect-OutlookRulesApp.ps1           # Authenticate
.\src\Manage-OutlookRules.ps1 -Operation Deploy  # Deploy rules
```

### Ongoing Management
```powershell
.\src\Connect-OutlookRulesApp.ps1           # If session expired
# Edit rules-config.json
.\src\Manage-OutlookRules.ps1 -Operation Compare  # Review changes
.\src\Manage-OutlookRules.ps1 -Operation Deploy   # Apply changes
```

### Backup/Restore
```powershell
# Create timestamped backup
.\src\Manage-OutlookRules.ps1 -Operation Backup
# Creates: ./backups/rules-YYYY-MM-DD_HHMMSS.json

# Restore from backup
.\src\Manage-OutlookRules.ps1 -Operation Import -ExportPath ".\backups\rules-2024-01-15_143022.json"
```

### Troubleshooting
```powershell
# Check mailbox statistics
.\src\Manage-OutlookRules.ps1 -Operation Stats

# Validate rules for issues
.\src\Manage-OutlookRules.ps1 -Operation Validate

# Pull deployed rules to sync config file
.\src\Manage-OutlookRules.ps1 -Operation Pull

# Disable all rules temporarily (for debugging)
.\src\Manage-OutlookRules.ps1 -Operation DisableAll

# Re-enable all rules
.\src\Manage-OutlookRules.ps1 -Operation EnableAll
```

### Multi-Account Management
```powershell
# Connect to personal email account
.\src\Connect-OutlookRulesApp.ps1 -ConfigProfile personal

# Deploy rules to personal account
.\src\Manage-OutlookRules.ps1 -Operation Deploy -ConfigProfile personal

# Connect to work email account
.\src\Connect-OutlookRulesApp.ps1 -ConfigProfile work

# List rules on work account
.\src\Manage-OutlookRules.ps1 -Operation List -ConfigProfile work
```

Profile files:
- `.env.personal` + `rules-config.personal.json` for personal email
- `.env.work` + `rules-config.work.json` for work email
- Each profile needs its own Azure AD app registration in the respective tenant

## Security Considerations

- No secrets stored in scripts (OAuth interactive flow)
- `.env` and `app-config.json` contain only public client ID and tenant ID (both gitignored for safety)
- User must authenticate interactively; no stored tokens
- All operations affect only the authenticated user's mailbox
- DeleteAll operation automatically creates backup before deletion
- Sensitive exports (exported-rules.json) are gitignored
