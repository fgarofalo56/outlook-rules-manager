# Outlook Rules Manager - Claude Code Instructions

## Project Overview

PowerShell-based automation for managing Outlook inbox organization via Microsoft Graph and Exchange Online. Creates server-side rules and folders to automatically triage email.

## Architecture

```
outlook/
├── prereqs.ps1                    # Module installation script
├── Register-OutlookRulesApp.ps1   # Azure AD app registration (run once)
├── Connect-OutlookRulesApp.ps1    # Authentication helper
├── Setup-OutlookRules.ps1         # One-shot rules/folders creation
├── Manage-OutlookRules.ps1        # Full management CLI (list/export/deploy/etc)
├── rules-config.json              # Declarative rule definitions
├── .env                           # Azure AD credentials (gitignored)
├── app-config.json                # Backup config JSON (gitignored)
├── .gitignore                     # Protects .env and generated files
└── docs/
    └── rules-cheatsheet.md        # Quick reference for rule logic
```

## Key Files

| File | Purpose | When to Modify |
|------|---------|----------------|
| `rules-config.json` | All rule definitions | Adding/changing rules |
| `Manage-OutlookRules.ps1` | Management operations | Adding new operations |
| `Setup-OutlookRules.ps1` | Legacy one-shot setup | Rarely - use config instead |

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
| `app-config.json` | Azure AD config backup (JSON format) | Yes |
| `rules-config.json` | Rule definitions, sender lists, keywords | No |
| `exported-rules.json` | Backup of deployed rules | Yes |

## Common Tasks

### Add a new priority sender
Edit `rules-config.json` > `senderLists.priority.addresses`:
```json
"addresses": [
  "existing@email.com",
  "new.sender@email.com"
]
```
Then: `.\Manage-OutlookRules.ps1 -Operation Deploy`

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
3. Run: `.\Manage-OutlookRules.ps1 -Operation Compare` (review)
4. Run: `.\Manage-OutlookRules.ps1 -Operation Deploy` (apply)

### Change noise handling
Edit `rules-config.json` > `settings.noiseAction`:
```json
"noiseAction": "Archive"   // or "Delete"
```

### Backup current rules
```powershell
# Automatic timestamped backup
.\Manage-OutlookRules.ps1 -Operation Backup

# Or export to specific path
.\Manage-OutlookRules.ps1 -Operation Export -ExportPath ".\my-rules-backup.json"
```

## Management Operations Reference

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
| Folders | `-Operation Folders` | List inbox subfolders with counts |
| Stats | `-Operation Stats` | Show mailbox statistics |
| Validate | `-Operation Validate` | Check rules for potential issues |
| Categories | `-Operation Categories` | List categories used in rules |

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
.\Manage-OutlookRules.ps1 -Operation List

# Full rule details
.\Manage-OutlookRules.ps1 -Operation Show -RuleName "01 - Priority Senders"

# Compare before deploy
.\Manage-OutlookRules.ps1 -Operation Compare
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
.\prereqs.ps1                           # Install modules
.\Register-OutlookRulesApp.ps1          # Create Azure AD app
.\Connect-OutlookRulesApp.ps1           # Authenticate
.\Manage-OutlookRules.ps1 -Operation Deploy  # Deploy rules
```

### Ongoing Management
```powershell
.\Connect-OutlookRulesApp.ps1           # If session expired
# Edit rules-config.json
.\Manage-OutlookRules.ps1 -Operation Compare  # Review changes
.\Manage-OutlookRules.ps1 -Operation Deploy   # Apply changes
```

### Backup/Restore
```powershell
# Create timestamped backup
.\Manage-OutlookRules.ps1 -Operation Backup
# Creates: ./backups/rules-YYYY-MM-DD_HHMMSS.json

# Restore from backup
.\Manage-OutlookRules.ps1 -Operation Import -ExportPath ".\backups\rules-2024-01-15_143022.json"
```

### Troubleshooting
```powershell
# Check mailbox statistics
.\Manage-OutlookRules.ps1 -Operation Stats

# Validate rules for issues
.\Manage-OutlookRules.ps1 -Operation Validate

# Pull deployed rules to sync config file
.\Manage-OutlookRules.ps1 -Operation Pull

# Disable all rules temporarily (for debugging)
.\Manage-OutlookRules.ps1 -Operation DisableAll

# Re-enable all rules
.\Manage-OutlookRules.ps1 -Operation EnableAll
```

## Security Considerations

- No secrets stored in scripts (OAuth interactive flow)
- `.env` and `app-config.json` contain only public client ID and tenant ID (both gitignored for safety)
- User must authenticate interactively; no stored tokens
- All operations affect only the authenticated user's mailbox
- DeleteAll operation automatically creates backup before deletion
- Sensitive exports (exported-rules.json) are gitignored
