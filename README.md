# Outlook Rules Manager

Automate Outlook inbox organization with server-side rules and folders using Microsoft Graph and Exchange Online PowerShell.

## Features

- Creates organized inbox subfolders (Priority, Action Required, Metrics, Leadership, Alerts, Low Priority)
- Sets up server-side rules that work across all devices
- Priority sender handling with stop-processing logic
- Keyword-based categorization and flagging
- Noise filtering for newsletters and marketing emails
- Idempotent execution (safe to re-run)

## Prerequisites

- PowerShell 7+ (recommended) or Windows PowerShell 5.1
- Microsoft 365 account with Exchange Online mailbox
- Azure AD tenant for app registration (can use personal tenant)

## Quick Start

### 1. Install Required Modules

```powershell
.\prereqs.ps1
```

### 2. Register Azure AD Application (One-Time)

```powershell
.\Register-OutlookRulesApp.ps1
```

This creates an app registration in your Azure tenant with the necessary permissions. No admin consent required for personal mailbox access.

### 3. Connect to Services

```powershell
.\Connect-OutlookRulesApp.ps1
```

Or use device code flow if browser popups are blocked:

```powershell
.\Connect-OutlookRulesApp.ps1 -UseDeviceCode
```

### 4. Run the Setup

```powershell
.\Setup-OutlookRules.ps1
```

## Folder Structure

| Folder | Purpose |
|--------|---------|
| **Priority** | Messages from VIP senders (manager, skip-level, key collaborators) |
| **Action Required** | Time-sensitive items requiring response |
| **Metrics** | Performance data, KPIs, Connect, QBR content |
| **Leadership** | Executive communications and staff meeting notes |
| **Alerts** | System notifications and digests (auto-read) |
| **Low Priority** | Newsletters, marketing, noise (archived or deleted) |

## Rules Summary

| Rule | Trigger | Actions |
|------|---------|---------|
| **01 - Priority Senders** | From VIP list | Move to Priority, High importance, Stop processing |
| **02 - Action Required** | Subject contains action keywords | Category, High importance, Flag, Move |
| **03 - Connect & Metrics** | Subject/Body contains metrics keywords | Category, Flag, Move |
| **04 - Leadership & Exec** | Subject contains leadership keywords | High importance, Move |
| **05 - Alerts & Notifications** | Subject contains alert keywords | Mark read, Move |
| **99 - Noise Filter** | From noise domains | Archive or Delete |

## Configuration

Edit `rules-config.json` to customize your rules:

### Priority Senders

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

### Keywords

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
  },
  "alerts": {
    "keywords": ["Alert", "Notification", "Digest"]
  }
}
```

### Noise Domains

```json
"senderLists": {
  "noiseDomains": {
    "domains": [
      "news.microsoft.com",
      "events.microsoft.com",
      "linkedin.com",
      "notifications.*"
    ]
  }
}
```

### Noise Handling

```json
"settings": {
  "noiseAction": "Archive"
}
```

Options: `"Archive"` (move to Low Priority folder) or `"Delete"` (permanently delete)

## File Reference

| File | Description |
|------|-------------|
| `prereqs.ps1` | Installs required PowerShell modules |
| `Register-OutlookRulesApp.ps1` | Creates Azure AD app registration |
| `Connect-OutlookRulesApp.ps1` | Authenticates to Graph and Exchange Online |
| `Setup-OutlookRules.ps1` | Creates folders and inbox rules (one-shot setup) |
| `Manage-OutlookRules.ps1` | Full rules management CLI |
| `rules-config.json` | Declarative rule definitions |
| `.env` | Azure AD credentials (ClientId/TenantId) - gitignored |
| `app-config.json` | Backup config in JSON format - gitignored |
| `docs/rules-cheatsheet.md` | Quick reference for rule logic |
| `docs/SDL.md` | Secure Development Lifecycle compliance documentation |
| `docs/SECURITY-QUESTIONNAIRE.md` | Highly Confidential permissions security questionnaire |

## Rules Management

The `Manage-OutlookRules.ps1` script provides comprehensive Outlook management capabilities.

### Operations Quick Reference

| Operation | Description |
|-----------|-------------|
| `List` | View all current inbox rules |
| `Show` | Show details of a specific rule |
| `Export` | Export rules to JSON file |
| `Backup` | Create timestamped backup in `./backups/` |
| `Import` | Restore rules from backup file |
| `Compare` | Compare deployed rules vs config file |
| `Deploy` | Deploy rules from config file |
| `Pull` | Pull deployed rules into config file |
| `Enable` | Enable a specific rule |
| `Disable` | Disable a specific rule |
| `EnableAll` | Enable all rules |
| `DisableAll` | Disable all rules |
| `Delete` | Delete a specific rule |
| `DeleteAll` | Delete ALL rules (creates backup first) |
| `Folders` | List inbox subfolders with counts |
| `Stats` | Show mailbox statistics |
| `Validate` | Check rules for potential issues |
| `Categories` | List categories used in rules |

### List All Rules

```powershell
.\Manage-OutlookRules.ps1 -Operation List
```

### Show Rule Details

```powershell
.\Manage-OutlookRules.ps1 -Operation Show -RuleName "01 - Priority Senders"
```

### Backup & Restore

```powershell
# Create timestamped backup
.\Manage-OutlookRules.ps1 -Operation Backup
# Creates: ./backups/rules-2024-01-15_143022.json

# Restore from backup
.\Manage-OutlookRules.ps1 -Operation Import -ExportPath ".\backups\rules-2024-01-15_143022.json"
```

### Pull Deployed Rules

Sync your config file with what's actually deployed:

```powershell
.\Manage-OutlookRules.ps1 -Operation Pull
# Overwrites rules-config.json with deployed rules
```

### Deploy from Config

```powershell
.\Manage-OutlookRules.ps1 -Operation Deploy

# Skip confirmation
.\Manage-OutlookRules.ps1 -Operation Deploy -Force
```

### Compare Deployed vs Config

```powershell
.\Manage-OutlookRules.ps1 -Operation Compare
```

### Bulk Operations

```powershell
# Disable all rules (e.g., for troubleshooting)
.\Manage-OutlookRules.ps1 -Operation DisableAll

# Re-enable all rules
.\Manage-OutlookRules.ps1 -Operation EnableAll

# Delete ALL rules (creates backup first, requires typing 'DELETE')
.\Manage-OutlookRules.ps1 -Operation DeleteAll
```

### Mailbox Statistics

```powershell
.\Manage-OutlookRules.ps1 -Operation Stats
```

Output:
```
=== Mailbox Statistics ===

Main Folders:
  Inbox                    1,234 items (15 unread)
  Sent Items               2,456 items
  ...

Inbox Subfolders:
  Priority                   45 items (3 unread)
  Action Required            12 items
  ...

Summary:
  Total items:  5,432
  Total unread: 18
  Inbox rules:  6 (6 enabled)
```

### Validate Rules

Check for potential issues:

```powershell
.\Manage-OutlookRules.ps1 -Operation Validate
```

Checks for:
- Disabled rules
- Rules without conditions (matches all emails!)
- Rules without actions
- Duplicate priorities
- Missing target folders

### Enable/Disable Rules

```powershell
# Disable a rule
.\Manage-OutlookRules.ps1 -Operation Disable -RuleName "05 - Alerts & Notifications"

# Enable a rule
.\Manage-OutlookRules.ps1 -Operation Enable -RuleName "05 - Alerts & Notifications"
```

### Delete a Rule

```powershell
.\Manage-OutlookRules.ps1 -Operation Delete -RuleName "99 - Noise Filter"

# Skip confirmation
.\Manage-OutlookRules.ps1 -Operation Delete -RuleName "99 - Noise Filter" -Force
```

## Rules Configuration File

Rules are defined declaratively in `rules-config.json`. This makes rules:
- **Version controllable** - track changes in git
- **Reviewable** - see exactly what will deploy
- **Portable** - share configurations across accounts

### Configuration Structure

```json
{
  "settings": {
    "noiseAction": "Archive",
    "categories": { "action": "Action Required", "metrics": "Metrics" }
  },
  "folders": [
    { "name": "Priority", "description": "VIP senders", "parent": "Inbox" }
  ],
  "senderLists": {
    "priority": {
      "description": "VIP senders",
      "addresses": ["manager@company.com", "skip@company.com"]
    }
  },
  "keywordLists": {
    "action": {
      "description": "Action keywords",
      "keywords": ["Action", "Approval", "Due", "Deadline"]
    }
  },
  "rules": [
    {
      "id": "rule-01",
      "name": "01 - Priority Senders",
      "enabled": true,
      "priority": 1,
      "conditions": { "from": "@senderLists.priority" },
      "actions": {
        "moveToFolder": "Inbox\\Priority",
        "markImportance": "High",
        "stopProcessingRules": true
      }
    }
  ]
}
```

### Reference Syntax

Use `@` references to reuse lists:
- `"from": "@senderLists.priority"` - references the priority sender list
- `"subjectContainsWords": "@keywordLists.action"` - references action keywords
- `"assignCategories": ["@settings.categories.action"]` - references category name

### Adding a New Rule

1. Edit `rules-config.json`
2. Add to the `rules` array:

```json
{
  "id": "rule-06",
  "name": "06 - Team Updates",
  "description": "Route team update emails",
  "enabled": true,
  "priority": 6,
  "conditions": {
    "subjectContainsWords": ["Team Update", "Weekly Sync", "Standup Notes"]
  },
  "actions": {
    "moveToFolder": "Inbox\\Team",
    "markAsRead": false
  }
}
```

3. Add the folder if needed:

```json
{ "name": "Team", "description": "Team communications", "parent": "Inbox" }
```

4. Deploy:

```powershell
.\Manage-OutlookRules.ps1 -Operation Compare  # Review changes
.\Manage-OutlookRules.ps1 -Operation Deploy   # Apply changes
```

### Available Conditions

| Condition | Description | Example |
|-----------|-------------|---------|
| `from` | Sender email addresses | `["user@domain.com"]` |
| `subjectContainsWords` | Words in subject | `["urgent", "asap"]` |
| `bodyContainsWords` | Words in body | `["please review"]` |
| `senderDomainIs` | Sender domain | `["marketing.com"]` |
| `hasAttachment` | Has attachments | `true` |
| `withImportance` | Message importance | `"High"` |

### Available Actions

| Action | Description | Example |
|--------|-------------|---------|
| `moveToFolder` | Move to folder | `"Inbox\\Folder"` |
| `copyToFolder` | Copy to folder | `"Inbox\\Archive"` |
| `deleteMessage` | Delete message | `true` |
| `markAsRead` | Mark as read | `true` |
| `markImportance` | Set importance | `"High"`, `"Low"` |
| `assignCategories` | Apply categories | `["Work", "Important"]` |
| `flagMessage` | Flag for follow-up | `true` |
| `stopProcessingRules` | Stop other rules | `true` |
| `forwardTo` | Forward to address | `["backup@domain.com"]` |

## Alternative: Without App Registration

If your tenant allows user consent for the Microsoft Graph PowerShell app:

```powershell
Connect-MgGraph -Scopes "Mail.ReadWrite" -NoWelcome
Connect-ExchangeOnline
.\Setup-OutlookRules.ps1
```

## Troubleshooting

### "Admin approval required" error

Your tenant has disabled user consent. Options:
1. Use your personal Azure tenant for the app registration
2. Request IT to consent to `Mail.ReadWrite` for Microsoft Graph PowerShell
3. Request an app registration from IT

### "Cannot bind parameter" on New-InboxRule

Ensure you're connected to Exchange Online:
```powershell
Get-ConnectionInformation
```

If not connected:
```powershell
Connect-ExchangeOnline
```

### Rules not applying

1. Verify rules are enabled: `Get-InboxRule | Select-Object Name, Enabled`
2. Check rule priority order: `Get-InboxRule | Select-Object Name, Priority | Sort-Object Priority`
3. Remember Priority Senders rule stops processing - VIP mail won't hit other rules

### Folder not found errors

Run the script again - it creates folders before rules. If issues persist:
```powershell
# Manually verify inbox access
Get-MgUserMailFolder -UserId me -MailFolderId Inbox
```

## Permissions

This solution uses **delegated permissions** only - no admin consent required for your own mailbox:

| Permission | Purpose |
|------------|---------|
| `Mail.ReadWrite` | Create mail folders under Inbox |
| `User.Read` | Basic profile for authentication |
| Exchange Online (implicit) | Manage inbox rules |

## SDL Compliance

For Azure AD admin consent requests, this application follows the **Shadow Org SDL Self-Attestation** process.

| Document | Purpose |
|----------|---------|
| [docs/SDL.md](docs/SDL.md) | SDL compliance, security controls, threat model |
| [docs/SECURITY-QUESTIONNAIRE.md](docs/SECURITY-QUESTIONNAIRE.md) | Highly Confidential permissions questionnaire |

**Key compliance artifacts**:
- Security controls assessment
- API calls documentation
- Code scanning requirements
- Component governance status
- Service Tree requirements

## Updating Rules

To modify rules after initial setup:

1. Edit `rules-config.json` with your changes
2. Run `.\Connect-OutlookRulesApp.ps1` (if session expired)
3. Review changes: `.\Manage-OutlookRules.ps1 -Operation Compare`
4. Apply changes: `.\Manage-OutlookRules.ps1 -Operation Deploy`

Or use the legacy one-shot script:

```powershell
.\Setup-OutlookRules.ps1
```

## Removing Rules

Delete a specific rule:

```powershell
.\Manage-OutlookRules.ps1 -Operation Delete -RuleName "99 - Noise Filter"
```

Delete ALL rules (creates backup first, requires confirmation):

```powershell
.\Manage-OutlookRules.ps1 -Operation DeleteAll
```

Or manually remove rules matching the naming pattern:

```powershell
Get-InboxRule | Where-Object { $_.Name -match "^\d{2} -" } | Remove-InboxRule
```

To remove specific folders (moves contents to Deleted Items):

```powershell
# Get folder ID first
$folders = Get-MgUserMailFolderChildFolder -UserId me -MailFolderId Inbox
$folderId = ($folders | Where-Object { $_.DisplayName -eq "Priority" }).Id
Remove-MgUserMailFolder -UserId me -MailFolderId $folderId
```

## License

MIT License - Use freely, modify as needed.
