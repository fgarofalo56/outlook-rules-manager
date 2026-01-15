# Outlook Rules Manager - Demo Walkthrough

This walkthrough demonstrates the key features of the Outlook Rules Manager with example commands and expected outputs.

---

## Prerequisites Check

Before starting, verify your environment is set up correctly:

```powershell
# Check PowerShell version (requires 7+)
$PSVersionTable.PSVersion

# Verify required modules
Get-Module -ListAvailable Microsoft.Graph.Authentication, ExchangeOnlineManagement
```

---

## Step 1: Initial Setup

### Install Prerequisites

```powershell
.\src\Install-Prerequisites.ps1
```

**Expected Output:**
```
Checking for Microsoft.Graph.Authentication... Installing...
Checking for Microsoft.Graph.Mail... Installing...
Checking for ExchangeOnlineManagement... Installing...
All prerequisites installed successfully!
```

### Register Azure AD Application

```powershell
.\src\Register-OutlookRulesApp.ps1
```

**Expected Output:**
```
Creating Azure AD application: Outlook Rules Manager
Adding API permissions...
Enabling public client flow...
Creating App Roles...
  - OutlookRules.Admin (Administrator)
  - OutlookRules.User (Standard User)
Enabling User Assignment Required...

Application registered successfully!
  App Name:   Outlook Rules Manager
  Client ID:  xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
  Tenant ID:  xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

Configuration saved to: .env
```

---

## Step 2: Connect to Your Mailbox

```powershell
.\src\Connect-OutlookRulesApp.ps1
```

**Expected Output:**
```
Loading configuration from .env...
Connecting to Microsoft Graph...

To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code XXXXXXXXX to authenticate.

Connected successfully!
  User: user@example.com
  Scopes: Mail.ReadWrite, User.Read

Connecting to Exchange Online...
Connected to Exchange Online!
```

---

## Step 3: View Existing Rules

### List All Rules

```powershell
.\src\Manage-OutlookRules.ps1 -Operation List
```

**Expected Output:**
```
=== Inbox Rules ===

Priority  Name                      Enabled  Actions
--------  ----                      -------  -------
1         01 - Priority Senders     True     Move to Inbox\Priority, Stop
2         02 - Action Required      True     Move to Inbox\Action, Category: Action Required
3         03 - Metrics & Reports    True     Move to Inbox\Metrics
4         04 - Meeting Responses    True     Mark as Read, Move to Inbox\Noise
5         05 - Automated Noise      True     Archive

Total: 5 rules
```

### Show Rule Details

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Show -RuleName "01 - Priority Senders"
```

**Expected Output:**
```
=== Rule: 01 - Priority Senders ===

Property              Value
--------              -----
Name                  01 - Priority Senders
Priority              1
Enabled               True
StopProcessingRules   True

Conditions:
  From                vip@example.com, boss@example.com

Actions:
  MoveToFolder        Inbox\Priority
  StopProcessingRules True
```

---

## Step 4: Configure Rules

### Edit rules-config.json

The configuration file defines your rules declaratively:

```json
{
  "settings": {
    "noiseAction": "Archive",
    "categories": {
      "action": "Action Required",
      "metrics": "Metrics"
    }
  },
  "folders": [
    { "name": "Priority", "parent": "Inbox" },
    { "name": "Action", "parent": "Inbox" },
    { "name": "Metrics", "parent": "Inbox" },
    { "name": "Noise", "parent": "Inbox" }
  ],
  "senderLists": {
    "priority": {
      "description": "VIP senders that bypass all other rules",
      "addresses": ["boss@example.com", "vip@example.com"]
    },
    "noise": {
      "description": "Automated notifications",
      "domains": ["notifications.example.com", "*.noreply.com"]
    }
  },
  "rules": [
    {
      "id": "priority-senders",
      "name": "01 - Priority Senders",
      "priority": 1,
      "conditions": {
        "from": "@senderLists.priority"
      },
      "actions": {
        "moveToFolder": "Inbox\\Priority",
        "stopProcessingRules": true
      }
    }
  ]
}
```

### Compare Configuration vs Deployed

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Compare
```

**Expected Output:**
```
=== Comparing Config vs Deployed Rules ===

Rule                      Config    Deployed   Status
----                      ------    --------   ------
01 - Priority Senders     Yes       Yes        In Sync
02 - Action Required      Yes       Yes        Modified (conditions differ)
03 - Metrics & Reports    Yes       No         New (will be created)
05 - Automated Noise      No        Yes        Orphaned (not in config)

Summary:
  In Sync:  1
  Modified: 1
  New:      1
  Orphaned: 1

Run with -Operation Deploy to apply changes.
```

### Deploy Configuration

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Deploy
```

**Expected Output:**
```
=== Deploying Rules from Configuration ===

Creating folders...
  [+] Inbox\Priority (exists)
  [+] Inbox\Action (exists)
  [+] Inbox\Metrics (created)
  [+] Inbox\Noise (exists)

Deploying rules...
  [=] 01 - Priority Senders (unchanged)
  [~] 02 - Action Required (updated)
  [+] 03 - Metrics & Reports (created)

Deployment complete!
  Created: 1
  Updated: 1
  Unchanged: 1
```

---

## Step 5: Backup and Restore

### Create Backup

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Backup
```

**Expected Output:**
```
Creating backup...
Backup saved to: ./backups/rules-2026-01-14_153022.json
  Rules: 5
  Size: 4.2 KB
```

### Restore from Backup

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Import -ExportPath "./backups/rules-2026-01-14_153022.json"
```

**Expected Output:**
```
Importing rules from: ./backups/rules-2026-01-14_153022.json
  [+] 01 - Priority Senders (restored)
  [+] 02 - Action Required (restored)
  [+] 03 - Metrics & Reports (restored)

Import complete! 3 rules restored.
```

---

## Step 6: Mailbox Settings

### View Out-of-Office Settings

```powershell
.\src\Manage-OutlookRules.ps1 -Operation OutOfOffice
```

**Expected Output:**
```
=== Out of Office Settings ===

Status:           Disabled
Internal Message: (not set)
External Message: (not set)
Scheduled:        No
```

### Enable Out-of-Office

```powershell
.\src\Manage-OutlookRules.ps1 -Operation OutOfOffice `
    -OOOEnabled $true `
    -OOOInternal "I'm currently out of office and will respond when I return." `
    -OOOExternal "Thank you for your email. I'm out of office with limited access."
```

**Expected Output:**
```
=== Updating Out of Office ===

Setting internal message...
Setting external message...
Enabling auto-reply...

Out of Office enabled successfully!
```

### View Forwarding Settings

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Forwarding
```

**Expected Output:**
```
=== Forwarding Settings ===

Forwarding Enabled:    No
Forwarding Address:    (none)
Deliver to Mailbox:    Yes
```

---

## Step 7: Validation and Troubleshooting

### Validate Rules

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Validate
```

**Expected Output:**
```
=== Validating Rules ===

Checking for common issues...

[OK] No duplicate rule names
[OK] No duplicate priorities
[OK] All referenced folders exist
[WARN] Rule "Forward Important" has forwarding action - verify this is intentional
[OK] No rules targeting deleted folders

Validation complete: 0 errors, 1 warning
```

### View Mailbox Statistics

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Stats
```

**Expected Output:**
```
=== Mailbox Statistics ===

Folder              Count    Unread   Size
------              -----    ------   ----
Inbox               1,234    45       125 MB
  Priority          89       3        12 MB
  Action            156      12       28 MB
  Metrics           423      0        45 MB
  Noise             566      30       40 MB
Sent Items          2,456    0        89 MB
Deleted Items       123      0        15 MB

Total Items: 3,813
Total Size: 229 MB
```

---

## Step 8: Audit Logging

### Enable Audit Logging

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Deploy -EnableAuditLog
```

**Expected Output:**
```
[AUDIT] 2026-01-14 15:30:22 - Deploy operation started
[AUDIT] 2026-01-14 15:30:23 - Created folder: Inbox\Metrics
[AUDIT] 2026-01-14 15:30:24 - Updated rule: 02 - Action Required
[AUDIT] 2026-01-14 15:30:25 - Deploy operation completed

Audit log saved to: ./logs/audit-2026-01-14.json
```

### View Audit Logs

```powershell
.\src\Manage-OutlookRules.ps1 -Operation AuditLog
```

**Expected Output:**
```
=== Audit Log (Last 7 Days) ===

Timestamp            Operation    Details                      User
---------            ---------    -------                      ----
2026-01-14 15:30:22  Deploy       Started deployment           user@example.com
2026-01-14 15:30:23  CreateFolder Inbox\Metrics                user@example.com
2026-01-14 15:30:24  UpdateRule   02 - Action Required         user@example.com
2026-01-14 15:30:25  Deploy       Completed successfully       user@example.com
2026-01-13 09:15:00  Backup       rules-2026-01-13.json        user@example.com

Total entries: 5
```

---

## Quick Reference

| Task | Command |
|------|---------|
| Connect | `.\src\Connect-OutlookRulesApp.ps1` |
| List rules | `.\src\Manage-OutlookRules.ps1 -Operation List` |
| Deploy config | `.\src\Manage-OutlookRules.ps1 -Operation Deploy` |
| Compare | `.\src\Manage-OutlookRules.ps1 -Operation Compare` |
| Backup | `.\src\Manage-OutlookRules.ps1 -Operation Backup` |
| Validate | `.\src\Manage-OutlookRules.ps1 -Operation Validate` |
| Statistics | `.\src\Manage-OutlookRules.ps1 -Operation Stats` |

For more details, see:
- [QUICKSTART.md](QUICKSTART.md) - 5-minute getting started guide
- [USER-GUIDE.md](USER-GUIDE.md) - Complete user documentation
- [rules-cheatsheet.md](rules-cheatsheet.md) - Rule conditions and actions reference
