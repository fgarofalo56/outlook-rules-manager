# Outlook Rules Quick Reference

## Folders (under Inbox)

| Folder | Purpose |
|--------|---------|
| **Priority** | VIP senders: manager, skip-level, key collaborators |
| **Action Required** | Time-sensitive items; auto-category: Action Required |
| **Metrics** | Connect, ACR, performance/KPI/attainment; auto-category: Metrics |
| **Leadership** | Leadership/executive/staff review comms |
| **Alerts** | Alerts/notifications/digests (marked read) |
| **Low Priority** | Newsletters/marketing (archived or deleted per preference) |

## Rules (Server-Side, Priority Order)

### 01 - Priority Senders
**IF** From is any of: Blake Badolato, Dan Coleman, PJ Kemp, Stephanie Joshi, Chris Peacock, Anthony Puca, Valarie Warburton, Pete Nash, Stephen Cyphers
**THEN** Move to Priority, mark High importance, **Stop processing**

### 02 - Action Required
**IF** Subject contains: Action, Approval, Response Needed, Due, Deadline, Sign-off, Decision Required
**THEN** Assign category "Action Required", mark High importance, Flag, move to Action Required

### 03 - Connect & Metrics
**IF** Subject/Body contains: Connect, ACR, Performance, Impact, KPI, Attainment, ADS, Azure Consumption, QBR, Scorecard
**THEN** Assign category "Metrics", Flag, move to Metrics

### 04 - Leadership & Exec
**IF** Subject contains: Leadership, Executive, LT, Review, Staff Meeting
**THEN** Mark High importance, move to Leadership

### 05 - Alerts & Notifications
**IF** Subject contains: Alert, Notification, Digest
**THEN** Mark as Read, move to Alerts

### 99 - Noise Filter
**IF** Sender domain matches: news.microsoft.com, events.microsoft.com, linkedin.com, notifications.*, mailer.*
**THEN** Move to Low Priority (or delete based on `settings.noiseAction` in rules-config.json)

## Quick Management

### Add new VIP sender
```powershell
# 1. Edit rules-config.json > senderLists.priority.addresses
# 2. Review and deploy changes:
.\Manage-OutlookRules.ps1 -Operation Compare
.\Manage-OutlookRules.ps1 -Operation Deploy
```

### Check current rules
```powershell
.\Manage-OutlookRules.ps1 -Operation List
```

### View rule details
```powershell
.\Manage-OutlookRules.ps1 -Operation Show -RuleName "01 - Priority Senders"
```

### View folder structure with counts
```powershell
.\Manage-OutlookRules.ps1 -Operation Folders
```

### Disable a rule temporarily
```powershell
.\Manage-OutlookRules.ps1 -Operation Disable -RuleName "02 - Action Required"
```

### Re-enable a rule
```powershell
.\Manage-OutlookRules.ps1 -Operation Enable -RuleName "02 - Action Required"
```

### Backup before making changes
```powershell
.\Manage-OutlookRules.ps1 -Operation Backup
```

### Check mailbox stats
```powershell
.\Manage-OutlookRules.ps1 -Operation Stats
```

### Validate rules for issues
```powershell
.\Manage-OutlookRules.ps1 -Operation Validate
```

## Tips

- **Priority Senders stays at top** - it stops all other rule processing for VIPs
- Use **Search Folders** in Outlook for "Unread" and "Flagged" at-a-glance views
- For **historical cleanup**: search by domain/subject, select all, move to target folders
- Rules are **server-side** - they apply even when Outlook is closed
- **Always backup first** before making significant changes
- Use `-Force` flag to skip confirmation prompts
