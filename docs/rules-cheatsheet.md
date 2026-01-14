# Outlook Rules Quick Reference

## Supported Rule Conditions

### Sender Conditions
| Config Property | Description | Example |
|-----------------|-------------|---------|
| `from` | Exact sender email addresses | `["boss@company.com"]` |
| `fromAddressContainsWords` | Words in sender address | `["newsletter", "noreply"]` |
| `senderDomainIs` | Sender domain matches | `["linkedin.com", "*.microsoft.com"]` |

### Recipient Conditions
| Config Property | Description | Example |
|-----------------|-------------|---------|
| `sentTo` | Message sent to specific recipients | `["team@company.com"]` |
| `recipientAddressContainsWords` | Words in recipient address | `["all-hands"]` |
| `myNameInToBox` | You are in the To field | `true` |
| `myNameInCcBox` | You are in the Cc field | `true` |
| `myNameInToOrCcBox` | You are in To or Cc | `true` |
| `myNameNotInToBox` | You are NOT in To field | `true` |
| `sentOnlyToMe` | Message sent only to you | `true` |

### Content Conditions
| Config Property | Description | Example |
|-----------------|-------------|---------|
| `subjectContainsWords` | Words in subject line | `["Action", "Urgent"]` |
| `bodyContainsWords` | Words in message body | `["deadline", "required"]` |
| `subjectOrBodyContainsWords` | Words in subject OR body | `["meeting", "review"]` |
| `headerContainsWords` | Words in email headers | `["X-Priority: 1"]` |

### Message Property Conditions
| Config Property | Description | Values |
|-----------------|-------------|--------|
| `hasAttachment` | Message has attachments | `true` / `false` |
| `messageTypeMatches` | Message type | `AutomaticReply`, `Calendaring`, `CalendaringResponse`, `Encrypted`, `Voicemail`, `ReadReceipt`, `NonDeliveryReport` |
| `withImportance` | Message importance | `High`, `Normal`, `Low` |
| `withSensitivity` | Message sensitivity | `Normal`, `Personal`, `Private`, `CompanyConfidential` |
| `flaggedForAction` | Flag status | `Any`, `Call`, `FollowUp`, `ForYourInformation`, `Reply`, `Review` |
| `hasClassification` | Message classification | Classification name |

### Size & Date Conditions
| Config Property | Description | Example |
|-----------------|-------------|---------|
| `withinSizeRangeMinimum` | Minimum message size (bytes) | `1048576` (1MB) |
| `withinSizeRangeMaximum` | Maximum message size (bytes) | `10485760` (10MB) |
| `receivedAfterDate` | Received after date | `"2024-01-01"` |
| `receivedBeforeDate` | Received before date | `"2024-12-31"` |

---

## Supported Rule Actions

### Move/Copy Actions
| Config Property | Description | Example |
|-----------------|-------------|---------|
| `moveToFolder` | Move to folder | `"Inbox\\Priority"` |
| `copyToFolder` | Copy to folder | `"Inbox\\Archive"` |
| `deleteMessage` | Delete (to Deleted Items) | `true` |
| `softDeleteMessage` | Soft delete (recoverable) | `true` |

### Mark Actions
| Config Property | Description | Example |
|-----------------|-------------|---------|
| `markAsRead` | Mark as read | `true` |
| `markImportance` | Set importance | `High`, `Normal`, `Low` |
| `flagMessage` | Flag for follow-up | `true` |
| `pinMessage` | Pin to top of folder | `true` |

### Category Actions
| Config Property | Description | Example |
|-----------------|-------------|---------|
| `assignCategories` | Apply custom categories | `["Action Required"]` |
| `applySystemCategory` | Apply system category | `Bills`, `Flight`, `Travel`, `Package`, `Shopping` |
| `deleteSystemCategory` | Remove system category | `Bills` |

### Forward/Redirect Actions
| Config Property | Description | Example |
|-----------------|-------------|---------|
| `forwardTo` | Forward to addresses | `["backup@company.com"]` |
| `redirectTo` | Redirect to addresses | `["delegate@company.com"]` |
| `forwardAsAttachmentTo` | Forward as attachment | `["archive@company.com"]` |

### Processing Control
| Config Property | Description | Example |
|-----------------|-------------|---------|
| `stopProcessingRules` | Stop other rules | `true` |

---

## Rule Configuration Example

```json
{
  "rules": [
    {
      "id": "rule-attachments",
      "name": "Large Attachments",
      "description": "Move emails with large attachments",
      "enabled": true,
      "priority": 10,
      "conditions": {
        "hasAttachment": true,
        "withinSizeRangeMinimum": 5242880
      },
      "actions": {
        "moveToFolder": "Inbox\\Large Files",
        "applySystemCategory": "Package"
      }
    },
    {
      "id": "rule-calendar",
      "name": "Calendar Invites",
      "description": "Mark calendar invites as high importance",
      "enabled": true,
      "priority": 15,
      "conditions": {
        "messageTypeMatches": "Calendaring",
        "myNameInToBox": true
      },
      "actions": {
        "markImportance": "High",
        "pinMessage": true
      }
    }
  ]
}
```

---

## Operations Reference

### Rule Operations
```powershell
# List all rules
.\src\Manage-OutlookRules.ps1 -Operation List

# Show rule details
.\src\Manage-OutlookRules.ps1 -Operation Show -RuleName "01 - Priority Senders"

# Deploy rules from config
.\src\Manage-OutlookRules.ps1 -Operation Deploy

# Compare deployed vs config
.\src\Manage-OutlookRules.ps1 -Operation Compare

# Pull deployed rules into config
.\src\Manage-OutlookRules.ps1 -Operation Pull

# Backup rules
.\src\Manage-OutlookRules.ps1 -Operation Backup

# Enable/Disable rules
.\src\Manage-OutlookRules.ps1 -Operation Enable -RuleName "MyRule"
.\src\Manage-OutlookRules.ps1 -Operation Disable -RuleName "MyRule"
.\src\Manage-OutlookRules.ps1 -Operation EnableAll
.\src\Manage-OutlookRules.ps1 -Operation DisableAll
```

### Mailbox Settings Operations
```powershell
# View Out-of-Office settings
.\src\Manage-OutlookRules.ps1 -Operation OutOfOffice

# Enable Out-of-Office
.\src\Manage-OutlookRules.ps1 -Operation OutOfOffice -OOOEnabled $true `
    -OOOInternal "I'm away from the office" `
    -OOOExternal "I'm currently out of office"

# Schedule Out-of-Office
.\src\Manage-OutlookRules.ps1 -Operation OutOfOffice -OOOEnabled $true `
    -OOOStartDate "2024-12-23" -OOOEndDate "2024-12-27" `
    -OOOInternal "Away for the holidays"

# Disable Out-of-Office
.\src\Manage-OutlookRules.ps1 -Operation OutOfOffice -OOOEnabled $false

# View forwarding settings
.\src\Manage-OutlookRules.ps1 -Operation Forwarding

# Enable forwarding
.\src\Manage-OutlookRules.ps1 -Operation Forwarding `
    -ForwardingAddress "backup@company.com" `
    -ForwardingEnabled $true `
    -DeliverToMailbox $true  # Keep a copy

# Disable forwarding
.\src\Manage-OutlookRules.ps1 -Operation Forwarding -ForwardingEnabled $false

# View junk mail settings
.\src\Manage-OutlookRules.ps1 -Operation JunkMail

# Add safe senders
.\src\Manage-OutlookRules.ps1 -Operation JunkMail `
    -SafeSenders "trusted@company.com","partner.com"

# Add blocked senders
.\src\Manage-OutlookRules.ps1 -Operation JunkMail `
    -BlockedSenders "spam@example.com"
```

### Utility Operations
```powershell
# Validate rules for issues
.\src\Manage-OutlookRules.ps1 -Operation Validate

# View mailbox statistics
.\src\Manage-OutlookRules.ps1 -Operation Stats

# List inbox folders
.\src\Manage-OutlookRules.ps1 -Operation Folders

# List categories used
.\src\Manage-OutlookRules.ps1 -Operation Categories
```

---

## Tips

- **Priority Senders stays at top** - it stops all other rule processing for VIPs
- Use **Search Folders** in Outlook for "Unread" and "Flagged" at-a-glance views
- For **historical cleanup**: search by domain/subject, select all, move to target folders
- Rules are **server-side** - they apply even when Outlook is closed
- **Always backup first** before making significant changes
- Use `-Force` flag to skip confirmation prompts
- Use `-Verbose` flag for detailed diagnostic output
- Use `-EnableAuditLog` to log operations for compliance
