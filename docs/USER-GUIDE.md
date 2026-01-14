# User Guide

Complete guide for implementing, configuring, and using the Outlook Rules Manager.

## Table of Contents

- [System Requirements](#system-requirements)
- [Installation](#installation)
- [Azure AD App Registration](#azure-ad-app-registration)
- [Configuration](#configuration)
- [Authentication](#authentication)
- [Managing Rules](#managing-rules)
- [Configuration Reference](#configuration-reference)
- [Workflow Patterns](#workflow-patterns)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)

---

## System Requirements

### Software Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| PowerShell | 5.1 (Windows) | 7.x (Cross-platform) |
| Operating System | Windows 10 | Windows 11 / macOS / Linux |

### Account Requirements

| Requirement | Details |
|-------------|---------|
| Microsoft 365 | Account with Exchange Online mailbox |
| Azure AD | Access to register applications (personal or organizational tenant) |
| Permissions | Ability to consent to `Mail.ReadWrite` and `User.Read` |

### Network Requirements

- Internet access to Microsoft services
- Access to `login.microsoftonline.com`
- Access to `graph.microsoft.com`
- Access to `outlook.office365.com`

---

## Installation

### Step 1: Clone or Download the Repository

```powershell
git clone https://github.com/fgarofalo56/outlook-rules-manager.git
cd outlook-rules-manager
```

### Step 2: Install PowerShell Modules

Run the prerequisite installer:

```powershell
.\Install-Prerequisites.ps1
```

This installs:

| Module | Purpose |
|--------|---------|
| `Microsoft.Graph.Authentication` | OAuth authentication to Graph API |
| `Microsoft.Graph.Mail` | Mail folder operations |
| `ExchangeOnlineManagement` | Inbox rule management |
| `Az.Accounts` | Azure authentication (optional) |
| `Az.Resources` | App registration (optional) |

**Manual Installation** (if needed):

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Mail -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

### Step 3: Verify Installation

```powershell
Get-Module -ListAvailable Microsoft.Graph.Authentication
Get-Module -ListAvailable ExchangeOnlineManagement
```

---

## Azure AD App Registration

You need an Azure AD application to authenticate. Choose one of these methods:

### Option A: Automated Registration (Recommended)

```powershell
.\Register-OutlookRulesApp.ps1
```

This script:
1. Authenticates to Azure
2. Creates an app registration named "Outlook Rules Manager"
3. Configures required permissions (`Mail.ReadWrite`, `User.Read`)
4. Enables public client flow (device code authentication)
5. Saves configuration to `.env` file

### Option B: Manual Registration (Azure Portal)

1. **Navigate to Azure Portal**
   - Go to [portal.azure.com](https://portal.azure.com)
   - Select **Microsoft Entra ID** > **App registrations**

2. **Create New Registration**
   - Click **New registration**
   - **Name**: `Outlook Rules Manager`
   - **Supported account types**: Choose based on your needs
     - Single tenant: `Accounts in this organizational directory only`
     - Multi-tenant: `Accounts in any organizational directory`
   - Click **Register**

3. **Configure Authentication**
   - Go to **Authentication** blade
   - Click **Add a platform** > **Mobile and desktop applications**
   - Select: `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - Enable **Allow public client flows** = **Yes**
   - Click **Save**

4. **Add API Permissions**
   - Go to **API permissions** blade
   - Click **Add a permission** > **Microsoft Graph** > **Delegated permissions**
   - Add:
     - `Mail.ReadWrite`
     - `User.Read`
   - Click **Add permissions**

5. **Record Configuration**
   - Note the **Application (client) ID**
   - Note the **Directory (tenant) ID**

6. **Create Configuration File**

   Create `.env` in the project root:

   ```powershell
   $ClientId = "your-client-id-here"
   $TenantId = "your-tenant-id-here"
   ```

---

## Configuration

### Environment Configuration (.env)

The `.env` file stores your Azure AD app credentials:

```powershell
# Azure AD App Configuration
$ClientId = "00000000-0000-0000-0000-000000000000"
$TenantId = "00000000-0000-0000-0000-000000000000"
```

**Important**: This file is gitignored and should never be committed.

### Rules Configuration (rules-config.json)

Copy the example configuration:

```powershell
Copy-Item examples/rules-config.example.json rules-config.json
```

#### Configuration Structure

```json
{
  "settings": {
    "noiseAction": "Archive",
    "categories": {
      "action": "Action Required",
      "metrics": "Metrics"
    }
  },
  "folders": [...],
  "senderLists": {...},
  "keywordLists": {...},
  "rules": [...]
}
```

#### Settings Section

| Setting | Values | Description |
|---------|--------|-------------|
| `noiseAction` | `Archive` or `Delete` | What to do with noise/newsletter emails |
| `categories.action` | String | Outlook category name for action items |
| `categories.metrics` | String | Outlook category name for metrics |

#### Folders Section

Defines inbox subfolders to create:

```json
"folders": [
  {
    "name": "Priority",
    "description": "VIP senders",
    "parent": "Inbox"
  }
]
```

#### Sender Lists Section

Define groups of email addresses:

```json
"senderLists": {
  "priority": {
    "description": "VIP senders who get priority treatment",
    "addresses": [
      "manager@company.com",
      "executive@company.com"
    ]
  },
  "noiseDomains": {
    "description": "Domains to filter",
    "domains": [
      "newsletter.example.com",
      "marketing.*"
    ]
  }
}
```

#### Keyword Lists Section

Define groups of keywords for matching:

```json
"keywordLists": {
  "action": {
    "description": "Action required keywords",
    "keywords": ["Action", "Approval", "Due", "Deadline"]
  }
}
```

#### Rules Section

Define inbox rules:

```json
{
  "id": "rule-01",
  "name": "01 - Priority Senders",
  "description": "Route VIP emails to Priority folder",
  "enabled": true,
  "priority": 1,
  "conditions": {
    "from": "@senderLists.priority"
  },
  "actions": {
    "moveToFolder": "Inbox\\Priority",
    "markImportance": "High",
    "stopProcessingRules": true
  }
}
```

##### Available Conditions

| Condition | Description | Example |
|-----------|-------------|---------|
| `from` | Sender email addresses | `"@senderLists.priority"` |
| `subjectContainsWords` | Words in subject | `"@keywordLists.action"` |
| `bodyContainsWords` | Words in body | `["urgent", "asap"]` |
| `senderDomainIs` | Sender domain | `"@senderLists.noiseDomains"` |

##### Available Actions

| Action | Description | Example |
|--------|-------------|---------|
| `moveToFolder` | Move to folder | `"Inbox\\Priority"` |
| `markImportance` | Set importance | `"High"`, `"Normal"`, `"Low"` |
| `assignCategories` | Apply categories | `["@settings.categories.action"]` |
| `flagMessage` | Flag for follow-up | `true` |
| `markAsRead` | Mark as read | `true` |
| `deleteMessage` | Delete message | `true` |
| `stopProcessingRules` | Stop rule processing | `true` |

##### Reference Syntax

Use `@` to reference other configuration sections:

| Syntax | Resolves To |
|--------|-------------|
| `@senderLists.priority` | `senderLists.priority.addresses` array |
| `@keywordLists.action` | `keywordLists.action.keywords` array |
| `@settings.categories.action` | `settings.categories.action` value |

---

## Authentication

### Connect to Services

```powershell
.\Connect-OutlookRulesApp.ps1
```

This script:
1. Loads configuration from `.env`
2. Connects to Microsoft Graph (device code flow)
3. Connects to Exchange Online
4. Verifies both connections

### Device Code Flow

1. Run the connect script
2. A code will be displayed (e.g., `ABCD1234`)
3. Open [microsoft.com/devicelogin](https://microsoft.com/devicelogin)
4. Enter the code
5. Sign in with your Microsoft 365 account
6. Consent to permissions (first time only)

### Interactive Flow (Alternative)

If your app has the native client redirect URI configured:

```powershell
.\Connect-OutlookRulesApp.ps1 -Interactive
```

### Verify Connection

```powershell
# Check Graph connection
Get-MgContext

# Check Exchange connection
Get-ConnectionInformation

# Test mailbox access
Get-MgUserMailFolder -UserId me -MailFolderId Inbox
Get-InboxRule
```

---

## Managing Rules

### Operation Reference

| Operation | Command | Description |
|-----------|---------|-------------|
| **List** | `-Operation List` | View all inbox rules |
| **Show** | `-Operation Show -RuleName "..."` | Show rule details |
| **Export** | `-Operation Export` | Export rules to JSON |
| **Backup** | `-Operation Backup` | Create timestamped backup |
| **Import** | `-Operation Import -ExportPath "..."` | Restore from backup |
| **Compare** | `-Operation Compare` | Compare config vs deployed |
| **Deploy** | `-Operation Deploy` | Deploy rules from config |
| **Pull** | `-Operation Pull` | Pull deployed rules to config |
| **Enable** | `-Operation Enable -RuleName "..."` | Enable a rule |
| **Disable** | `-Operation Disable -RuleName "..."` | Disable a rule |
| **EnableAll** | `-Operation EnableAll` | Enable all rules |
| **DisableAll** | `-Operation DisableAll` | Disable all rules |
| **Delete** | `-Operation Delete -RuleName "..."` | Delete a rule |
| **DeleteAll** | `-Operation DeleteAll` | Delete ALL rules |
| **Folders** | `-Operation Folders` | List inbox folders |
| **Stats** | `-Operation Stats` | Show mailbox statistics |
| **Validate** | `-Operation Validate` | Check for rule issues |
| **Categories** | `-Operation Categories` | List available categories |

### Common Workflows

#### Deploy New Rules

```powershell
# Review what will change
.\Manage-OutlookRules.ps1 -Operation Compare

# Deploy changes
.\Manage-OutlookRules.ps1 -Operation Deploy

# Verify deployment
.\Manage-OutlookRules.ps1 -Operation List
```

#### Backup and Restore

```powershell
# Create backup
.\Manage-OutlookRules.ps1 -Operation Backup
# Creates: ./backups/rules-YYYY-MM-DD_HHMMSS.json

# Restore from backup
.\Manage-OutlookRules.ps1 -Operation Import -ExportPath ".\backups\rules-2024-01-15_143022.json"
```

#### Temporarily Disable Rules

```powershell
# Disable all rules for debugging
.\Manage-OutlookRules.ps1 -Operation DisableAll

# Test email flow...

# Re-enable all rules
.\Manage-OutlookRules.ps1 -Operation EnableAll
```

#### Sync Deployed Rules to Config

```powershell
# Pull current rules into config file
.\Manage-OutlookRules.ps1 -Operation Pull
```

---

## Configuration Reference

### Complete rules-config.json Example

```json
{
  "$schema": "./rules-schema.json",
  "_metadata": {
    "description": "Outlook Inbox Rules Configuration",
    "version": "1.0.0",
    "lastModified": "2024-01-15"
  },
  "settings": {
    "noiseAction": "Archive",
    "categories": {
      "action": "Action Required",
      "metrics": "Metrics"
    }
  },
  "folders": [
    { "name": "Priority", "description": "VIP senders", "parent": "Inbox" },
    { "name": "Action Required", "description": "Items needing response", "parent": "Inbox" },
    { "name": "Metrics", "description": "Performance content", "parent": "Inbox" },
    { "name": "Leadership", "description": "Executive comms", "parent": "Inbox" },
    { "name": "Alerts", "description": "System notifications", "parent": "Inbox" },
    { "name": "Low Priority", "description": "Newsletters, noise", "parent": "Inbox" }
  ],
  "senderLists": {
    "priority": {
      "description": "VIP senders",
      "addresses": [
        "manager@company.com",
        "executive@company.com"
      ]
    },
    "noiseDomains": {
      "description": "Newsletter domains",
      "domains": [
        "newsletter.*",
        "marketing.*"
      ]
    }
  },
  "keywordLists": {
    "action": {
      "keywords": ["Action", "Approval", "Due", "Deadline", "Response Needed"]
    },
    "leadership": {
      "keywords": ["Leadership", "Executive", "Staff Meeting", "Review"]
    },
    "metrics": {
      "keywords": ["Performance", "KPI", "QBR", "Scorecard"]
    },
    "alerts": {
      "keywords": ["Alert", "Notification", "Digest"]
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
    },
    {
      "id": "rule-02",
      "name": "02 - Action Required",
      "enabled": true,
      "priority": 2,
      "conditions": { "subjectContainsWords": "@keywordLists.action" },
      "actions": {
        "moveToFolder": "Inbox\\Action Required",
        "assignCategories": ["@settings.categories.action"],
        "markImportance": "High",
        "flagMessage": true
      }
    },
    {
      "id": "rule-99",
      "name": "99 - Noise Filter",
      "enabled": true,
      "priority": 99,
      "conditions": { "senderDomainIs": "@senderLists.noiseDomains" },
      "actions": {
        "moveToFolder": "Inbox\\Low Priority",
        "markAsRead": true
      }
    }
  ]
}
```

---

## Workflow Patterns

### Initial Setup

```powershell
# 1. Install modules
.\Install-Prerequisites.ps1

# 2. Register app (one-time)
.\Register-OutlookRulesApp.ps1

# 3. Create your configuration
Copy-Item examples/rules-config.example.json rules-config.json
# Edit rules-config.json with your senders and preferences

# 4. Connect and deploy
.\Connect-OutlookRulesApp.ps1
.\Manage-OutlookRules.ps1 -Operation Deploy
```

### Daily Operations

```powershell
# Connect (if session expired)
.\Connect-OutlookRulesApp.ps1

# Check current state
.\Manage-OutlookRules.ps1 -Operation List

# Make changes to rules-config.json...

# Preview and deploy
.\Manage-OutlookRules.ps1 -Operation Compare
.\Manage-OutlookRules.ps1 -Operation Deploy
```

### Adding a New Priority Sender

1. Edit `rules-config.json`:

   ```json
   "senderLists": {
     "priority": {
       "addresses": [
         "existing@company.com",
         "new.sender@company.com"
       ]
     }
   }
   ```

2. Deploy the change:

   ```powershell
   .\Manage-OutlookRules.ps1 -Operation Deploy
   ```

### Adding a New Rule

1. Add to `rules-config.json`:

   ```json
   {
     "id": "rule-06",
     "name": "06 - My New Rule",
     "enabled": true,
     "priority": 6,
     "conditions": {
       "subjectContainsWords": ["keyword1", "keyword2"]
     },
     "actions": {
       "moveToFolder": "Inbox\\MyFolder",
       "flagMessage": true
     }
   }
   ```

2. Add the folder if needed:

   ```json
   "folders": [
     { "name": "MyFolder", "parent": "Inbox" }
   ]
   ```

3. Deploy:

   ```powershell
   .\Manage-OutlookRules.ps1 -Operation Compare
   .\Manage-OutlookRules.ps1 -Operation Deploy
   ```

---

## Best Practices

### Rule Priority Order

- **1-10**: High priority rules (VIP senders, urgent items)
- **11-50**: Normal processing rules
- **51-98**: Lower priority rules
- **99**: Catch-all/noise filter (last)

### Use Stop Processing Wisely

The `stopProcessingRules: true` action prevents subsequent rules from running. Use it for:
- VIP senders (they should only go to Priority folder)
- Any rule where you want exclusive handling

### Backup Before Major Changes

```powershell
.\Manage-OutlookRules.ps1 -Operation Backup
# Then make your changes
```

### Validate Configuration

```powershell
.\Manage-OutlookRules.ps1 -Operation Validate
```

### Security Checklist

Before committing changes:

```powershell
.\scripts\Check-BeforeCommit.ps1
```

---

## Troubleshooting

### "Admin approval required"

Your tenant has restricted user consent.

**Solutions**:
1. Use a personal Azure tenant for app registration
2. Request IT to consent to permissions
3. Request an IT-managed app registration

### "Cannot bind parameter"

Exchange Online connection issue.

```powershell
# Check connection
Get-ConnectionInformation

# Reconnect if needed
Connect-ExchangeOnline
```

### Rules Not Applying

1. Verify rules are enabled:
   ```powershell
   Get-InboxRule | Select-Object Name, Enabled
   ```

2. Check priority order:
   ```powershell
   Get-InboxRule | Select-Object Name, Priority | Sort-Object Priority
   ```

3. Remember: `stopProcessingRules` on Priority Senders prevents other rules from running on VIP mail

### Device Code Authentication Fails

1. Verify app has **Allow public client flows** = **Yes**
2. Check redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`
3. Ensure user has access to the tenant

### Graph Connection Fails

```powershell
# Check current context
Get-MgContext

# Disconnect and reconnect
Disconnect-MgGraph
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Scopes "Mail.ReadWrite","User.Read" -UseDeviceCode
```

### Folders Not Created

Ensure Graph connection has `Mail.ReadWrite` scope:

```powershell
(Get-MgContext).Scopes
```

If missing, reconnect with proper scopes.
