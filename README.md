# Outlook Rules Manager

PowerShell-based automation for managing Outlook inbox organization via Microsoft Graph and Exchange Online. Creates server-side rules and folders to automatically triage email.

## Features

- **Server-Side Rules** - Rules work across all devices (desktop, mobile, web)
- **Smart Folders** - Organized inbox subfolders (Priority, Action Required, Metrics, Leadership, Alerts, Low Priority)
- **Priority Handling** - VIP senders get priority treatment with stop-processing logic
- **Declarative Config** - Rules defined in JSON for version control and review
- **Idempotent** - Safe to re-run without creating duplicates

## Prerequisites

| Requirement | Details |
|-------------|---------|
| PowerShell | 7+ (recommended) or Windows PowerShell 5.1 |
| Microsoft 365 | Account with Exchange Online mailbox |
| Azure AD | Tenant for app registration |

## Quick Start

```powershell
# 1. Install required modules
.\Install-Prerequisites.ps1

# 2. Register Azure AD app (one-time)
.\Register-OutlookRulesApp.ps1

# 3. Connect to services
.\Connect-OutlookRulesApp.ps1

# 4. Deploy rules
.\Manage-OutlookRules.ps1 -Operation Deploy
```

## Repository Structure

```
outlook-rules-manager/
├── .github/workflows/          # CI/CD and security scanning
├── docs/                       # Documentation
├── examples/                   # Example configuration files
├── scripts/                    # Utility scripts
├── Install-Prerequisites.ps1   # Module installation
├── Register-OutlookRulesApp.ps1# Azure AD app registration
├── Connect-OutlookRulesApp.ps1 # Authentication helper
├── Setup-OutlookRules.ps1      # One-shot setup (legacy)
├── Manage-OutlookRules.ps1     # Full management CLI
└── rules-config.json           # Rule definitions (gitignored)
```

## Core Scripts

| Script | Purpose |
|--------|---------|
| `Install-Prerequisites.ps1` | Installs required PowerShell modules |
| `Register-OutlookRulesApp.ps1` | Creates Azure AD app registration |
| `Connect-OutlookRulesApp.ps1` | Authenticates to Graph and Exchange Online |
| `Manage-OutlookRules.ps1` | Full rules management CLI |

## Management Operations

```powershell
# List all rules
.\Manage-OutlookRules.ps1 -Operation List

# Compare config vs deployed
.\Manage-OutlookRules.ps1 -Operation Compare

# Deploy from config
.\Manage-OutlookRules.ps1 -Operation Deploy

# Create backup
.\Manage-OutlookRules.ps1 -Operation Backup

# Show mailbox stats
.\Manage-OutlookRules.ps1 -Operation Stats
```

All operations: `List`, `Show`, `Export`, `Backup`, `Import`, `Compare`, `Deploy`, `Pull`, `Enable`, `Disable`, `EnableAll`, `DisableAll`, `Delete`, `DeleteAll`, `Folders`, `Stats`, `Validate`, `Categories`

## Configuration

Rules are defined in `rules-config.json`. Copy the example to get started:

```powershell
Copy-Item examples/rules-config.example.json rules-config.json
```

Edit the file to customize:
- **senderLists** - VIP email addresses
- **keywordLists** - Keywords for categorization
- **rules** - Rule definitions with conditions and actions
- **folders** - Inbox subfolders to create

## Security

- **Delegated permissions only** - No admin consent required for your own mailbox
- **Device code flow** - SPACE-compliant authentication (no localhost redirects)
- **No secrets stored** - Public client, interactive authentication only
- **Gitignored configs** - Sensitive files excluded from repository

### Required Permissions

| Permission | Purpose | Admin Consent |
|------------|---------|:-------------:|
| `Mail.ReadWrite` | Create mail folders | No |
| `User.Read` | Basic profile | No |

## Documentation

See the [docs/](docs/) folder for detailed documentation:

- [User Guide](docs/USER-GUIDE.md) - Implementation, configuration, and usage
- [Testing Guide](docs/TESTING-GUIDE.md) - Demo environment setup and validation
- [SDL Compliance](docs/SDL.md) - Security development lifecycle documentation
- [Security Questionnaire](docs/SECURITY-QUESTIONNAIRE.md) - Admin consent documentation

## Pre-Commit Security

Run before committing to catch credentials/PII:

```powershell
.\scripts\Check-BeforeCommit.ps1
```

## License

MIT License - See [LICENSE](LICENSE) for details.
