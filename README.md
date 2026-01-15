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
.\src\Install-Prerequisites.ps1

# 2. Register Azure AD app (one-time)
.\src\Register-OutlookRulesApp.ps1

# 3. Connect to services
.\src\Connect-OutlookRulesApp.ps1

# 4. Deploy rules
.\src\Manage-OutlookRules.ps1 -Operation Deploy
```

### Using an Existing App Registration

If you cannot create new app registrations (IT restrictions):

```powershell
# Validate and configure an existing app
.\src\Register-OutlookRulesApp.ps1 -UseExisting -ClientId "your-client-id" -TenantId "your-tenant-id"

# With auto-fix for common issues
.\src\Register-OutlookRulesApp.ps1 -UseExisting -ClientId "your-client-id" -TenantId "your-tenant-id" -AutoFix
```

See [docs/EXISTING-APP-SETUP.md](docs/EXISTING-APP-SETUP.md) for detailed requirements.

## Repository Structure

```
outlook-rules-manager/
├── src/                        # Core PowerShell scripts
│   ├── Manage-OutlookRules.ps1     # Main management CLI
│   ├── Manage-AppAuthorization.ps1 # Authorization management
│   ├── Connect-OutlookRulesApp.ps1 # Authentication helper
│   ├── Register-OutlookRulesApp.ps1# Azure AD app registration
│   ├── Install-Prerequisites.ps1   # Module installation
│   └── modules/
│       └── SecurityHelpers.psm1    # Security module
├── scripts/                    # Utility scripts
│   └── Check-BeforeCommit.ps1      # Pre-commit security check
├── tests/                      # Pester tests
├── docs/                       # Documentation
├── examples/                   # Example configuration files
├── .env                        # Azure AD config (gitignored)
└── rules-config.json           # Rule definitions (gitignored)
```

## Core Scripts

| Script | Purpose |
|--------|---------|
| `src/Install-Prerequisites.ps1` | Installs required PowerShell modules |
| `src/Register-OutlookRulesApp.ps1` | Creates or validates Azure AD app registration |
| `src/Test-ExistingAppRegistration.ps1` | Validates existing app registration configuration |
| `src/Connect-OutlookRulesApp.ps1` | Authenticates to Graph and Exchange Online |
| `src/Manage-OutlookRules.ps1` | Full rules management CLI |
| `src/Manage-AppAuthorization.ps1` | User authorization management |

## Management Operations

```powershell
# List all rules
.\src\Manage-OutlookRules.ps1 -Operation List

# Compare config vs deployed
.\src\Manage-OutlookRules.ps1 -Operation Compare

# Deploy from config
.\src\Manage-OutlookRules.ps1 -Operation Deploy

# Create backup
.\src\Manage-OutlookRules.ps1 -Operation Backup

# Show mailbox stats
.\src\Manage-OutlookRules.ps1 -Operation Stats
```

All operations: `List`, `Show`, `Export`, `Backup`, `Import`, `Compare`, `Deploy`, `Pull`, `Enable`, `Disable`, `EnableAll`, `DisableAll`, `Delete`, `DeleteAll`, `Folders`, `Stats`, `Validate`, `Categories`, `AuditLog`

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

- [Quick Start](docs/QUICKSTART.md) - Get running in 5 minutes
- [User Guide](docs/USER-GUIDE.md) - Implementation, configuration, and usage
- [Existing App Setup](docs/EXISTING-APP-SETUP.md) - Use existing Azure AD app registration
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
