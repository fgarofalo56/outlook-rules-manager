# Outlook Rules Manager Documentation

Welcome to the Outlook Rules Manager documentation. This folder contains comprehensive guides for implementing, configuring, testing, and maintaining the solution.

## Documentation Index

### Getting Started

| Document | Description |
|----------|-------------|
| [User Guide](USER-GUIDE.md) | Complete implementation guide covering installation, configuration, and daily usage |
| [Testing Guide](TESTING-GUIDE.md) | Demo environment setup, validation procedures, and troubleshooting |

### Reference

| Document | Description |
|----------|-------------|
| [Rules Cheatsheet](rules-cheatsheet.md) | Quick reference for rule conditions and actions |

### Compliance & Security

| Document | Description |
|----------|-------------|
| [Security Policy](SECURITY.md) | Security model, threat analysis, and controls |
| [SDL Compliance](SDL.md) | Security Development Lifecycle self-attestation |
| [Security Questionnaire](SECURITY-QUESTIONNAIRE.md) | Admin consent request documentation |

## Quick Links

### Installation

```powershell
# Install prerequisites
.\Install-Prerequisites.ps1

# Register Azure AD app
.\Register-OutlookRulesApp.ps1

# Connect to services
.\Connect-OutlookRulesApp.ps1
```

### Common Operations

```powershell
# Deploy rules from config
.\Manage-OutlookRules.ps1 -Operation Deploy

# List current rules
.\Manage-OutlookRules.ps1 -Operation List

# Create backup
.\Manage-OutlookRules.ps1 -Operation Backup

# Validate rules
.\Manage-OutlookRules.ps1 -Operation Validate
```

### Configuration Files

| File | Location | Purpose |
|------|----------|---------|
| `.env` | Root (gitignored) | Azure AD Client/Tenant IDs |
| `rules-config.json` | Root (gitignored) | Rule definitions |
| `.env.example` | `examples/` | Template for .env |
| `rules-config.example.json` | `examples/` | Template for rules config |

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    Outlook Rules Manager                     │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐     │
│  │ PowerShell  │───▶│  Microsoft  │───▶│   Outlook   │     │
│  │   Scripts   │    │    Graph    │    │   Folders   │     │
│  └─────────────┘    └─────────────┘    └─────────────┘     │
│         │                                                   │
│         │           ┌─────────────┐    ┌─────────────┐     │
│         └──────────▶│  Exchange   │───▶│   Inbox     │     │
│                     │   Online    │    │   Rules     │     │
│                     └─────────────┘    └─────────────┘     │
│                                                             │
├─────────────────────────────────────────────────────────────┤
│  Authentication: Device Code Flow (SPACE Compliant)         │
│  Permissions: Mail.ReadWrite, User.Read (Delegated)         │
└─────────────────────────────────────────────────────────────┘
```

## API & Module Reference

### Microsoft Graph

| Module | Purpose |
|--------|---------|
| `Microsoft.Graph.Authentication` | OAuth connection |
| `Microsoft.Graph.Mail` | Mail folder operations |

### Exchange Online

| Module | Purpose |
|--------|---------|
| `ExchangeOnlineManagement` | Inbox rule management |

### Azure (Optional)

| Module | Purpose |
|--------|---------|
| `Az.Accounts` | Azure authentication |
| `Az.Resources` | App registration |

## Support

For issues or questions:

1. Check the [Testing Guide](TESTING-GUIDE.md) troubleshooting section
2. Review the [User Guide](USER-GUIDE.md) for configuration help
3. Open an issue on [GitHub](https://github.com/fgarofalo56/outlook-rules-manager/issues)
