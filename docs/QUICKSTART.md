# Quick Start Guide

Get Outlook Rules Manager running in 5 minutes.

---

## Prerequisites

- PowerShell 7+ (or Windows PowerShell 5.1)
- Microsoft 365 account with Exchange Online mailbox
- Azure AD tenant access (for app registration)

---

## Step 1: Clone and Install (2 min)

```powershell
# Clone the repository
git clone https://github.com/fgarofalo56/outlook-rules-manager.git
cd outlook-rules-manager

# Install required PowerShell modules
.\src\Install-Prerequisites.ps1
```

This installs:
- `Microsoft.Graph.Authentication` - For folder management
- `Microsoft.Graph.Mail` - For mailbox operations
- `ExchangeOnlineManagement` - For inbox rules

---

## Step 2: Register Azure AD App (2 min)

### Option A: Create New App Registration (Recommended)

```powershell
# Run the registration script
.\src\Register-OutlookRulesApp.ps1
```

This creates an Azure AD app registration with:
- Required permissions (Mail.ReadWrite, User.Read)
- Device code flow authentication
- App roles for access control

**Output**: Creates `.env` file with ClientId and TenantId.

### Option B: Use Existing App Registration

If you cannot create new app registrations (IT restrictions, existing app available):

```powershell
# Validate and configure an existing app
.\src\Register-OutlookRulesApp.ps1 -UseExisting -ClientId "your-client-id" -TenantId "your-tenant-id"

# Or with auto-fix for common issues
.\src\Register-OutlookRulesApp.ps1 -UseExisting -ClientId "your-client-id" -TenantId "your-tenant-id" -AutoFix
```

**Requirements for existing apps:**
- Public client flow enabled
- Delegated permissions: `Mail.ReadWrite`, `User.Read`
- Redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`

See [EXISTING-APP-SETUP.md](EXISTING-APP-SETUP.md) for detailed instructions.

---

## Step 3: Configure Your Rules (1 min)

```powershell
# Copy the example config
Copy-Item examples/rules-config.example.json rules-config.json
```

Edit `rules-config.json` to customize:

```json
{
  "senderLists": {
    "priority": {
      "addresses": [
        "your.manager@company.com",
        "important.contact@company.com"
      ]
    }
  }
}
```

---

## Step 4: Connect and Deploy

```powershell
# Connect to Microsoft 365
.\src\Connect-OutlookRulesApp.ps1

# Review what will be created
.\src\Manage-OutlookRules.ps1 -Operation Compare

# Deploy rules
.\src\Manage-OutlookRules.ps1 -Operation Deploy
```

---

## Done! Your Rules Are Active

Verify your rules are working:

```powershell
# List all deployed rules
.\src\Manage-OutlookRules.ps1 -Operation List

# Check folder structure
.\src\Manage-OutlookRules.ps1 -Operation Folders

# View mailbox statistics
.\src\Manage-OutlookRules.ps1 -Operation Stats
```

---

## Common Operations

### Add a VIP sender

1. Edit `rules-config.json`:
   ```json
   "priority": {
     "addresses": [
       "existing@company.com",
       "new.vip@company.com"
     ]
   }
   ```

2. Redeploy:
   ```powershell
   .\src\Manage-OutlookRules.ps1 -Operation Deploy
   ```

### Disable all rules temporarily

```powershell
.\src\Manage-OutlookRules.ps1 -Operation DisableAll
```

### Re-enable all rules

```powershell
.\src\Manage-OutlookRules.ps1 -Operation EnableAll
```

### Create a backup

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Backup
# Creates: ./backups/rules-YYYY-MM-DD_HHMMSS.json
```

### Restore from backup

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Import -ExportPath ".\backups\rules-2024-01-15.json"
```

### Set Out-of-Office

```powershell
.\src\Manage-OutlookRules.ps1 -Operation OutOfOffice -OOOEnabled $true `
    -OOOInternal "I'm away from the office." `
    -OOOExternal "I'm currently out of office."
```

---

## Multi-Account Setup

For managing multiple email accounts:

```powershell
# Create profile-specific configs
Copy-Item .env .env.personal
Copy-Item .env .env.work
Copy-Item rules-config.json rules-config.personal.json
Copy-Item rules-config.json rules-config.work.json

# Edit each with appropriate values, then:

# Connect to personal account
.\src\Connect-OutlookRulesApp.ps1 -ConfigProfile personal

# Deploy personal rules
.\src\Manage-OutlookRules.ps1 -Operation Deploy -ConfigProfile personal

# Switch to work account
.\src\Connect-OutlookRulesApp.ps1 -ConfigProfile work

# Deploy work rules
.\src\Manage-OutlookRules.ps1 -Operation Deploy -ConfigProfile work
```

---

## Troubleshooting

### "Not connected to Exchange Online"

```powershell
.\src\Connect-OutlookRulesApp.ps1
```

### App registration issues

If connection fails with authentication errors:

```powershell
# Validate your app registration
.\src\Test-ExistingAppRegistration.ps1 -ClientId "your-client-id" -TenantId "your-tenant-id"

# Attempt to auto-fix common issues
.\src\Test-ExistingAppRegistration.ps1 -ClientId "your-client-id" -TenantId "your-tenant-id" -AutoFix
```

Common issues:
- Public client flow not enabled
- Missing API permissions
- Incorrect redirect URI

### "Config file not found"

```powershell
Copy-Item examples/rules-config.example.json rules-config.json
```

### "Invalid email in senderLists"

Check that all email addresses in `rules-config.json` are valid format.

### View detailed errors

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Deploy -Verbose
```

### Enable audit logging

```powershell
.\src\Manage-OutlookRules.ps1 -Operation Deploy -EnableAuditLog

# View logs
.\src\Manage-OutlookRules.ps1 -Operation AuditLog
```

---

## Security Best Practices

1. **Run pre-commit checks** before any commits:
   ```powershell
   .\scripts\Check-BeforeCommit.ps1
   ```

2. **Enable authorization** for team deployments:
   ```powershell
   .\src\Manage-AppAuthorization.ps1 -Operation Setup
   ```

3. **Use audit logging** for production:
   ```powershell
   .\src\Manage-OutlookRules.ps1 -Operation Deploy -EnableAuditLog
   ```

---

## Next Steps

- **[User Guide](USER-GUIDE.md)** - Complete feature reference
- **[Rules Cheatsheet](rules-cheatsheet.md)** - All rule conditions and actions
- **[Security Guide](SECURITY.md)** - Security model and best practices
- **[Testing Guide](TESTING-GUIDE.md)** - Demo environment setup
- **[Existing App Setup](EXISTING-APP-SETUP.md)** - Use existing Azure AD app registration

---

## All Operations Reference

| Category | Operations |
|----------|-----------|
| **Rules** | List, Show, Deploy, Compare, Pull, Enable, Disable, EnableAll, DisableAll, Delete, DeleteAll |
| **Backup** | Export, Backup, Import |
| **Folders** | Folders, Stats |
| **Mailbox** | OutOfOffice, Forwarding, JunkMail |
| **Utility** | Validate, Categories, AuditLog |

Full syntax:
```powershell
Get-Help .\src\Manage-OutlookRules.ps1 -Full
```
