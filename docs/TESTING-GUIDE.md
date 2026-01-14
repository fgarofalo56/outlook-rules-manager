# Testing Guide

Complete guide for setting up demo environments, testing procedures, and validation workflows.

## Table of Contents

- [Overview](#overview)
- [Demo Environment Setup](#demo-environment-setup)
- [Development Environment](#development-environment)
- [Testing Procedures](#testing-procedures)
- [Validation Checklist](#validation-checklist)
- [Troubleshooting Tests](#troubleshooting-tests)
- [CI/CD Integration](#cicd-integration)

---

## Overview

This guide covers three testing scenarios:

| Environment | Purpose | Tenant |
|-------------|---------|--------|
| **Demo** | End-to-end testing with real mailbox | Separate demo tenant |
| **Development** | Script development and debugging | Personal or dev tenant |
| **Production** | Live deployment | Work/production tenant |

**Recommendation**: Always test in a demo environment before deploying to production.

---

## Demo Environment Setup

### Prerequisites

| Requirement | Details |
|-------------|---------|
| Demo Tenant | Azure AD tenant with Exchange Online |
| Test User | User account with mailbox in demo tenant |
| App Registration | Separate app registration in demo tenant |

### Step 1: Create Demo App Registration

#### Option A: Automated

```powershell
# Connect to demo tenant
Connect-AzAccount -TenantId "<demo-tenant-id>"

# Register app with custom name
.\Register-OutlookRulesApp.ps1 -AppName "Outlook Rules Manager - Demo"
```

#### Option B: Manual (Azure Portal)

1. **Go to Azure Portal**
   - Navigate to [portal.azure.com](https://portal.azure.com)
   - Switch to your demo tenant
   - Go to **Microsoft Entra ID** > **App registrations**

2. **Create Registration**
   - Click **New registration**
   - **Name**: `Outlook Rules Manager - Demo`
   - **Supported account types**: `Accounts in this organizational directory only`
   - Click **Register**

3. **Configure Authentication**

   In the **Authentication** blade:

   | Setting | Value |
   |---------|-------|
   | Platform | Mobile and desktop applications |
   | Redirect URI | `https://login.microsoftonline.com/common/oauth2/nativeclient` |
   | Allow public client flows | **Yes** |

4. **Add API Permissions**

   In the **API permissions** blade:

   | API | Permission | Type |
   |-----|------------|------|
   | Microsoft Graph | `Mail.ReadWrite` | Delegated |
   | Microsoft Graph | `User.Read` | Delegated |

5. **Record IDs**
   - **Application (client) ID**: Copy this value
   - **Directory (tenant) ID**: Copy this value

### Step 2: Create Demo Configuration

Create `demo.env` in the project root:

```powershell
# Demo Tenant Configuration
$ClientId = "<demo-app-client-id>"
$TenantId = "<demo-tenant-id>"
```

**Note**: `demo.env` is gitignored and will not be committed.

### Step 3: Test Demo Connection

```powershell
# Load demo config
. .\demo.env

# Import modules
Import-Module Microsoft.Graph.Authentication

# Connect to Graph
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Scopes "Mail.ReadWrite","User.Read" -UseDeviceCode

# Verify connection
Get-MgContext

# Test mailbox access
Get-MgUserMailFolder -UserId me -MailFolderId Inbox

# Connect to Exchange Online
Connect-ExchangeOnline -Device

# Verify Exchange access
Get-InboxRule
```

### Step 4: Deploy to Demo Environment

```powershell
# Switch to demo config
Copy-Item demo.env .env

# Create demo rules config (or use example)
Copy-Item examples/rules-config.example.json rules-config.json

# Connect
.\Connect-OutlookRulesApp.ps1

# Deploy
.\Manage-OutlookRules.ps1 -Operation Deploy

# Verify
.\Manage-OutlookRules.ps1 -Operation List
.\Manage-OutlookRules.ps1 -Operation Folders
```

### Step 5: Clean Up Demo Environment

```powershell
# Remove all rules from demo
.\Manage-OutlookRules.ps1 -Operation DeleteAll

# Restore production config (if applicable)
Copy-Item .env.backup .env
```

---

## Development Environment

### Local Development Setup

1. **Clone Repository**

   ```powershell
   git clone https://github.com/fgarofalo56/outlook-rules-manager.git
   cd outlook-rules-manager
   ```

2. **Install Dependencies**

   ```powershell
   .\Install-Prerequisites.ps1
   ```

3. **Create Development Config**

   ```powershell
   # Create from example
   Copy-Item examples/.env.example .env
   Copy-Item examples/rules-config.example.json rules-config.json

   # Edit .env with your dev tenant credentials
   notepad .env
   ```

4. **Install Pre-Commit Hooks** (Optional but recommended)

   ```bash
   pip install pre-commit
   pre-commit install
   ```

### Development Workflow

```powershell
# 1. Make code changes

# 2. Run security check
.\scripts\Check-BeforeCommit.ps1

# 3. Test changes
.\Connect-OutlookRulesApp.ps1
.\Manage-OutlookRules.ps1 -Operation <your-operation>

# 4. Verify no sensitive data
git diff

# 5. Commit changes
git add .
git commit -m "Your commit message"
```

### Debugging Tips

#### Enable Verbose Output

```powershell
$VerbosePreference = "Continue"
.\Connect-OutlookRulesApp.ps1
```

#### Check Module Versions

```powershell
Get-Module Microsoft.Graph* -ListAvailable | Select-Object Name, Version
Get-Module ExchangeOnlineManagement -ListAvailable | Select-Object Name, Version
```

#### Test Individual Operations

```powershell
# Test Graph connection
$inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox
$inbox | Format-List *

# Test Exchange connection
Get-InboxRule | Format-Table Name, Priority, Enabled

# Test folder creation
Get-MgUserMailFolderChildFolder -UserId me -MailFolderId $inbox.Id
```

---

## Testing Procedures

### Unit Tests

#### Test 1: Configuration Loading

```powershell
# Verify .env loads correctly
. .\.env
Write-Host "ClientId: $ClientId"
Write-Host "TenantId: $TenantId"

# Verify rules-config.json is valid JSON
$config = Get-Content rules-config.json | ConvertFrom-Json
$config | Format-List
```

#### Test 2: Connection Verification

```powershell
# Test Graph connection
$graphCtx = Get-MgContext
if ($graphCtx) {
    Write-Host "Graph OK: $($graphCtx.Account)" -ForegroundColor Green
} else {
    Write-Host "Graph FAILED" -ForegroundColor Red
}

# Test Exchange connection
$exoConn = Get-ConnectionInformation | Where-Object { $_.Name -like "*ExchangeOnline*" }
if ($exoConn) {
    Write-Host "Exchange OK" -ForegroundColor Green
} else {
    Write-Host "Exchange FAILED" -ForegroundColor Red
}
```

#### Test 3: Mailbox Access

```powershell
# Test folder read
try {
    $inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox -ErrorAction Stop
    Write-Host "Folder Read OK: $($inbox.DisplayName)" -ForegroundColor Green
} catch {
    Write-Host "Folder Read FAILED: $($_.Exception.Message)" -ForegroundColor Red
}

# Test rule read
try {
    $rules = Get-InboxRule -ErrorAction Stop
    Write-Host "Rule Read OK: $($rules.Count) rules" -ForegroundColor Green
} catch {
    Write-Host "Rule Read FAILED: $($_.Exception.Message)" -ForegroundColor Red
}
```

### Integration Tests

#### Test 1: Full Deployment Cycle

```powershell
# Step 1: Backup existing rules
.\Manage-OutlookRules.ps1 -Operation Backup

# Step 2: Delete all rules
.\Manage-OutlookRules.ps1 -Operation DeleteAll -Force

# Step 3: Verify deletion
$rules = Get-InboxRule
if ($rules.Count -eq 0) {
    Write-Host "Deletion OK" -ForegroundColor Green
}

# Step 4: Deploy from config
.\Manage-OutlookRules.ps1 -Operation Deploy

# Step 5: Verify deployment
$rules = Get-InboxRule
Write-Host "Deployed $($rules.Count) rules"

# Step 6: Compare config vs deployed
.\Manage-OutlookRules.ps1 -Operation Compare
```

#### Test 2: Folder Creation

```powershell
# Get inbox ID
$inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox

# List child folders before
$beforeFolders = Get-MgUserMailFolderChildFolder -UserId me -MailFolderId $inbox.Id
Write-Host "Folders before: $($beforeFolders.Count)"

# Deploy (should create folders)
.\Manage-OutlookRules.ps1 -Operation Deploy

# List child folders after
$afterFolders = Get-MgUserMailFolderChildFolder -UserId me -MailFolderId $inbox.Id
Write-Host "Folders after: $($afterFolders.Count)"

# Verify specific folders
$expectedFolders = @("Priority", "Action Required", "Metrics", "Leadership", "Alerts", "Low Priority")
foreach ($name in $expectedFolders) {
    $folder = $afterFolders | Where-Object { $_.DisplayName -eq $name }
    if ($folder) {
        Write-Host "Folder OK: $name" -ForegroundColor Green
    } else {
        Write-Host "Folder MISSING: $name" -ForegroundColor Red
    }
}
```

#### Test 3: Rule Functionality

```powershell
# Get rule details
$rules = Get-InboxRule | Sort-Object Priority

foreach ($rule in $rules) {
    Write-Host "`n=== $($rule.Name) ===" -ForegroundColor Cyan
    Write-Host "Priority: $($rule.Priority)"
    Write-Host "Enabled: $($rule.Enabled)"
    Write-Host "Move To: $($rule.MoveToFolder)"
    Write-Host "Stop Processing: $($rule.StopProcessingRules)"

    # Check conditions
    if ($rule.From) { Write-Host "From: $($rule.From -join ', ')" }
    if ($rule.SubjectContainsWords) { Write-Host "Subject Contains: $($rule.SubjectContainsWords -join ', ')" }
    if ($rule.SenderDomainIs) { Write-Host "Domain Is: $($rule.SenderDomainIs -join ', ')" }
}
```

### End-to-End Test

#### Test with Real Email

1. **Send Test Emails**

   Send emails to your test mailbox that match each rule:

   | Test Case | Email Setup |
   |-----------|-------------|
   | Priority Sender | Send from a VIP address in your config |
   | Action Required | Subject contains "Action" or "Approval" |
   | Metrics | Subject contains "KPI" or "QBR" |
   | Leadership | Subject contains "Leadership" or "Executive" |
   | Alerts | Subject contains "Alert" or "Notification" |
   | Noise | Send from a newsletter domain |

2. **Verify Routing**

   ```powershell
   # Check folder contents
   .\Manage-OutlookRules.ps1 -Operation Folders
   ```

3. **Verify Actions**

   Check in Outlook:
   - Email landed in correct folder
   - Importance set correctly
   - Categories applied
   - Flags set
   - Mark as read (for alerts)

---

## Validation Checklist

### Pre-Deployment Checklist

- [ ] App registered in target tenant
- [ ] Authentication configured (public client flow enabled)
- [ ] API permissions added (Mail.ReadWrite, User.Read)
- [ ] `.env` file created with correct IDs
- [ ] `rules-config.json` customized with your senders/keywords
- [ ] Security check passes: `.\scripts\Check-BeforeCommit.ps1`

### Connection Checklist

- [ ] Device code authentication works
- [ ] Graph connection successful
- [ ] Exchange Online connection successful
- [ ] Can read inbox folders
- [ ] Can read inbox rules

### Deployment Checklist

- [ ] Compare shows expected changes
- [ ] Deploy completes without errors
- [ ] All expected folders created
- [ ] All expected rules created
- [ ] Rules enabled and in correct priority order

### Functional Checklist

- [ ] Priority sender email routes to Priority folder
- [ ] Action keywords trigger Action Required folder
- [ ] Metrics keywords trigger Metrics folder
- [ ] Leadership keywords trigger Leadership folder
- [ ] Alert keywords trigger Alerts folder + mark as read
- [ ] Noise domains route to Low Priority / Archive / Delete
- [ ] Priority sender rule stops processing (VIP mail only hits one rule)

### Security Checklist

- [ ] No real email addresses in tracked files
- [ ] No real GUIDs in tracked files (except Microsoft API constants)
- [ ] `.env` is gitignored
- [ ] `rules-config.json` is gitignored
- [ ] Pre-commit check passes

---

## Troubleshooting Tests

### Authentication Issues

```powershell
# Test 1: Verify app configuration
Write-Host "ClientId: $ClientId"
Write-Host "TenantId: $TenantId"

# Test 2: Test Graph auth directly
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Scopes "User.Read" -UseDeviceCode

# Test 3: Check scopes granted
(Get-MgContext).Scopes
```

### Permission Issues

```powershell
# Test: Verify Mail.ReadWrite scope
$scopes = (Get-MgContext).Scopes
if ($scopes -contains "Mail.ReadWrite") {
    Write-Host "Mail.ReadWrite granted" -ForegroundColor Green
} else {
    Write-Host "Mail.ReadWrite MISSING - reconnect with correct scopes" -ForegroundColor Red
}
```

### Rule Creation Issues

```powershell
# Test: Create a simple test rule
try {
    New-InboxRule -Name "ZZ_Test_Rule" -SubjectContainsWords "TEST12345" -MarkAsRead $true
    Write-Host "Rule creation OK" -ForegroundColor Green

    # Clean up
    Remove-InboxRule -Identity "ZZ_Test_Rule" -Confirm:$false
    Write-Host "Rule deletion OK" -ForegroundColor Green
} catch {
    Write-Host "Rule creation FAILED: $($_.Exception.Message)" -ForegroundColor Red
}
```

### Folder Creation Issues

```powershell
# Test: Create a test folder
$inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox

try {
    $testFolder = New-MgUserMailFolderChildFolder -UserId me -MailFolderId $inbox.Id -BodyParameter @{
        DisplayName = "ZZ_Test_Folder"
    }
    Write-Host "Folder creation OK: $($testFolder.Id)" -ForegroundColor Green

    # Clean up
    Remove-MgUserMailFolder -UserId me -MailFolderId $testFolder.Id
    Write-Host "Folder deletion OK" -ForegroundColor Green
} catch {
    Write-Host "Folder creation FAILED: $($_.Exception.Message)" -ForegroundColor Red
}
```

---

## CI/CD Integration

### GitHub Actions Security Scan

The repository includes automated security scanning on every push:

| Scan | Purpose |
|------|---------|
| Gitleaks | Secret detection |
| PII Detection | Email address scanning |
| PSScriptAnalyzer | PowerShell code quality |
| Blocked Files | Prevent sensitive file commits |

### Running Security Scan Locally

```powershell
# Run pre-commit check
.\scripts\Check-BeforeCommit.ps1

# Expected output:
# [PASS] No blocked files detected
# [PASS] No blocked email domains found
# [PASS] No exposed Azure GUIDs found
# [PASS] No secret patterns found
# [PASS] ALL CHECKS PASSED - Safe to commit!
```

### Pre-Commit Hooks

Install pre-commit hooks for automatic checking:

```bash
# Install pre-commit
pip install pre-commit

# Install hooks
pre-commit install

# Run manually
pre-commit run --all-files
```

### Branch Protection

The `main` branch is protected with:

- Required pull request reviews
- Required status checks (security scan)
- No direct pushes
- No force pushes

---

## Quick Reference

### Test Commands Cheatsheet

```powershell
# Connection tests
Get-MgContext                              # Check Graph connection
Get-ConnectionInformation                   # Check Exchange connection

# Mailbox tests
Get-MgUserMailFolder -UserId me -MailFolderId Inbox    # Test folder access
Get-InboxRule                              # Test rule access

# Deployment tests
.\Manage-OutlookRules.ps1 -Operation Compare    # Preview changes
.\Manage-OutlookRules.ps1 -Operation Validate   # Check for issues
.\Manage-OutlookRules.ps1 -Operation List       # View deployed rules
.\Manage-OutlookRules.ps1 -Operation Folders    # View folders

# Security tests
.\scripts\Check-BeforeCommit.ps1           # Security scan
git status                                 # Check for uncommitted sensitive files
```
