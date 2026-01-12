<#
.SYNOPSIS
    Connects to Microsoft Graph and Exchange Online using your registered app.
    Run this BEFORE running Setup-OutlookRules.ps1

.DESCRIPTION
    Uses the app registration created by Register-OutlookRulesApp.ps1 to:
    - Connect to Microsoft Graph (for folder management)
    - Connect to Exchange Online (for inbox rules)

    Supports interactive auth and device code flow (for restricted environments).

.PARAMETER UseDeviceCode
    Use device code flow instead of interactive browser popup.
    Useful in environments where browser popups are blocked.

.EXAMPLE
    .\Connect-OutlookRulesApp.ps1
    # Interactive browser login

.EXAMPLE
    .\Connect-OutlookRulesApp.ps1 -UseDeviceCode
    # Device code flow (displays code to enter at microsoft.com/devicelogin)
#>

param(
    [switch]$UseDeviceCode
)

# ---------------------------
# CONFIGURATION
# Load from .env (primary) or app-config.json (fallback)
# ---------------------------
$envFile = Join-Path $PSScriptRoot ".env"
$configFile = Join-Path $PSScriptRoot "app-config.json"

if (Test-Path $envFile) {
    # Parse .env file (supports both formats: KEY=value and $KEY = "value")
    $envContent = Get-Content $envFile
    foreach ($line in $envContent) {
        $line = $line.Trim()
        if ($line -and -not $line.StartsWith("#")) {
            # Match: ClientId="value" or ClientId=value or $ClientId = "value"
            if ($line -match '^\$?(\w+)\s*=\s*"?([^"]+)"?$') {
                $key = $matches[1]
                $value = $matches[2]
                Set-Variable -Name $key -Value $value -Scope Script
            }
        }
    }
    Write-Host "Loaded config from .env" -ForegroundColor Green
} elseif (Test-Path $configFile) {
    # Fallback to JSON config
    $config = Get-Content $configFile | ConvertFrom-Json
    $ClientId = $config.ClientId
    $TenantId = $config.TenantId
    Write-Host "Loaded config from app-config.json" -ForegroundColor Green
} else {
    Write-Host "ERROR: No configuration found!" -ForegroundColor Red
    Write-Host "Either:" -ForegroundColor Yellow
    Write-Host "  1. Run Register-OutlookRulesApp.ps1 first to create .env" -ForegroundColor Yellow
    Write-Host "  2. Create .env with ClientId and TenantId values" -ForegroundColor Yellow
    exit 1
}

# Validate config was loaded
if (-not $ClientId -or -not $TenantId) {
    Write-Host "ERROR: ClientId or TenantId not found in config!" -ForegroundColor Red
    exit 1
}

Write-Host "`n=== Outlook Rules Connection ===" -ForegroundColor Cyan
Write-Host "Client ID: $ClientId" -ForegroundColor Gray
Write-Host "Tenant ID: $TenantId" -ForegroundColor Gray

# ---------------------------
# ENSURE MODULES INSTALLED
# ---------------------------
$modules = @("Microsoft.Graph.Authentication", "ExchangeOnlineManagement")
foreach ($mod in $modules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing $mod..." -ForegroundColor Yellow
        Install-Module $mod -Scope CurrentUser -Repository PSGallery -Force
    }
}

# ---------------------------
# CONNECT TO MICROSOFT GRAPH
# ---------------------------
Write-Host "`n[1/2] Connecting to Microsoft Graph..." -ForegroundColor Yellow

# Disconnect existing session if any
$existingContext = Get-MgContext -ErrorAction SilentlyContinue
if ($existingContext) {
    Write-Host "  Disconnecting existing Graph session..." -ForegroundColor Gray
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

# Use specific TenantId for single-tenant apps, or 'common' for multi-tenant
$targetTenant = $TenantId  # Use the tenant where your app is registered

# Option 1: Use custom app (set to $false)
# Option 2: Use Microsoft Graph PowerShell SDK's built-in app (set to $true)
$useBuiltInApp = $false  # Using custom work tenant app

if ($useBuiltInApp) {
    # Don't specify ClientId - let Connect-MgGraph use the Microsoft Graph PowerShell SDK's
    # built-in app registration which has proper redirect URIs configured
    Write-Host "  Using Microsoft Graph PowerShell SDK (built-in)" -ForegroundColor Cyan
    $graphParams = @{
        TenantId  = $targetTenant
        Scopes    = @("Mail.ReadWrite", "User.Read")
        NoWelcome = $true
    }
} else {
    Write-Host "  Using custom app: $ClientId" -ForegroundColor Cyan
    $graphParams = @{
        ClientId  = $ClientId
        TenantId  = $targetTenant
        Scopes    = @("Mail.ReadWrite", "User.Read")
        NoWelcome = $true
    }
}

if ($UseDeviceCode) {
    $graphParams["UseDeviceCode"] = $true
    Write-Host "  Using device code flow - check console for instructions" -ForegroundColor Cyan
}

try {
    Connect-MgGraph @graphParams
    $ctx = Get-MgContext
    Write-Host "  Connected to Graph as: $($ctx.Account)" -ForegroundColor Green
    Write-Host "  Scopes: $($ctx.Scopes -join ', ')" -ForegroundColor Gray
} catch {
    Write-Host "  ERROR connecting to Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ---------------------------
# CONNECT TO EXCHANGE ONLINE
# ---------------------------
Write-Host "`n[2/2] Connecting to Exchange Online..." -ForegroundColor Yellow

# Check for existing connection
$exoConn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*ExchangeOnline*" }
if ($exoConn) {
    Write-Host "  Already connected to Exchange Online" -ForegroundColor Green
} else {
    $exoParams = @{}

    if ($UseDeviceCode) {
        $exoParams["Device"] = $true
        Write-Host "  Using device code flow for Exchange Online" -ForegroundColor Cyan
    }

    try {
        Connect-ExchangeOnline @exoParams -ShowBanner:$false
        Write-Host "  Connected to Exchange Online" -ForegroundColor Green
    } catch {
        Write-Host "  ERROR connecting to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Note: Exchange Online connection is separate from Graph" -ForegroundColor Yellow
        Write-Host "  You may need to authenticate again" -ForegroundColor Yellow
        exit 1
    }
}

# ---------------------------
# VERIFY CONNECTIONS
# ---------------------------
Write-Host "`n=== Connection Verification ===" -ForegroundColor Cyan

# Verify Graph - try to get inbox
try {
    $inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox -ErrorAction Stop
    Write-Host "Graph OK - Inbox ID: $($inbox.Id.Substring(0,20))..." -ForegroundColor Green
} catch {
    Write-Host "Graph FAILED - Cannot access mailbox: $($_.Exception.Message)" -ForegroundColor Red
}

# Verify Exchange - try to list rules
try {
    $ruleCount = (Get-InboxRule -ErrorAction Stop | Measure-Object).Count
    Write-Host "Exchange OK - Found $ruleCount existing inbox rules" -ForegroundColor Green
} catch {
    Write-Host "Exchange FAILED - Cannot access rules: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== Ready ===" -ForegroundColor Green
Write-Host "You can now run: .\Setup-OutlookRules.ps1" -ForegroundColor White
