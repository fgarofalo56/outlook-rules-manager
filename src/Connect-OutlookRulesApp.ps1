<#
.SYNOPSIS
    Connects to Microsoft Graph and Exchange Online using your registered app.
    Run this BEFORE running Setup-OutlookRules.ps1

.DESCRIPTION
    Uses the app registration created by Register-OutlookRulesApp.ps1 to:
    - Connect to Microsoft Graph (for folder management)
    - Connect to Exchange Online (for inbox rules)

    DEFAULTS TO DEVICE CODE FLOW for SPACE/ACE compliance (localhost redirect URIs not allowed).

.PARAMETER ConfigProfile
    Profile name for multi-account support. Loads .env.{profile} instead of .env.
    Example: -ConfigProfile personal loads .env.personal
    Example: -ConfigProfile work loads .env.work

.PARAMETER UseDeviceCode
    Use device code flow (DEFAULT). Displays a code to enter at microsoft.com/devicelogin.
    This is the default because localhost redirect URIs are not allowed by SPACE compliance.

.PARAMETER Interactive
    Use interactive browser popup instead of device code.
    Only use if you've added this redirect URI to your app registration:
    https://login.microsoftonline.com/common/oauth2/nativeclient

.EXAMPLE
    .\Connect-OutlookRulesApp.ps1
    # Device code flow (default) - displays code to enter at microsoft.com/devicelogin

.EXAMPLE
    .\Connect-OutlookRulesApp.ps1 -ConfigProfile personal
    # Connect using .env.personal configuration for personal email management

.EXAMPLE
    .\Connect-OutlookRulesApp.ps1 -ConfigProfile work
    # Connect using .env.work configuration for work email management

.EXAMPLE
    .\Connect-OutlookRulesApp.ps1 -Interactive
    # Interactive browser login (requires native client redirect URI configured)
#>

param(
    [string]$ConfigProfile,
    [switch]$UseDeviceCode,
    [switch]$Interactive  # Use interactive browser flow (requires native client redirect URI)
)

# Default to Device Code flow since localhost redirect URIs are not allowed by SPACE compliance
# Use -Interactive flag only if you've configured https://login.microsoftonline.com/common/oauth2/nativeclient
if (-not $Interactive) {
    $UseDeviceCode = $true
}

# ---------------------------
# PATH RESOLUTION
# ---------------------------

# Project root is parent of src/ directory where this script lives
$ProjectRoot = Split-Path $PSScriptRoot -Parent

# ---------------------------
# CONFIGURATION
# Load from .env.{profile} (if profile specified), .env (primary), or app-config.json (fallback)
# Config files are stored at project root for easy access
# ---------------------------

# Determine which env file to load based on profile
if ($ConfigProfile) {
    $envFile = Join-Path $ProjectRoot ".env.$ConfigProfile"
    $configFile = Join-Path $ProjectRoot "app-config.$ConfigProfile.json"
    $profileDisplay = "[$ConfigProfile]"
} else {
    $envFile = Join-Path $ProjectRoot ".env"
    $configFile = Join-Path $ProjectRoot "app-config.json"
    $profileDisplay = "[default]"
}

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
    Write-Host "Loaded config from $(Split-Path $envFile -Leaf) $profileDisplay" -ForegroundColor Green
} elseif (Test-Path $configFile) {
    # Fallback to JSON config
    $config = Get-Content $configFile | ConvertFrom-Json
    $ClientId = $config.ClientId
    $TenantId = $config.TenantId
    Write-Host "Loaded config from $(Split-Path $configFile -Leaf) $profileDisplay" -ForegroundColor Green
} else {
    Write-Host "ERROR: No configuration found!" -ForegroundColor Red
    if ($ConfigProfile) {
        Write-Host "Profile '$ConfigProfile' specified but no config file found." -ForegroundColor Red
        Write-Host "Expected: .env.$ConfigProfile or app-config.$ConfigProfile.json" -ForegroundColor Yellow
    }
    Write-Host "Either:" -ForegroundColor Yellow
    Write-Host "  1. Run Register-OutlookRulesApp.ps1 first to create .env" -ForegroundColor Yellow
    Write-Host "  2. Create .env with ClientId and TenantId values" -ForegroundColor Yellow
    Write-Host "  3. For multi-account: Create .env.<profile> (e.g., .env.personal, .env.work)" -ForegroundColor Yellow
    exit 1
}

# Validate config was loaded
if (-not $ClientId -or -not $TenantId) {
    Write-Host "ERROR: ClientId or TenantId not found in config!" -ForegroundColor Red
    exit 1
}

Write-Host "`n=== Outlook Rules Connection $profileDisplay ===" -ForegroundColor Cyan
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

# ---------------------------
# VERIFY AUTHORIZATION (App Role Assignment)
# ---------------------------
Write-Host "`n=== Authorization Check ===" -ForegroundColor Cyan

try {
    # Get the service principal for this app
    $sp = Get-MgServicePrincipal -Filter "AppId eq '$ClientId'" -ErrorAction SilentlyContinue

    if ($sp) {
        $currentUserUpn = (Get-MgContext).Account
        $currentUser = Get-MgUser -Filter "userPrincipalName eq '$currentUserUpn'" -ErrorAction SilentlyContinue
        if (-not $currentUser) {
            $currentUser = Get-MgUser -Filter "mail eq '$currentUserUpn'" -ErrorAction SilentlyContinue
        }

        if ($currentUser) {
            $assignment = Get-MgUserAppRoleAssignment -UserId $currentUser.Id -ErrorAction SilentlyContinue |
                Where-Object { $_.ResourceId -eq $sp.Id }

            if ($assignment) {
                $roleName = switch ($assignment.AppRoleId) {
                    "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f" { "Administrator" }
                    "a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d" { "User" }
                    default { "Unknown" }
                }
                Write-Host "Authorization OK - Role: $roleName" -ForegroundColor Green
            } else {
                if ($sp.AppRoleAssignmentRequired) {
                    Write-Host "Authorization FAILED - User not assigned to application" -ForegroundColor Red
                    Write-Host "Contact an administrator or run: .\Manage-AppAuthorization.ps1 -Operation Setup" -ForegroundColor Yellow
                } else {
                    Write-Host "Authorization: Not configured (User Assignment not required)" -ForegroundColor Yellow
                    Write-Host "For better security, run: .\Manage-AppAuthorization.ps1 -Operation Setup" -ForegroundColor Gray
                }
            }
        } else {
            Write-Host "Authorization: Could not verify (user lookup failed)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Authorization: Not configured (service principal not found)" -ForegroundColor Yellow
        Write-Host "This may be normal for first-time setup" -ForegroundColor Gray
    }
} catch {
    Write-Host "Authorization: Could not verify - $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`n=== Ready ===" -ForegroundColor Green
Write-Host "You can now run: .\Manage-OutlookRules.ps1 -Operation List" -ForegroundColor White
