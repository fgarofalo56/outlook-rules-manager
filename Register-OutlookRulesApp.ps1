<#
.SYNOPSIS
    Registers an Azure AD application for Outlook Rules management.
    Run this ONCE in your personal Azure tenant to create the app registration.

.DESCRIPTION
    Creates an app registration with:
    - Microsoft Graph: Mail.ReadWrite (delegated) - for folder creation
    - Public client flow enabled (for interactive/device code auth)

    No admin consent required for delegated permissions on your own mailbox.

.NOTES
    Requires: Az.Accounts, Az.Resources modules
    Run from PowerShell 7+ recommended
#>

param(
    [string]$AppName = "Outlook Rules Manager",
    [string]$TenantId  # Optional: specify tenant, otherwise uses current
)

# ---------------------------
# INSTALL REQUIRED MODULES
# ---------------------------
$requiredModules = @("Az.Accounts", "Az.Resources")
foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing $mod..." -ForegroundColor Cyan
        Install-Module $mod -Scope CurrentUser -Repository PSGallery -Force
    }
}

Import-Module Az.Accounts
Import-Module Az.Resources

# ---------------------------
# AUTHENTICATE TO AZURE
# ---------------------------
Write-Host "`n=== Azure Authentication ===" -ForegroundColor Yellow

$context = Get-AzContext
if (-not $context) {
    if ($TenantId) {
        Connect-AzAccount -TenantId $TenantId
    } else {
        Connect-AzAccount
    }
    $context = Get-AzContext
}

$tenantIdResolved = $context.Tenant.Id
Write-Host "Using Tenant: $tenantIdResolved" -ForegroundColor Green
Write-Host "Account: $($context.Account.Id)" -ForegroundColor Green

# ---------------------------
# CHECK FOR EXISTING APP
# ---------------------------
Write-Host "`n=== Checking for existing app ===" -ForegroundColor Yellow

$existingApp = Get-AzADApplication -DisplayName $AppName -ErrorAction SilentlyContinue
if ($existingApp) {
    Write-Host "App '$AppName' already exists!" -ForegroundColor Yellow
    Write-Host "  Application (client) ID: $($existingApp.AppId)" -ForegroundColor Cyan
    Write-Host "  Object ID: $($existingApp.Id)" -ForegroundColor Cyan

    $response = Read-Host "Do you want to delete and recreate it? (y/N)"
    if ($response -eq 'y') {
        Remove-AzADApplication -ObjectId $existingApp.Id -Force
        Write-Host "Deleted existing app." -ForegroundColor Yellow
    } else {
        Write-Host "`nUsing existing app. Update your config with:" -ForegroundColor Green
        Write-Host "  `$ClientId = '$($existingApp.AppId)'" -ForegroundColor White
        Write-Host "  `$TenantId = '$tenantIdResolved'" -ForegroundColor White
        exit 0
    }
}

# ---------------------------
# DEFINE API PERMISSIONS
# ---------------------------
# Microsoft Graph API ID
$graphApiId = "00000003-0000-0000-c000-000000000000"

# Permission IDs (delegated/scope permissions)
# Mail.ReadWrite: e383f46e-2787-4529-855e-0e479a3ffac0
# Mail.Read: 570282fd-fa5c-430d-a7fd-fc8dc98a9dca
# User.Read: e1fe6dd8-ba31-4d61-89e7-88639da4683d

$requiredResourceAccess = @(
    @{
        ResourceAppId = $graphApiId
        ResourceAccess = @(
            @{
                Id = "e383f46e-2787-4529-855e-0e479a3ffac0"  # Mail.ReadWrite
                Type = "Scope"  # Delegated
            },
            @{
                Id = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"  # User.Read
                Type = "Scope"  # Delegated
            }
        )
    }
)

# ---------------------------
# CREATE APP REGISTRATION
# ---------------------------
Write-Host "`n=== Creating App Registration ===" -ForegroundColor Yellow

# Create the application with basic settings first
# Using AzureADMultipleOrgs to allow cross-tenant auth (app in personal tenant, mailbox in work tenant)
$app = New-AzADApplication `
    -DisplayName $AppName `
    -SignInAudience "AzureADMultipleOrgs" `
    -IsFallbackPublicClient:$true `
    -RequiredResourceAccess $requiredResourceAccess

if (-not $app -or -not $app.AppId) {
    Write-Host "ERROR: Failed to create application" -ForegroundColor Red
    exit 1
}

Write-Host "Created app registration: $($app.DisplayName)" -ForegroundColor Green
Write-Host "  Application (client) ID: $($app.AppId)" -ForegroundColor Cyan

# Add public client redirect URIs (SPACE-compliant - no localhost)
Write-Host "  Adding redirect URIs..." -ForegroundColor Gray
$redirectUris = @(
    "https://login.microsoftonline.com/common/oauth2/nativeclient"
)

try {
    Update-AzADApplication -ObjectId $app.Id -PublicClientRedirectUri $redirectUris -ErrorAction Stop
    Write-Host "  Redirect URIs configured" -ForegroundColor Green
} catch {
    Write-Host "  Warning: Could not set redirect URIs (may not be required): $($_.Exception.Message)" -ForegroundColor Yellow
}

# ---------------------------
# CREATE SERVICE PRINCIPAL
# ---------------------------
Write-Host "`n=== Creating Service Principal ===" -ForegroundColor Yellow

try {
    $sp = New-AzADServicePrincipal -ApplicationId $app.AppId -ErrorAction Stop
    Write-Host "Created service principal: $($sp.Id)" -ForegroundColor Green
} catch {
    Write-Host "Warning: Could not create service principal: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "This is usually fine - it may be created automatically on first sign-in." -ForegroundColor Gray
}

# ---------------------------
# OUTPUT CONFIGURATION
# ---------------------------
Write-Host ""
Write-Host ("=" * 60) -ForegroundColor Green
Write-Host "  APP REGISTRATION COMPLETE" -ForegroundColor Green
Write-Host ("=" * 60) -ForegroundColor Green

$configOutput = @"

Configuration saved! Next steps:
--------------------------------
1. Run Connect-OutlookRulesApp.ps1 to authenticate
2. On first run, you'll be prompted to consent to permissions
3. Then run Setup-OutlookRules.ps1 or Manage-OutlookRules.ps1

Note: No admin consent required - you're consenting for your own mailbox only.

"@

Write-Host $configOutput -ForegroundColor White

# Save to .env file (primary config)
$envFile = Join-Path $PSScriptRoot ".env"
$envContent = @"
`$ClientId = "$($app.AppId)"
`$TenantId = "$tenantIdResolved"
"@
Set-Content -Path $envFile -Value $envContent
Write-Host "Configuration saved to: $envFile" -ForegroundColor Cyan

# Also save to JSON for backwards compatibility
$configFile = Join-Path $PSScriptRoot "app-config.json"
@{
    AppName = $AppName
    ClientId = $app.AppId
    TenantId = $tenantIdResolved
    Created = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
} | ConvertTo-Json | Set-Content $configFile
Write-Host "Also saved to: $configFile (backup)" -ForegroundColor Gray
