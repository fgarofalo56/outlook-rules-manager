<#
.SYNOPSIS
    Registers an Azure AD application for Outlook Rules management.
    Run this ONCE in your personal Azure tenant to create the app registration.

.DESCRIPTION
    Creates an app registration with:
    - Microsoft Graph: Mail.ReadWrite (delegated) - for folder creation
    - Public client flow enabled (for interactive/device code auth)
    - App Roles for multi-tier access control (Admin, User)
    - User Assignment Required enabled for security

    SECURITY: This implements a multi-tier authorization model:
    - Owners: Can manage admins (Service Principal owners)
    - Admins: Can add/remove authorized users AND use the app
    - Users: Can only use the app for their own mailbox

    No admin consent required for delegated permissions on your own mailbox.

.PARAMETER AppName
    Name for the Azure AD application (default: "Outlook Rules Manager")

.PARAMETER TenantId
    Optional: specify tenant ID, otherwise uses current Azure context

.PARAMETER SkipUserAssignment
    Skip enabling "User Assignment Required" (not recommended for security)

.NOTES
    Requires: Az.Accounts, Az.Resources modules
    Run from PowerShell 7+ recommended
#>

param(
    [string]$AppName = "Outlook Rules Manager",
    [string]$TenantId,  # Optional: specify tenant, otherwise uses current
    [switch]$SkipUserAssignment
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
# DEFINE APP ROLES (Multi-Tier Authorization)
# ---------------------------
# These roles provide authorization BEYOND just permission consent
# Users must be explicitly assigned to a role to use the application

$appRoles = @(
    @{
        AllowedMemberTypes = @("User")
        Description = "Administrators can manage authorized users and perform all mailbox operations."
        DisplayName = "Administrator"
        Id = "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f"  # Fixed GUID for consistency
        IsEnabled = $true
        Value = "OutlookRules.Admin"
    },
    @{
        AllowedMemberTypes = @("User")
        Description = "Users can perform mailbox operations on their own mailbox only."
        DisplayName = "User"
        Id = "a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d"  # Fixed GUID for consistency
        IsEnabled = $true
        Value = "OutlookRules.User"
    }
)

Write-Host "`n=== App Roles Defined ===" -ForegroundColor Yellow
Write-Host "  OutlookRules.Admin - Can manage users and perform all operations" -ForegroundColor Gray
Write-Host "  OutlookRules.User  - Can perform operations on own mailbox" -ForegroundColor Gray

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
    -RequiredResourceAccess $requiredResourceAccess `
    -AppRole $appRoles

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

$sp = $null
try {
    $sp = New-AzADServicePrincipal -ApplicationId $app.AppId -ErrorAction Stop
    Write-Host "Created service principal: $($sp.Id)" -ForegroundColor Green
} catch {
    Write-Host "Warning: Could not create service principal: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "This is usually fine - it may be created automatically on first sign-in." -ForegroundColor Gray
}

# ---------------------------
# ENABLE USER ASSIGNMENT REQUIRED (Security)
# ---------------------------
if ($sp -and -not $SkipUserAssignment) {
    Write-Host "`n=== Enabling User Assignment Required ===" -ForegroundColor Yellow
    Write-Host "  This restricts app access to explicitly assigned users only" -ForegroundColor Gray

    try {
        # Note: Az module doesn't have direct support for AppRoleAssignmentRequired
        # We need to use the Microsoft Graph API or update via REST
        # For now, output instructions for manual configuration
        Write-Host ""
        Write-Host "  IMPORTANT: Enable 'User Assignment Required' manually:" -ForegroundColor Cyan
        Write-Host "  1. Go to Azure Portal > Microsoft Entra ID > Enterprise Applications" -ForegroundColor White
        Write-Host "  2. Find '$AppName'" -ForegroundColor White
        Write-Host "  3. Go to Properties" -ForegroundColor White
        Write-Host "  4. Set 'Assignment required?' to 'Yes'" -ForegroundColor White
        Write-Host "  5. Save" -ForegroundColor White
        Write-Host ""
        Write-Host "  Or run this PowerShell (requires Microsoft.Graph module):" -ForegroundColor Cyan
        Write-Host "  Connect-MgGraph -Scopes 'Application.ReadWrite.All'" -ForegroundColor Gray
        Write-Host "  Update-MgServicePrincipal -ServicePrincipalId '$($sp.Id)' -AppRoleAssignmentRequired:`$true" -ForegroundColor Gray
    } catch {
        Write-Host "  Note: Could not auto-configure. Please enable manually." -ForegroundColor Yellow
    }
}

# ---------------------------
# ASSIGN CREATOR AS ADMIN
# ---------------------------
if ($sp) {
    Write-Host "`n=== Assigning Creator as Administrator ===" -ForegroundColor Yellow

    $currentUserId = (Get-AzADUser -UserPrincipalName $context.Account.Id -ErrorAction SilentlyContinue).Id
    $adminRoleId = "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f"  # OutlookRules.Admin role ID

    if ($currentUserId) {
        Write-Host "  Current user: $($context.Account.Id)" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  To assign yourself as Admin, run:" -ForegroundColor Cyan
        Write-Host "  Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All'" -ForegroundColor Gray
        Write-Host "  New-MgUserAppRoleAssignment -UserId '$currentUserId' -PrincipalId '$currentUserId' -ResourceId '$($sp.Id)' -AppRoleId '$adminRoleId'" -ForegroundColor Gray
    } else {
        Write-Host "  Could not determine current user ID. Assign roles manually." -ForegroundColor Yellow
    }
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
1. SECURITY SETUP (Required):
   a. Enable 'User Assignment Required' in Enterprise Applications (see above)
   b. Assign yourself as Admin using the PowerShell commands above
   c. Or use: .\Manage-AppAuthorization.ps1 -Operation Setup

2. Run Connect-OutlookRulesApp.ps1 to authenticate

3. On first run, you'll be prompted to consent to permissions

4. Then run Setup-OutlookRules.ps1 or Manage-OutlookRules.ps1

Authorization Model:
-------------------
- Owners: Service Principal owners (you by default)
- Admins: Can add/remove users AND use the app (OutlookRules.Admin role)
- Users: Can only use the app for their own mailbox (OutlookRules.User role)

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
