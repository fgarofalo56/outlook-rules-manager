<#
.SYNOPSIS
    Validates an existing Azure AD app registration for use with Outlook Rules Manager.
    Use this if you cannot create a new app registration and need to use an existing one.

.DESCRIPTION
    Checks that an existing app registration has all required settings:
    - Public client flow enabled
    - Required API permissions (Mail.ReadWrite, User.Read)
    - Proper redirect URIs (optional but recommended)
    - App roles (optional - for multi-user authorization)

    Provides detailed feedback on what's configured correctly and what needs fixing.

.PARAMETER ClientId
    The Application (client) ID of the existing app registration to validate.

.PARAMETER TenantId
    The tenant ID where the app registration exists.

.PARAMETER AutoFix
    Attempt to automatically fix issues (requires appropriate permissions).

.PARAMETER SaveConfig
    Save the ClientId and TenantId to .env after successful validation.

.PARAMETER ConfigProfile
    Profile name for multi-account support. Saves to .env.{profile} instead of .env.
    Example: -ConfigProfile personal saves to .env.personal

.EXAMPLE
    .\Test-ExistingAppRegistration.ps1 -ClientId "your-client-id" -TenantId "your-tenant-id"
    # Validate an existing app registration

.EXAMPLE
    .\Test-ExistingAppRegistration.ps1 -ClientId "abc123" -TenantId "xyz789" -SaveConfig
    # Validate and save configuration if successful

.EXAMPLE
    .\Test-ExistingAppRegistration.ps1 -ClientId "abc123" -TenantId "xyz789" -AutoFix
    # Validate and attempt to fix any issues automatically

.NOTES
    Requires: Az.Accounts, Az.Resources modules
    Run from PowerShell 7+ recommended
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [switch]$AutoFix,
    [switch]$SaveConfig,
    [string]$ConfigProfile
)

# ---------------------------
# CONSTANTS
# ---------------------------
$GRAPH_API_ID = "00000003-0000-0000-c000-000000000000"
$MAIL_READWRITE_ID = "e383f46e-2787-4529-855e-0e479a3ffac0"
$USER_READ_ID = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"
$NATIVE_REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient"

$ADMIN_ROLE_ID = "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f"
$USER_ROLE_ID = "a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d"

# Track validation results
$validationResults = @{
    Passed   = @()
    Warnings = @()
    Failed   = @()
    Fixed    = @()
}

function Add-Result {
    param([string]$Type, [string]$Message, [string]$Details = "")
    $result = @{ Message = $Message; Details = $Details }
    $validationResults[$Type] += $result
}

# ---------------------------
# INSTALL/IMPORT REQUIRED MODULES
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
if (-not $context -or $context.Tenant.Id -ne $TenantId) {
    Write-Host "Connecting to Azure tenant: $TenantId" -ForegroundColor Cyan
    try {
        Connect-AzAccount -TenantId $TenantId -ErrorAction Stop | Out-Null
        $context = Get-AzContext
    }
    catch {
        Write-Host "ERROR: Failed to connect to Azure: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

Write-Host "Connected as: $($context.Account.Id)" -ForegroundColor Green
Write-Host "Tenant: $($context.Tenant.Id)" -ForegroundColor Gray

# ---------------------------
# FIND THE APP REGISTRATION
# ---------------------------
Write-Host "`n=== Finding App Registration ===" -ForegroundColor Yellow

$app = Get-AzADApplication -Filter "AppId eq '$ClientId'" -ErrorAction SilentlyContinue

if (-not $app) {
    Write-Host "ERROR: No app registration found with Client ID: $ClientId" -ForegroundColor Red
    Write-Host "`nPossible causes:" -ForegroundColor Yellow
    Write-Host "  1. The Client ID is incorrect" -ForegroundColor Gray
    Write-Host "  2. The app exists in a different tenant" -ForegroundColor Gray
    Write-Host "  3. You don't have permission to view the app registration" -ForegroundColor Gray
    Write-Host "`nTo find your app:" -ForegroundColor Yellow
    Write-Host "  1. Go to Azure Portal > Microsoft Entra ID > App registrations" -ForegroundColor Gray
    Write-Host "  2. Search for your app name" -ForegroundColor Gray
    Write-Host "  3. Copy the 'Application (client) ID'" -ForegroundColor Gray
    exit 1
}

Write-Host "Found app: $($app.DisplayName)" -ForegroundColor Green
Write-Host "  Object ID: $($app.Id)" -ForegroundColor Gray
Write-Host "  Client ID: $($app.AppId)" -ForegroundColor Gray

# ---------------------------
# VALIDATION CHECKS
# ---------------------------
Write-Host "`n=== Validating App Configuration ===" -ForegroundColor Yellow

# Check 1: Public Client Flow
Write-Host "`n[1/6] Checking Public Client Configuration..." -ForegroundColor Cyan

if ($app.IsFallbackPublicClient -eq $true) {
    Add-Result -Type "Passed" -Message "Public client flow is ENABLED" -Details "Required for device code authentication"
}
else {
    if ($AutoFix) {
        Write-Host "  Attempting to enable public client flow..." -ForegroundColor Yellow
        try {
            Update-AzADApplication -ObjectId $app.Id -IsFallbackPublicClient:$true -ErrorAction Stop
            Add-Result -Type "Fixed" -Message "Public client flow ENABLED" -Details "Was disabled, now fixed"
        }
        catch {
            Add-Result -Type "Failed" -Message "Public client flow is DISABLED" -Details "AutoFix failed: $($_.Exception.Message)"
        }
    }
    else {
        Add-Result -Type "Failed" -Message "Public client flow is DISABLED" -Details "Required for device code authentication. Enable in: App Registration > Authentication > Allow public client flows"
    }
}

# Check 2: Sign-In Audience
Write-Host "[2/6] Checking Sign-In Audience..." -ForegroundColor Cyan

$signInAudience = $app.SignInAudience
$audienceOk = $signInAudience -in @("AzureADMultipleOrgs", "AzureADandPersonalMicrosoftAccount", "PersonalMicrosoftAccount", "AzureADMyOrg")

if ($audienceOk) {
    Add-Result -Type "Passed" -Message "Sign-in audience: $signInAudience" -Details "Appropriate for the use case"
}
else {
    Add-Result -Type "Warning" -Message "Sign-in audience: $signInAudience" -Details "May limit which accounts can authenticate"
}

# Check 3: Required API Permissions
Write-Host "[3/6] Checking API Permissions..." -ForegroundColor Cyan

$graphPermissions = $app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq $GRAPH_API_ID }

$hasMailReadWrite = $false
$hasUserRead = $false

if ($graphPermissions) {
    foreach ($access in $graphPermissions.ResourceAccess) {
        if ($access.Id -eq $MAIL_READWRITE_ID -and $access.Type -eq "Scope") {
            $hasMailReadWrite = $true
        }
        if ($access.Id -eq $USER_READ_ID -and $access.Type -eq "Scope") {
            $hasUserRead = $true
        }
    }
}

if ($hasMailReadWrite) {
    Add-Result -Type "Passed" -Message "Mail.ReadWrite permission is configured" -Details "Delegated permission for mailbox access"
}
else {
    if ($AutoFix) {
        Write-Host "  Attempting to add Mail.ReadWrite permission..." -ForegroundColor Yellow
        try {
            # Build new resource access
            $newResourceAccess = @(
                @{
                    ResourceAppId  = $GRAPH_API_ID
                    ResourceAccess = @(
                        @{ Id = $MAIL_READWRITE_ID; Type = "Scope" },
                        @{ Id = $USER_READ_ID; Type = "Scope" }
                    )
                }
            )
            Update-AzADApplication -ObjectId $app.Id -RequiredResourceAccess $newResourceAccess -ErrorAction Stop
            Add-Result -Type "Fixed" -Message "Mail.ReadWrite permission ADDED" -Details "User must consent on next login"
            $hasMailReadWrite = $true
            $hasUserRead = $true
        }
        catch {
            Add-Result -Type "Failed" -Message "Mail.ReadWrite permission MISSING" -Details "AutoFix failed: $($_.Exception.Message)"
        }
    }
    else {
        Add-Result -Type "Failed" -Message "Mail.ReadWrite permission MISSING" -Details "Add in: App Registration > API permissions > Add > Microsoft Graph > Delegated > Mail.ReadWrite"
    }
}

if ($hasUserRead) {
    Add-Result -Type "Passed" -Message "User.Read permission is configured" -Details "Delegated permission for user profile"
}
else {
    Add-Result -Type "Failed" -Message "User.Read permission MISSING" -Details "Add in: App Registration > API permissions > Add > Microsoft Graph > Delegated > User.Read"
}

# Check 4: Redirect URIs
Write-Host "[4/6] Checking Redirect URIs..." -ForegroundColor Cyan

$publicClientUris = $app.PublicClient.RedirectUris
$hasNativeUri = $publicClientUris -contains $NATIVE_REDIRECT_URI

if ($hasNativeUri) {
    Add-Result -Type "Passed" -Message "Native client redirect URI configured" -Details $NATIVE_REDIRECT_URI
}
else {
    if ($AutoFix) {
        Write-Host "  Attempting to add redirect URI..." -ForegroundColor Yellow
        try {
            $newUris = @($publicClientUris) + @($NATIVE_REDIRECT_URI) | Select-Object -Unique
            Update-AzADApplication -ObjectId $app.Id -PublicClientRedirectUri $newUris -ErrorAction Stop
            Add-Result -Type "Fixed" -Message "Native client redirect URI ADDED" -Details $NATIVE_REDIRECT_URI
        }
        catch {
            Add-Result -Type "Warning" -Message "Native client redirect URI not configured" -Details "May work with device code flow only. Add for interactive auth: $NATIVE_REDIRECT_URI"
        }
    }
    else {
        Add-Result -Type "Warning" -Message "Native client redirect URI not configured" -Details "May work with device code flow only. Add for interactive auth: $NATIVE_REDIRECT_URI"
    }
}

# Check 5: App Roles (Optional)
Write-Host "[5/6] Checking App Roles (Optional)..." -ForegroundColor Cyan

$hasAdminRole = $false
$hasUserRole = $false

foreach ($role in $app.AppRole) {
    if ($role.Value -eq "OutlookRules.Admin") {
        $hasAdminRole = $true
    }
    if ($role.Value -eq "OutlookRules.User") {
        $hasUserRole = $true
    }
}

if ($hasAdminRole -and $hasUserRole) {
    Add-Result -Type "Passed" -Message "App roles are configured (Admin + User)" -Details "Multi-user authorization enabled"
}
elseif ($hasAdminRole -or $hasUserRole) {
    Add-Result -Type "Warning" -Message "Partial app roles configured" -Details "Both OutlookRules.Admin and OutlookRules.User recommended"
}
else {
    Add-Result -Type "Warning" -Message "No app roles configured" -Details "App roles enable multi-user authorization. Not required for single-user use."
}

# Check 6: Service Principal
Write-Host "[6/6] Checking Service Principal..." -ForegroundColor Cyan

$sp = Get-AzADServicePrincipal -Filter "AppId eq '$ClientId'" -ErrorAction SilentlyContinue

if ($sp) {
    Add-Result -Type "Passed" -Message "Service principal exists" -Details "SP ID: $($sp.Id)"

    # Check if User Assignment Required is enabled
    # Note: Az module may not expose this property directly
    Write-Host "  Service Principal ID: $($sp.Id)" -ForegroundColor Gray
}
else {
    Add-Result -Type "Warning" -Message "Service principal not found in this tenant" -Details "Will be created automatically on first sign-in"
}

# ---------------------------
# DISPLAY RESULTS
# ---------------------------
Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan
Write-Host "  VALIDATION RESULTS" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan

# Passed
if ($validationResults.Passed.Count -gt 0) {
    Write-Host "`n[PASSED] $($validationResults.Passed.Count) checks passed:" -ForegroundColor Green
    foreach ($item in $validationResults.Passed) {
        Write-Host "  [OK] $($item.Message)" -ForegroundColor Green
        if ($item.Details) {
            Write-Host "       $($item.Details)" -ForegroundColor Gray
        }
    }
}

# Fixed (if AutoFix was used)
if ($validationResults.Fixed.Count -gt 0) {
    Write-Host "`n[FIXED] $($validationResults.Fixed.Count) issues auto-fixed:" -ForegroundColor Cyan
    foreach ($item in $validationResults.Fixed) {
        Write-Host "  [FIXED] $($item.Message)" -ForegroundColor Cyan
        if ($item.Details) {
            Write-Host "          $($item.Details)" -ForegroundColor Gray
        }
    }
}

# Warnings
if ($validationResults.Warnings.Count -gt 0) {
    Write-Host "`n[WARNINGS] $($validationResults.Warnings.Count) warnings:" -ForegroundColor Yellow
    foreach ($item in $validationResults.Warnings) {
        Write-Host "  [!] $($item.Message)" -ForegroundColor Yellow
        if ($item.Details) {
            Write-Host "      $($item.Details)" -ForegroundColor Gray
        }
    }
}

# Failed
if ($validationResults.Failed.Count -gt 0) {
    Write-Host "`n[FAILED] $($validationResults.Failed.Count) critical issues:" -ForegroundColor Red
    foreach ($item in $validationResults.Failed) {
        Write-Host "  [X] $($item.Message)" -ForegroundColor Red
        if ($item.Details) {
            Write-Host "      $($item.Details)" -ForegroundColor Yellow
        }
    }
}

# ---------------------------
# OVERALL STATUS
# ---------------------------
$criticalIssues = $validationResults.Failed.Count
$warningCount = $validationResults.Warnings.Count

Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan

if ($criticalIssues -eq 0) {
    Write-Host "  APP REGISTRATION IS READY FOR USE" -ForegroundColor Green
    Write-Host ("=" * 60) -ForegroundColor Cyan

    if ($warningCount -gt 0) {
        Write-Host "`nNote: $warningCount warning(s) found but app should work." -ForegroundColor Yellow
    }

    # Save config if requested
    if ($SaveConfig) {
        $ProjectRoot = Split-Path $PSScriptRoot -Parent

        if ($ConfigProfile) {
            $envFile = Join-Path $ProjectRoot ".env.$ConfigProfile"
            $configFile = Join-Path $ProjectRoot "app-config.$ConfigProfile.json"
            $profileDisplay = "[$ConfigProfile]"
        }
        else {
            $envFile = Join-Path $ProjectRoot ".env"
            $configFile = Join-Path $ProjectRoot "app-config.json"
            $profileDisplay = "[default]"
        }

        # Save .env file
        $envContent = @"
`$ClientId = "$ClientId"
`$TenantId = "$TenantId"
"@
        Set-Content -Path $envFile -Value $envContent

        # Save JSON backup
        @{
            AppName   = $app.DisplayName
            ClientId  = $ClientId
            TenantId  = $TenantId
            Validated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Source    = "Existing app registration"
        } | ConvertTo-Json | Set-Content $configFile

        Write-Host "`nConfiguration saved $profileDisplay:" -ForegroundColor Green
        Write-Host "  $envFile" -ForegroundColor Cyan
        Write-Host "  $configFile" -ForegroundColor Gray
    }

    Write-Host "`nNext steps:" -ForegroundColor White
    if (-not $SaveConfig) {
        Write-Host "  1. Run this command with -SaveConfig to save the configuration:" -ForegroundColor Gray
        Write-Host "     .\Test-ExistingAppRegistration.ps1 -ClientId '$ClientId' -TenantId '$TenantId' -SaveConfig" -ForegroundColor Cyan
        Write-Host "  2. Or manually create .env file with:" -ForegroundColor Gray
        Write-Host "     `$ClientId = `"$ClientId`"" -ForegroundColor Cyan
        Write-Host "     `$TenantId = `"$TenantId`"" -ForegroundColor Cyan
        Write-Host "  3. Then run: .\src\Connect-OutlookRulesApp.ps1" -ForegroundColor Gray
    }
    else {
        Write-Host "  1. Run: .\src\Connect-OutlookRulesApp.ps1" -ForegroundColor Gray
        Write-Host "  2. Consent to permissions when prompted" -ForegroundColor Gray
        Write-Host "  3. Run: .\src\Manage-OutlookRules.ps1 -Operation List" -ForegroundColor Gray
    }

    exit 0
}
else {
    Write-Host "  APP REGISTRATION NEEDS CONFIGURATION" -ForegroundColor Red
    Write-Host ("=" * 60) -ForegroundColor Cyan

    Write-Host "`n$criticalIssues critical issue(s) must be resolved before using this app." -ForegroundColor Red

    if (-not $AutoFix) {
        Write-Host "`nTip: Run with -AutoFix to attempt automatic fixes:" -ForegroundColor Yellow
        Write-Host "  .\Test-ExistingAppRegistration.ps1 -ClientId '$ClientId' -TenantId '$TenantId' -AutoFix" -ForegroundColor Cyan
    }

    Write-Host "`nManual fix instructions:" -ForegroundColor Yellow
    Write-Host "  1. Go to Azure Portal > Microsoft Entra ID > App registrations" -ForegroundColor Gray
    Write-Host "  2. Find '$($app.DisplayName)'" -ForegroundColor Gray
    Write-Host "  3. Address each failed check listed above" -ForegroundColor Gray
    Write-Host "  4. Re-run this validation script" -ForegroundColor Gray

    exit 1
}
