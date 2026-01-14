<#
.SYNOPSIS
    Installs all prerequisites for Outlook Rules management scripts.
    Run this ONCE on a new machine.

.DESCRIPTION
    Installs:
    - Microsoft.Graph modules (for folder management)
    - ExchangeOnlineManagement (for inbox rules)
    - Az modules (for app registration, optional)

.PARAMETER SkipVersionPin
    Skip version pinning and install latest versions.
    Use this only if you need the latest features or fixes.

.PARAMETER UpdateModules
    Update existing modules to the pinned versions.
#>

param(
    [switch]$SkipVersionPin,
    [switch]$UpdateModules
)

Write-Host "=== Installing Outlook Rules Prerequisites ===" -ForegroundColor Cyan

# Core modules with pinned versions for stability
# These versions are tested and known to work with this project
$coreModules = @{
    "Microsoft.Graph.Authentication" = "2.24.0"
    "Microsoft.Graph.Mail"           = "2.24.0"
    "ExchangeOnlineManagement"       = "3.6.0"
}

# Optional modules for app registration (Register-OutlookRulesApp.ps1)
$optionalModules = @{
    "Az.Accounts"  = "3.0.5"
    "Az.Resources" = "7.5.0"
}

function Install-ModuleWithVersion {
    param(
        [string]$ModuleName,
        [string]$Version,
        [switch]$SkipPin,
        [switch]$Update
    )

    $installed = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1

    if ($installed) {
        if ($Update -and -not $SkipPin) {
            # Check if we need to update to pinned version
            if ($installed.Version -ne $Version) {
                Write-Host "  $ModuleName - updating from $($installed.Version) to $Version..." -ForegroundColor Yellow
                Install-Module $ModuleName -RequiredVersion $Version -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
                Write-Host "  $ModuleName - updated to $Version" -ForegroundColor Green
            } else {
                Write-Host "  $ModuleName - already at version $Version" -ForegroundColor Green
            }
        } else {
            Write-Host "  $ModuleName - already installed (v$($installed.Version))" -ForegroundColor Green
        }
    } else {
        if ($SkipPin) {
            Write-Host "  $ModuleName - installing latest..." -ForegroundColor Cyan
            Install-Module $ModuleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
        } else {
            Write-Host "  $ModuleName - installing v$Version..." -ForegroundColor Cyan
            Install-Module $ModuleName -RequiredVersion $Version -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
        }
        Write-Host "  $ModuleName - installed" -ForegroundColor Green
    }
}

Write-Host "`n[1/2] Installing core modules..." -ForegroundColor Yellow
if (-not $SkipVersionPin) {
    Write-Host "       (Using pinned versions for stability)" -ForegroundColor Gray
}

foreach ($mod in $coreModules.GetEnumerator()) {
    Install-ModuleWithVersion -ModuleName $mod.Key -Version $mod.Value -SkipPin:$SkipVersionPin -Update:$UpdateModules
}

Write-Host "`n[2/2] Installing optional modules (for app registration)..." -ForegroundColor Yellow
foreach ($mod in $optionalModules.GetEnumerator()) {
    Install-ModuleWithVersion -ModuleName $mod.Key -Version $mod.Value -SkipPin:$SkipVersionPin -Update:$UpdateModules
}

Write-Host "`n=== Prerequisites Complete ===" -ForegroundColor Green

# Display installed versions
Write-Host "`nInstalled Module Versions:" -ForegroundColor Cyan
$allModules = $coreModules.Keys + $optionalModules.Keys
foreach ($modName in $allModules) {
    $mod = Get-Module -ListAvailable -Name $modName | Sort-Object Version -Descending | Select-Object -First 1
    if ($mod) {
        $pinnedVersion = if ($coreModules[$modName]) { $coreModules[$modName] } else { $optionalModules[$modName] }
        $status = if ($mod.Version.ToString() -eq $pinnedVersion) { "(pinned)" } else { "(different from pin: $pinnedVersion)" }
        Write-Host "  $modName : $($mod.Version) $status" -ForegroundColor Gray
    }
}

Write-Host @"

Next Steps:
-----------
1. Register your app (one-time):
   .\Register-OutlookRulesApp.ps1

2. Connect to services:
   .\Connect-OutlookRulesApp.ps1

3. Run the rules setup:
   .\Manage-OutlookRules.ps1 -Operation Deploy

Alternative (without app registration):
---------------------------------------
If your tenant allows user consent, you can skip app registration:

   Connect-MgGraph -Scopes "Mail.ReadWrite" -NoWelcome
   Connect-ExchangeOnline
   .\Manage-OutlookRules.ps1 -Operation Deploy

Module Version Management:
--------------------------
To update to pinned versions:  .\Install-Prerequisites.ps1 -UpdateModules
To install latest versions:    .\Install-Prerequisites.ps1 -SkipVersionPin

"@
