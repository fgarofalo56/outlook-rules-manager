<#
.SYNOPSIS
    Installs all prerequisites for Outlook Rules management scripts.
    Run this ONCE on a new machine.

.DESCRIPTION
    Installs:
    - Microsoft.Graph modules (for folder management)
    - ExchangeOnlineManagement (for inbox rules)
    - Az modules (for app registration, optional)
#>

Write-Host "=== Installing Outlook Rules Prerequisites ===" -ForegroundColor Cyan

# Core modules for running Setup-OutlookRules.ps1
$coreModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Mail",
    "ExchangeOnlineManagement"
)

# Optional modules for app registration (Register-OutlookRulesApp.ps1)
$optionalModules = @(
    "Az.Accounts",
    "Az.Resources"
)

Write-Host "`n[1/2] Installing core modules..." -ForegroundColor Yellow
foreach ($mod in $coreModules) {
    if (Get-Module -ListAvailable -Name $mod) {
        Write-Host "  $mod - already installed" -ForegroundColor Green
    } else {
        Write-Host "  $mod - installing..." -ForegroundColor Cyan
        Install-Module $mod -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
        Write-Host "  $mod - installed" -ForegroundColor Green
    }
}

Write-Host "`n[2/2] Installing optional modules (for app registration)..." -ForegroundColor Yellow
foreach ($mod in $optionalModules) {
    if (Get-Module -ListAvailable -Name $mod) {
        Write-Host "  $mod - already installed" -ForegroundColor Green
    } else {
        Write-Host "  $mod - installing..." -ForegroundColor Cyan
        Install-Module $mod -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
        Write-Host "  $mod - installed" -ForegroundColor Green
    }
}

Write-Host "`n=== Prerequisites Complete ===" -ForegroundColor Green
Write-Host @"

Next Steps:
-----------
1. Register your app (one-time):
   .\Register-OutlookRulesApp.ps1

2. Connect to services:
   .\Connect-OutlookRulesApp.ps1

3. Run the rules setup:
   .\Setup-OutlookRules.ps1

Alternative (without app registration):
---------------------------------------
If your tenant allows user consent, you can skip app registration:

   Connect-MgGraph -Scopes "Mail.ReadWrite" -NoWelcome
   Connect-ExchangeOnline
   .\Setup-OutlookRules.ps1

"@
