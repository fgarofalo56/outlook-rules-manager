<#
.SYNOPSIS
    Manages user authorization for the Outlook Rules Manager application.
    Implements multi-tier access control: Owners > Admins > Users.

.DESCRIPTION
    This script provides operations to manage who can use the Outlook Rules Manager:

    - Setup:        Initial security setup (enable assignment required, assign creator as admin)
    - List:         List all authorized users and their roles
    - AddAdmin:     Add a user as Administrator (can manage other users)
    - AddUser:      Add a user as standard User (can only manage own mailbox)
    - Remove:       Remove a user's authorization
    - Check:        Check if a specific user is authorized
    - Status:       Show current authorization configuration status

    SECURITY MODEL:
    ---------------
    Tier 1 - Owners:  Service Principal owners in Azure AD (manage admins)
    Tier 2 - Admins:  OutlookRules.Admin role (can add/remove users + use app)
    Tier 3 - Users:   OutlookRules.User role (can only use app for own mailbox)

.PARAMETER Operation
    The operation to perform

.PARAMETER UserPrincipalName
    User's email/UPN for Add/Remove/Check operations

.PARAMETER ConfigProfile
    Profile name for multi-account support (loads .env.{profile})

.PARAMETER Force
    Skip confirmation prompts

.EXAMPLE
    .\Manage-AppAuthorization.ps1 -Operation Setup
    # Initial security setup - enables user assignment required and assigns creator as admin

.EXAMPLE
    .\Manage-AppAuthorization.ps1 -Operation List
    # Lists all authorized users and their roles

.EXAMPLE
    .\Manage-AppAuthorization.ps1 -Operation AddAdmin -UserPrincipalName "admin@example.com"
    # Adds a user as Administrator

.EXAMPLE
    .\Manage-AppAuthorization.ps1 -Operation AddUser -UserPrincipalName "user@example.com"
    # Adds a user as standard User

.EXAMPLE
    .\Manage-AppAuthorization.ps1 -Operation Remove -UserPrincipalName "user@example.com"
    # Removes a user's authorization

.EXAMPLE
    .\Manage-AppAuthorization.ps1 -Operation Check -UserPrincipalName "user@example.com"
    # Checks if a user is authorized and shows their role

.EXAMPLE
    .\Manage-AppAuthorization.ps1 -Operation Status
    # Shows current authorization configuration

.NOTES
    Requires: Microsoft.Graph.Applications module
    Must be run by someone with admin rights (Service Principal owner or OutlookRules.Admin)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("Setup", "List", "AddAdmin", "AddUser", "Remove", "Check", "Status")]
    [string]$Operation,

    [Parameter(Mandatory = $false)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$ConfigProfile,

    [switch]$Force
)

# ---------------------------
# CONSTANTS - App Role IDs
# ---------------------------
$script:AdminRoleId = "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f"
$script:UserRoleId = "a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d"

# ---------------------------
# PROFILE RESOLUTION
# ---------------------------
if ($ConfigProfile) {
    $envFile = Join-Path $PSScriptRoot ".env.$ConfigProfile"
    $profileDisplay = "[$ConfigProfile]"
} else {
    $envFile = Join-Path $PSScriptRoot ".env"
    $profileDisplay = "[default]"
}

# ---------------------------
# LOAD CONFIGURATION
# ---------------------------
if (Test-Path $envFile) {
    $envContent = Get-Content $envFile
    foreach ($line in $envContent) {
        $line = $line.Trim()
        if ($line -and -not $line.StartsWith("#")) {
            if ($line -match '^\$?(\w+)\s*=\s*"?([^"]+)"?$') {
                $key = $matches[1]
                $value = $matches[2]
                Set-Variable -Name $key -Value $value -Scope Script
            }
        }
    }
    Write-Host "Loaded config from $(Split-Path $envFile -Leaf) $profileDisplay" -ForegroundColor Green
} else {
    Write-Host "ERROR: Configuration not found: $envFile" -ForegroundColor Red
    Write-Host "Run Register-OutlookRulesApp.ps1 first to create the app registration." -ForegroundColor Yellow
    exit 1
}

if (-not $ClientId) {
    Write-Host "ERROR: ClientId not found in configuration" -ForegroundColor Red
    exit 1
}

# ---------------------------
# HELPER FUNCTIONS
# ---------------------------

function Test-GraphConnection {
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Host "ERROR: Not connected to Microsoft Graph." -ForegroundColor Red
        Write-Host "Run: Connect-MgGraph -Scopes 'Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All','User.Read.All'" -ForegroundColor Yellow
        exit 1
    }
    return $ctx
}

function Get-ServicePrincipalByAppId {
    param([string]$AppId)

    try {
        $sp = Get-MgServicePrincipal -Filter "AppId eq '$AppId'" -ErrorAction Stop
        return $sp
    } catch {
        Write-Host "ERROR: Could not find service principal for app: $AppId" -ForegroundColor Red
        Write-Host "The app may not have been used yet. Try signing in first." -ForegroundColor Yellow
        return $null
    }
}

function Get-AppRoleDisplayName {
    param([string]$RoleId)

    switch ($RoleId) {
        $script:AdminRoleId { return "Administrator (OutlookRules.Admin)" }
        $script:UserRoleId { return "User (OutlookRules.User)" }
        default { return "Unknown Role ($RoleId)" }
    }
}

# ---------------------------
# OPERATION: Setup
# ---------------------------
function Invoke-Setup {
    Write-Host "`n=== Authorization Setup $profileDisplay ===" -ForegroundColor Cyan

    $ctx = Test-GraphConnection
    Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Gray

    $sp = Get-ServicePrincipalByAppId -AppId $ClientId
    if (-not $sp) { exit 1 }

    Write-Host "`nService Principal: $($sp.DisplayName)" -ForegroundColor White
    Write-Host "  ID: $($sp.Id)" -ForegroundColor Gray
    Write-Host "  App ID: $($sp.AppId)" -ForegroundColor Gray

    # Step 1: Enable User Assignment Required
    Write-Host "`n[1/2] Enabling User Assignment Required..." -ForegroundColor Yellow

    if ($sp.AppRoleAssignmentRequired) {
        Write-Host "  Already enabled" -ForegroundColor Green
    } else {
        try {
            Update-MgServicePrincipal -ServicePrincipalId $sp.Id -AppRoleAssignmentRequired:$true -ErrorAction Stop
            Write-Host "  Enabled successfully" -ForegroundColor Green
        } catch {
            Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  You may need to enable this manually in Azure Portal" -ForegroundColor Yellow
        }
    }

    # Step 2: Assign current user as Admin
    Write-Host "`n[2/2] Assigning current user as Administrator..." -ForegroundColor Yellow

    $currentUser = Get-MgUser -Filter "userPrincipalName eq '$($ctx.Account)'" -ErrorAction SilentlyContinue
    if (-not $currentUser) {
        # Try by mail
        $currentUser = Get-MgUser -Filter "mail eq '$($ctx.Account)'" -ErrorAction SilentlyContinue
    }

    if ($currentUser) {
        # Check if already assigned
        $existingAssignment = Get-MgUserAppRoleAssignment -UserId $currentUser.Id -ErrorAction SilentlyContinue |
            Where-Object { $_.ResourceId -eq $sp.Id }

        if ($existingAssignment) {
            $roleName = Get-AppRoleDisplayName -RoleId $existingAssignment.AppRoleId
            Write-Host "  Already assigned: $roleName" -ForegroundColor Green
        } else {
            try {
                $params = @{
                    PrincipalId = $currentUser.Id
                    ResourceId = $sp.Id
                    AppRoleId = $script:AdminRoleId
                }
                New-MgUserAppRoleAssignment -UserId $currentUser.Id -BodyParameter $params -ErrorAction Stop | Out-Null
                Write-Host "  Assigned as Administrator" -ForegroundColor Green
            } catch {
                Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "  Could not find current user in directory" -ForegroundColor Yellow
        Write-Host "  You may need to assign yourself manually" -ForegroundColor Yellow
    }

    Write-Host "`n=== Setup Complete ===" -ForegroundColor Green
    Write-Host "Run: .\Manage-AppAuthorization.ps1 -Operation Status" -ForegroundColor Gray
}

# ---------------------------
# OPERATION: List
# ---------------------------
function Invoke-List {
    Write-Host "`n=== Authorized Users $profileDisplay ===" -ForegroundColor Cyan

    Test-GraphConnection | Out-Null

    $sp = Get-ServicePrincipalByAppId -AppId $ClientId
    if (-not $sp) { exit 1 }

    $assignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue

    if (-not $assignments -or $assignments.Count -eq 0) {
        Write-Host "`nNo users are currently authorized." -ForegroundColor Yellow
        Write-Host "Run: .\Manage-AppAuthorization.ps1 -Operation Setup" -ForegroundColor Gray
        return
    }

    Write-Host "`nApplication: $($sp.DisplayName)" -ForegroundColor White
    Write-Host "User Assignment Required: $($sp.AppRoleAssignmentRequired)" -ForegroundColor $(if ($sp.AppRoleAssignmentRequired) { 'Green' } else { 'Yellow' })
    Write-Host ""

    $admins = @()
    $users = @()

    foreach ($assignment in $assignments) {
        $roleDisplay = Get-AppRoleDisplayName -RoleId $assignment.AppRoleId
        $entry = [PSCustomObject]@{
            DisplayName = $assignment.PrincipalDisplayName
            Type = $assignment.PrincipalType
            Role = $roleDisplay
            AssignedDate = $assignment.CreatedDateTime
        }

        if ($assignment.AppRoleId -eq $script:AdminRoleId) {
            $admins += $entry
        } else {
            $users += $entry
        }
    }

    if ($admins.Count -gt 0) {
        Write-Host "ADMINISTRATORS ($($admins.Count)):" -ForegroundColor Cyan
        $admins | Format-Table -AutoSize
    }

    if ($users.Count -gt 0) {
        Write-Host "USERS ($($users.Count)):" -ForegroundColor Blue
        $users | Format-Table -AutoSize
    }

    Write-Host "Total authorized: $($assignments.Count)" -ForegroundColor Gray
}

# ---------------------------
# OPERATION: AddAdmin / AddUser
# ---------------------------
function Invoke-AddUser {
    param(
        [string]$UPN,
        [string]$RoleId,
        [string]$RoleName
    )

    if (-not $UPN) {
        Write-Host "ERROR: -UserPrincipalName is required for this operation" -ForegroundColor Red
        exit 1
    }

    Write-Host "`n=== Add $RoleName $profileDisplay ===" -ForegroundColor Cyan

    Test-GraphConnection | Out-Null

    $sp = Get-ServicePrincipalByAppId -AppId $ClientId
    if (-not $sp) { exit 1 }

    # Find the user
    $user = Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction SilentlyContinue
    if (-not $user) {
        $user = Get-MgUser -Filter "mail eq '$UPN'" -ErrorAction SilentlyContinue
    }

    if (-not $user) {
        Write-Host "ERROR: User not found: $UPN" -ForegroundColor Red
        exit 1
    }

    Write-Host "User: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor White
    Write-Host "Role: $RoleName" -ForegroundColor White

    # Check if already assigned
    $existingAssignment = Get-MgUserAppRoleAssignment -UserId $user.Id -ErrorAction SilentlyContinue |
        Where-Object { $_.ResourceId -eq $sp.Id }

    if ($existingAssignment) {
        $currentRole = Get-AppRoleDisplayName -RoleId $existingAssignment.AppRoleId
        Write-Host "`nUser is already assigned: $currentRole" -ForegroundColor Yellow

        if (-not $Force) {
            $response = Read-Host "Do you want to change their role? (y/N)"
            if ($response -ne 'y') {
                Write-Host "Cancelled." -ForegroundColor Gray
                return
            }
        }

        # Remove existing assignment
        Remove-MgUserAppRoleAssignment -UserId $user.Id -AppRoleAssignmentId $existingAssignment.Id -ErrorAction SilentlyContinue
    }

    # Create new assignment
    try {
        $params = @{
            PrincipalId = $user.Id
            ResourceId = $sp.Id
            AppRoleId = $RoleId
        }
        New-MgUserAppRoleAssignment -UserId $user.Id -BodyParameter $params -ErrorAction Stop | Out-Null
        Write-Host "`nSuccessfully assigned $($user.DisplayName) as $RoleName" -ForegroundColor Green
    } catch {
        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ---------------------------
# OPERATION: Remove
# ---------------------------
function Invoke-Remove {
    param([string]$UPN)

    if (-not $UPN) {
        Write-Host "ERROR: -UserPrincipalName is required for this operation" -ForegroundColor Red
        exit 1
    }

    Write-Host "`n=== Remove User Authorization $profileDisplay ===" -ForegroundColor Cyan

    Test-GraphConnection | Out-Null

    $sp = Get-ServicePrincipalByAppId -AppId $ClientId
    if (-not $sp) { exit 1 }

    # Find the user
    $user = Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction SilentlyContinue
    if (-not $user) {
        $user = Get-MgUser -Filter "mail eq '$UPN'" -ErrorAction SilentlyContinue
    }

    if (-not $user) {
        Write-Host "ERROR: User not found: $UPN" -ForegroundColor Red
        exit 1
    }

    # Find their assignment
    $assignment = Get-MgUserAppRoleAssignment -UserId $user.Id -ErrorAction SilentlyContinue |
        Where-Object { $_.ResourceId -eq $sp.Id }

    if (-not $assignment) {
        Write-Host "User is not currently authorized: $UPN" -ForegroundColor Yellow
        return
    }

    $roleName = Get-AppRoleDisplayName -RoleId $assignment.AppRoleId
    Write-Host "User: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor White
    Write-Host "Current Role: $roleName" -ForegroundColor White

    if (-not $Force) {
        $response = Read-Host "`nAre you sure you want to remove this user's authorization? (y/N)"
        if ($response -ne 'y') {
            Write-Host "Cancelled." -ForegroundColor Gray
            return
        }
    }

    try {
        Remove-MgUserAppRoleAssignment -UserId $user.Id -AppRoleAssignmentId $assignment.Id -ErrorAction Stop
        Write-Host "`nSuccessfully removed authorization for $($user.DisplayName)" -ForegroundColor Green
    } catch {
        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ---------------------------
# OPERATION: Check
# ---------------------------
function Invoke-Check {
    param([string]$UPN)

    if (-not $UPN) {
        Write-Host "ERROR: -UserPrincipalName is required for this operation" -ForegroundColor Red
        exit 1
    }

    Write-Host "`n=== Check User Authorization $profileDisplay ===" -ForegroundColor Cyan

    Test-GraphConnection | Out-Null

    $sp = Get-ServicePrincipalByAppId -AppId $ClientId
    if (-not $sp) { exit 1 }

    # Find the user
    $user = Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction SilentlyContinue
    if (-not $user) {
        $user = Get-MgUser -Filter "mail eq '$UPN'" -ErrorAction SilentlyContinue
    }

    if (-not $user) {
        Write-Host "ERROR: User not found: $UPN" -ForegroundColor Red
        exit 1
    }

    Write-Host "User: $($user.DisplayName)" -ForegroundColor White
    Write-Host "UPN: $($user.UserPrincipalName)" -ForegroundColor Gray
    Write-Host "ID: $($user.Id)" -ForegroundColor Gray
    Write-Host ""

    # Check assignment
    $assignment = Get-MgUserAppRoleAssignment -UserId $user.Id -ErrorAction SilentlyContinue |
        Where-Object { $_.ResourceId -eq $sp.Id }

    if ($assignment) {
        $roleName = Get-AppRoleDisplayName -RoleId $assignment.AppRoleId
        Write-Host "AUTHORIZED: Yes" -ForegroundColor Green
        Write-Host "Role: $roleName" -ForegroundColor Green
        Write-Host "Assigned: $($assignment.CreatedDateTime)" -ForegroundColor Gray
    } else {
        Write-Host "AUTHORIZED: No" -ForegroundColor Red
        if (-not $sp.AppRoleAssignmentRequired) {
            Write-Host "Note: User Assignment Required is disabled - all users can access!" -ForegroundColor Yellow
        }
    }
}

# ---------------------------
# OPERATION: Status
# ---------------------------
function Invoke-Status {
    Write-Host "`n=== Authorization Status $profileDisplay ===" -ForegroundColor Cyan

    $ctx = Test-GraphConnection
    Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Gray

    $sp = Get-ServicePrincipalByAppId -AppId $ClientId
    if (-not $sp) { exit 1 }

    Write-Host "`nApplication: $($sp.DisplayName)" -ForegroundColor White
    Write-Host "  Service Principal ID: $($sp.Id)" -ForegroundColor Gray
    Write-Host "  Application ID: $($sp.AppId)" -ForegroundColor Gray

    # Check User Assignment Required
    Write-Host ""
    if ($sp.AppRoleAssignmentRequired) {
        Write-Host "User Assignment Required: ENABLED" -ForegroundColor Green
        Write-Host "  Only explicitly assigned users can access this application" -ForegroundColor Gray
    } else {
        Write-Host "User Assignment Required: DISABLED" -ForegroundColor Red
        Write-Host "  WARNING: Any user who can authenticate can access this application!" -ForegroundColor Yellow
        Write-Host "  Run: .\Manage-AppAuthorization.ps1 -Operation Setup" -ForegroundColor Yellow
    }

    # List app roles
    Write-Host "`nApp Roles Defined:" -ForegroundColor White
    foreach ($role in $sp.AppRoles) {
        Write-Host "  - $($role.DisplayName) ($($role.Value))" -ForegroundColor Gray
    }

    # Count assignments
    $assignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
    $adminCount = ($assignments | Where-Object { $_.AppRoleId -eq $script:AdminRoleId }).Count
    $userCount = ($assignments | Where-Object { $_.AppRoleId -eq $script:UserRoleId }).Count

    Write-Host "`nAuthorized Users:" -ForegroundColor White
    Write-Host "  Administrators: $adminCount" -ForegroundColor Cyan
    Write-Host "  Users: $userCount" -ForegroundColor Blue
    Write-Host "  Total: $($assignments.Count)" -ForegroundColor Gray

    # List owners
    Write-Host "`nService Principal Owners:" -ForegroundColor White
    $owners = Get-MgServicePrincipalOwner -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
    if ($owners) {
        foreach ($owner in $owners) {
            Write-Host "  - $($owner.AdditionalProperties.displayName) ($($owner.AdditionalProperties.userPrincipalName))" -ForegroundColor Gray
        }
    } else {
        Write-Host "  (No owners found or unable to retrieve)" -ForegroundColor Yellow
    }
}

# ---------------------------
# MAIN EXECUTION
# ---------------------------

Write-Host "`n=== Outlook Rules Manager - Authorization ===" -ForegroundColor Cyan

switch ($Operation) {
    "Setup" {
        Invoke-Setup
    }
    "List" {
        Invoke-List
    }
    "AddAdmin" {
        Invoke-AddUser -UPN $UserPrincipalName -RoleId $script:AdminRoleId -RoleName "Administrator"
    }
    "AddUser" {
        Invoke-AddUser -UPN $UserPrincipalName -RoleId $script:UserRoleId -RoleName "User"
    }
    "Remove" {
        Invoke-Remove -UPN $UserPrincipalName
    }
    "Check" {
        Invoke-Check -UPN $UserPrincipalName
    }
    "Status" {
        Invoke-Status
    }
}
