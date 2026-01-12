<#
.SYNOPSIS
    Comprehensive Outlook Rules & Mailbox Management Tool.

.DESCRIPTION
    Rule Operations:
    - List:        View all current Exchange Online inbox rules
    - Show:        Show details of a specific rule
    - Export:      Export current rules to JSON config format
    - Backup:      Export rules with timestamp to backup folder
    - Import:      Import/restore rules from a backup file
    - Compare:     Compare deployed rules vs config file
    - Deploy:      Deploy rules from config file (create/update)
    - Pull:        Pull deployed rules into config file (opposite of Deploy)
    - Enable:      Enable a specific rule
    - Disable:     Disable a specific rule
    - EnableAll:   Enable all rules
    - DisableAll:  Disable all rules
    - Delete:      Delete a specific rule
    - DeleteAll:   Delete ALL rules (dangerous!)

    Folder Operations:
    - Folders:     List inbox subfolders with item counts
    - Stats:       Show mailbox statistics

    Utility Operations:
    - Validate:    Check rules for potential issues
    - Categories:  List available Outlook categories

.PARAMETER Operation
    The operation to perform

.PARAMETER RuleName
    Rule name for Enable, Disable, Delete, or Show operations

.PARAMETER ConfigPath
    Path to rules-config.json (default: ./rules-config.json)

.PARAMETER ExportPath
    Path for Export operation output (default: ./exported-rules.json)

.PARAMETER Force
    Skip confirmation prompts for destructive operations

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation List
    # Lists all current inbox rules

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation Backup
    # Creates timestamped backup in ./backups/ folder

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation Import -ExportPath ".\backups\rules-2024-01-15.json"
    # Restores rules from backup file

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation Stats
    # Shows mailbox folder statistics

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation Pull
    # Updates rules-config.json with currently deployed rules

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation Validate
    # Checks rules for potential conflicts or issues

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation DisableAll
    # Disables all inbox rules

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation DeleteAll -Force
    # Deletes ALL inbox rules (use with caution!)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet(
        "List", "Show", "Export", "Backup", "Import", "Compare", "Deploy", "Pull",
        "Enable", "Disable", "EnableAll", "DisableAll", "Delete", "DeleteAll",
        "Folders", "Stats", "Validate", "Categories"
    )]
    [string]$Operation,

    [Parameter(Mandatory = $false)]
    [string]$RuleName,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = (Join-Path $PSScriptRoot "rules-config.json"),

    [Parameter(Mandatory = $false)]
    [string]$ExportPath = (Join-Path $PSScriptRoot "exported-rules.json"),

    [switch]$Force
)

# ---------------------------
# HELPER FUNCTIONS
# ---------------------------

function Test-ExchangeConnection {
    $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*ExchangeOnline*" }
    if (-not $conn) {
        Write-Host "ERROR: Not connected to Exchange Online." -ForegroundColor Red
        Write-Host "Run: .\Connect-OutlookRulesApp.ps1" -ForegroundColor Yellow
        exit 1
    }
}

function Test-GraphConnection {
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Host "ERROR: Not connected to Microsoft Graph." -ForegroundColor Red
        Write-Host "Run: .\Connect-OutlookRulesApp.ps1" -ForegroundColor Yellow
        exit 1
    }
}

function Get-RulesList {
    Get-InboxRule -ErrorAction Stop | Sort-Object Priority
}

function Format-RuleForDisplay {
    param($Rule)

    $conditions = @()
    if ($Rule.From) { $conditions += "From: $($Rule.From -join ', ')" }
    if ($Rule.FromAddressContainsWords) { $conditions += "FromContains: $($Rule.FromAddressContainsWords -join ', ')" }
    if ($Rule.SubjectContainsWords) { $conditions += "Subject: $($Rule.SubjectContainsWords -join ', ')" }
    if ($Rule.BodyContainsWords) { $conditions += "Body: $($Rule.BodyContainsWords -join ', ')" }
    if ($Rule.SenderDomainIs) { $conditions += "Domain: $($Rule.SenderDomainIs -join ', ')" }

    $actions = @()
    if ($Rule.MoveToFolder) { $actions += "MoveTo: $($Rule.MoveToFolder)" }
    if ($Rule.DeleteMessage) { $actions += "Delete" }
    if ($Rule.MarkAsRead) { $actions += "MarkRead" }
    if ($Rule.MarkImportance) { $actions += "Importance: $($Rule.MarkImportance)" }
    if ($Rule.ApplyCategory) { $actions += "Category: $($Rule.ApplyCategory -join ', ')" }
    if ($Rule.FlagMessage) { $actions += "Flag" }
    if ($Rule.StopProcessingRules) { $actions += "StopProcessing" }

    [PSCustomObject]@{
        Priority   = $Rule.Priority
        Name       = $Rule.Name
        Enabled    = $Rule.Enabled
        Conditions = ($conditions -join "; ")
        Actions    = ($actions -join "; ")
    }
}

function Read-Config {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        Write-Host "ERROR: Config file not found: $Path" -ForegroundColor Red
        exit 1
    }

    $config = Get-Content $Path -Raw | ConvertFrom-Json
    return $config
}

function Resolve-ConfigReferences {
    param($Config, $Value)

    if ($Value -is [string] -and $Value.StartsWith("@")) {
        $path = $Value.Substring(1).Split(".")
        $resolved = $Config
        foreach ($segment in $path) {
            $resolved = $resolved.$segment
        }

        # Return appropriate property based on the reference type
        if ($resolved.addresses) { return $resolved.addresses }
        if ($resolved.domains) { return $resolved.domains }
        if ($resolved.keywords) { return $resolved.keywords }
        return $resolved
    }

    if ($Value -is [array]) {
        return $Value | ForEach-Object { Resolve-ConfigReferences -Config $Config -Value $_ }
    }

    return $Value
}

function Convert-ConfigRuleToParams {
    param($Config, $Rule)

    $params = @{}

    # Process conditions
    if ($Rule.conditions.from) {
        $params["From"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.from
    }
    if ($Rule.conditions.subjectContainsWords) {
        $params["SubjectContainsWords"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.subjectContainsWords
    }
    if ($Rule.conditions.bodyContainsWords) {
        $params["BodyContainsWords"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.bodyContainsWords
    }
    if ($Rule.conditions.senderDomainIs) {
        $params["SenderDomainIs"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.senderDomainIs
    }

    # Process actions
    if ($Rule.actions.moveToFolder) {
        # Handle noise action override
        if ($Rule.id -eq "rule-99" -and $Config.settings.noiseAction -eq "Delete") {
            $params["DeleteMessage"] = $true
            $params["StopProcessingRules"] = $true
        } else {
            $params["MoveToFolder"] = $Rule.actions.moveToFolder
        }
    }
    if ($Rule.actions.markImportance) {
        $params["MarkImportance"] = $Rule.actions.markImportance
    }
    if ($Rule.actions.markAsRead) {
        $params["MarkAsRead"] = $Rule.actions.markAsRead
    }
    if ($Rule.actions.flagMessage) {
        $params["FlagMessage"] = $Rule.actions.flagMessage
    }
    if ($Rule.actions.stopProcessingRules) {
        $params["StopProcessingRules"] = $Rule.actions.stopProcessingRules
    }
    if ($Rule.actions.assignCategories) {
        $categories = $Rule.actions.assignCategories | ForEach-Object {
            Resolve-ConfigReferences -Config $Config -Value $_
        }
        $params["ApplyCategory"] = $categories
    }
    if ($Rule.actions.deleteMessage) {
        $params["DeleteMessage"] = $Rule.actions.deleteMessage
    }

    return $params
}

function Export-RuleToConfig {
    param($Rule)

    $conditions = @{}
    $actions = @{}

    # Map conditions
    if ($Rule.From) { $conditions["from"] = @($Rule.From) }
    if ($Rule.FromAddressContainsWords) { $conditions["fromAddressContainsWords"] = @($Rule.FromAddressContainsWords) }
    if ($Rule.SubjectContainsWords) { $conditions["subjectContainsWords"] = @($Rule.SubjectContainsWords) }
    if ($Rule.SubjectOrBodyContainsWords) { $conditions["subjectOrBodyContainsWords"] = @($Rule.SubjectOrBodyContainsWords) }
    if ($Rule.BodyContainsWords) { $conditions["bodyContainsWords"] = @($Rule.BodyContainsWords) }
    if ($Rule.SenderDomainIs) { $conditions["senderDomainIs"] = @($Rule.SenderDomainIs) }
    if ($Rule.RecipientAddressContainsWords) { $conditions["recipientAddressContainsWords"] = @($Rule.RecipientAddressContainsWords) }
    if ($Rule.HasAttachment) { $conditions["hasAttachment"] = $Rule.HasAttachment }
    if ($Rule.MessageTypeMatches) { $conditions["messageTypeMatches"] = $Rule.MessageTypeMatches }
    if ($Rule.WithImportance) { $conditions["withImportance"] = $Rule.WithImportance }
    if ($Rule.FlaggedForAction) { $conditions["flaggedForAction"] = $Rule.FlaggedForAction }

    # Map actions
    if ($Rule.MoveToFolder) { $actions["moveToFolder"] = $Rule.MoveToFolder }
    if ($Rule.CopyToFolder) { $actions["copyToFolder"] = $Rule.CopyToFolder }
    if ($Rule.DeleteMessage) { $actions["deleteMessage"] = $true }
    if ($Rule.MarkAsRead) { $actions["markAsRead"] = $true }
    if ($Rule.MarkImportance) { $actions["markImportance"] = $Rule.MarkImportance }
    if ($Rule.ApplyCategory) { $actions["assignCategories"] = @($Rule.ApplyCategory) }
    if ($Rule.FlagMessage) { $actions["flagMessage"] = $true }
    if ($Rule.StopProcessingRules) { $actions["stopProcessingRules"] = $true }
    if ($Rule.ForwardTo) { $actions["forwardTo"] = @($Rule.ForwardTo) }
    if ($Rule.RedirectTo) { $actions["redirectTo"] = @($Rule.RedirectTo) }
    if ($Rule.ForwardAsAttachmentTo) { $actions["forwardAsAttachmentTo"] = @($Rule.ForwardAsAttachmentTo) }

    # Create safe ID from name
    $safeId = ($Rule.Name -replace '[^a-zA-Z0-9-]', '-').ToLower()

    return [ordered]@{
        id          = "rule-$safeId"
        name        = $Rule.Name
        description = "Exported from Exchange Online"
        enabled     = $Rule.Enabled
        priority    = $Rule.Priority
        conditions  = $conditions
        actions     = $actions
    }
}

# ---------------------------
# OPERATIONS
# ---------------------------

function Invoke-ListOperation {
    Test-ExchangeConnection

    Write-Host "`n=== Current Inbox Rules ===" -ForegroundColor Cyan

    $rules = Get-RulesList

    if ($rules.Count -eq 0) {
        Write-Host "No inbox rules found." -ForegroundColor Yellow
        return
    }

    $formatted = $rules | ForEach-Object { Format-RuleForDisplay $_ }
    $formatted | Format-Table -AutoSize -Wrap

    Write-Host "Total: $($rules.Count) rules" -ForegroundColor Gray
}

function Invoke-ShowOperation {
    param([string]$Name)

    if (-not $Name) {
        Write-Host "ERROR: -RuleName required for Show operation" -ForegroundColor Red
        exit 1
    }

    Test-ExchangeConnection

    $rule = Get-InboxRule -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $Name }

    if (-not $rule) {
        Write-Host "ERROR: Rule not found: $Name" -ForegroundColor Red
        exit 1
    }

    Write-Host "`n=== Rule Details: $Name ===" -ForegroundColor Cyan

    Write-Host "`nGeneral:" -ForegroundColor Yellow
    Write-Host "  Name:      $($rule.Name)"
    Write-Host "  Priority:  $($rule.Priority)"
    Write-Host "  Enabled:   $($rule.Enabled)"
    Write-Host "  Identity:  $($rule.Identity)"

    Write-Host "`nConditions:" -ForegroundColor Yellow
    if ($rule.From) { Write-Host "  From:                    $($rule.From -join ', ')" }
    if ($rule.FromAddressContainsWords) { Write-Host "  FromAddressContains:     $($rule.FromAddressContainsWords -join ', ')" }
    if ($rule.SubjectContainsWords) { Write-Host "  SubjectContains:         $($rule.SubjectContainsWords -join ', ')" }
    if ($rule.BodyContainsWords) { Write-Host "  BodyContains:            $($rule.BodyContainsWords -join ', ')" }
    if ($rule.SubjectOrBodyContainsWords) { Write-Host "  SubjectOrBodyContains:   $($rule.SubjectOrBodyContainsWords -join ', ')" }
    if ($rule.SenderDomainIs) { Write-Host "  SenderDomain:            $($rule.SenderDomainIs -join ', ')" }
    if ($rule.HasAttachment) { Write-Host "  HasAttachment:           $($rule.HasAttachment)" }
    if ($rule.WithImportance) { Write-Host "  WithImportance:          $($rule.WithImportance)" }

    Write-Host "`nActions:" -ForegroundColor Yellow
    if ($rule.MoveToFolder) { Write-Host "  MoveToFolder:            $($rule.MoveToFolder)" }
    if ($rule.CopyToFolder) { Write-Host "  CopyToFolder:            $($rule.CopyToFolder)" }
    if ($rule.DeleteMessage) { Write-Host "  DeleteMessage:           $($rule.DeleteMessage)" }
    if ($rule.MarkAsRead) { Write-Host "  MarkAsRead:              $($rule.MarkAsRead)" }
    if ($rule.MarkImportance) { Write-Host "  MarkImportance:          $($rule.MarkImportance)" }
    if ($rule.ApplyCategory) { Write-Host "  ApplyCategory:           $($rule.ApplyCategory -join ', ')" }
    if ($rule.FlagMessage) { Write-Host "  FlagMessage:             $($rule.FlagMessage)" }
    if ($rule.StopProcessingRules) { Write-Host "  StopProcessingRules:     $($rule.StopProcessingRules)" }
    if ($rule.ForwardTo) { Write-Host "  ForwardTo:               $($rule.ForwardTo -join ', ')" }
    if ($rule.RedirectTo) { Write-Host "  RedirectTo:              $($rule.RedirectTo -join ', ')" }
}

function Invoke-ExportOperation {
    param([string]$Path)

    Test-ExchangeConnection

    Write-Host "`n=== Exporting Rules ===" -ForegroundColor Cyan

    $rules = Get-RulesList

    if ($rules.Count -eq 0) {
        Write-Host "No rules to export." -ForegroundColor Yellow
        return
    }

    $exportedRules = $rules | ForEach-Object { Export-RuleToConfig $_ }

    $exportConfig = [ordered]@{
        "_metadata" = [ordered]@{
            description  = "Exported Outlook Inbox Rules"
            exportedAt   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            ruleCount    = $rules.Count
            source       = "Exchange Online"
        }
        "rules" = $exportedRules
    }

    $exportConfig | ConvertTo-Json -Depth 10 | Set-Content $Path

    Write-Host "Exported $($rules.Count) rules to: $Path" -ForegroundColor Green

    Write-Host "`nExported rules:" -ForegroundColor Yellow
    $rules | ForEach-Object {
        $status = if ($_.Enabled) { "[ON] " } else { "[OFF]" }
        Write-Host "  $status $($_.Priority.ToString().PadLeft(2)) - $($_.Name)"
    }
}

function Invoke-CompareOperation {
    param([string]$Path)

    Test-ExchangeConnection

    Write-Host "`n=== Comparing Deployed vs Config ===" -ForegroundColor Cyan

    $config = Read-Config -Path $Path
    $deployedRules = Get-RulesList

    $configRuleNames = $config.rules | ForEach-Object { $_.name }
    $deployedRuleNames = $deployedRules | ForEach-Object { $_.Name }

    # Find differences
    $onlyInConfig = $configRuleNames | Where-Object { $_ -notin $deployedRuleNames }
    $onlyDeployed = $deployedRuleNames | Where-Object { $_ -notin $configRuleNames }
    $inBoth = $configRuleNames | Where-Object { $_ -in $deployedRuleNames }

    Write-Host "`nIn Config Only (will be created):" -ForegroundColor Yellow
    if ($onlyInConfig.Count -eq 0) {
        Write-Host "  (none)" -ForegroundColor Gray
    } else {
        $onlyInConfig | ForEach-Object { Write-Host "  + $_" -ForegroundColor Green }
    }

    Write-Host "`nDeployed Only (not in config):" -ForegroundColor Yellow
    if ($onlyDeployed.Count -eq 0) {
        Write-Host "  (none)" -ForegroundColor Gray
    } else {
        $onlyDeployed | ForEach-Object { Write-Host "  ? $_" -ForegroundColor Magenta }
    }

    Write-Host "`nIn Both (will be updated if different):" -ForegroundColor Yellow
    if ($inBoth.Count -eq 0) {
        Write-Host "  (none)" -ForegroundColor Gray
    } else {
        $inBoth | ForEach-Object { Write-Host "  = $_" -ForegroundColor Cyan }
    }

    Write-Host "`nSummary:" -ForegroundColor Cyan
    Write-Host "  Config rules:   $($configRuleNames.Count)"
    Write-Host "  Deployed rules: $($deployedRuleNames.Count)"
    Write-Host "  To create:      $($onlyInConfig.Count)"
    Write-Host "  Not in config:  $($onlyDeployed.Count)"
    Write-Host "  To update:      $($inBoth.Count)"
}

function Invoke-DeployOperation {
    param([string]$Path, [switch]$ForceFlag)

    Test-ExchangeConnection
    Test-GraphConnection

    Write-Host "`n=== Deploying from Config ===" -ForegroundColor Cyan

    $config = Read-Config -Path $Path

    if (-not $ForceFlag) {
        Write-Host "`nThis will create/update $($config.rules.Count) rules and $($config.folders.Count) folders."
        $confirm = Read-Host "Continue? (y/N)"
        if ($confirm -ne 'y') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    # Create folders first
    Write-Host "`n--- Creating Folders ---" -ForegroundColor Yellow
    $inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox

    foreach ($folder in $config.folders) {
        $existing = Get-MgUserMailFolderChildFolder -UserId me -MailFolderId $inbox.Id -All |
            Where-Object { $_.DisplayName -eq $folder.name }

        if (-not $existing) {
            Write-Host "  Creating: $($folder.name)" -ForegroundColor Green
            New-MgUserMailFolderChildFolder -UserId me -MailFolderId $inbox.Id -DisplayName $folder.name | Out-Null
        } else {
            Write-Host "  Exists:   $($folder.name)" -ForegroundColor Gray
        }
    }

    # Create/update rules
    Write-Host "`n--- Deploying Rules ---" -ForegroundColor Yellow

    foreach ($rule in $config.rules) {
        if (-not $rule.enabled) {
            Write-Host "  Skipping (disabled in config): $($rule.name)" -ForegroundColor Gray
            continue
        }

        $params = Convert-ConfigRuleToParams -Config $config -Rule $rule

        $existing = Get-InboxRule -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $rule.name }

        if ($existing) {
            Write-Host "  Updating: $($rule.name)" -ForegroundColor Yellow
            Set-InboxRule -Identity $existing.Identity @params -Priority $rule.priority -Enabled:$true
        } else {
            Write-Host "  Creating: $($rule.name)" -ForegroundColor Green
            New-InboxRule -Name $rule.name @params -Priority $rule.priority -Enabled:$true
        }
    }

    Write-Host "`nDeployment complete." -ForegroundColor Green
}

function Invoke-EnableOperation {
    param([string]$Name, [switch]$ForceFlag)

    if (-not $Name) {
        Write-Host "ERROR: -RuleName required for Enable operation" -ForegroundColor Red
        exit 1
    }

    Test-ExchangeConnection

    $rule = Get-InboxRule -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $Name }

    if (-not $rule) {
        Write-Host "ERROR: Rule not found: $Name" -ForegroundColor Red
        exit 1
    }

    if ($rule.Enabled) {
        Write-Host "Rule '$Name' is already enabled." -ForegroundColor Yellow
        return
    }

    Set-InboxRule -Identity $rule.Identity -Enabled:$true
    Write-Host "Enabled rule: $Name" -ForegroundColor Green
}

function Invoke-DisableOperation {
    param([string]$Name, [switch]$ForceFlag)

    if (-not $Name) {
        Write-Host "ERROR: -RuleName required for Disable operation" -ForegroundColor Red
        exit 1
    }

    Test-ExchangeConnection

    $rule = Get-InboxRule -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $Name }

    if (-not $rule) {
        Write-Host "ERROR: Rule not found: $Name" -ForegroundColor Red
        exit 1
    }

    if (-not $rule.Enabled) {
        Write-Host "Rule '$Name' is already disabled." -ForegroundColor Yellow
        return
    }

    Set-InboxRule -Identity $rule.Identity -Enabled:$false
    Write-Host "Disabled rule: $Name" -ForegroundColor Green
}

function Invoke-DeleteOperation {
    param([string]$Name, [switch]$ForceFlag)

    if (-not $Name) {
        Write-Host "ERROR: -RuleName required for Delete operation" -ForegroundColor Red
        exit 1
    }

    Test-ExchangeConnection

    $rule = Get-InboxRule -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $Name }

    if (-not $rule) {
        Write-Host "ERROR: Rule not found: $Name" -ForegroundColor Red
        exit 1
    }

    if (-not $ForceFlag) {
        Write-Host "WARNING: This will delete rule '$Name'" -ForegroundColor Red
        $confirm = Read-Host "Are you sure? (y/N)"
        if ($confirm -ne 'y') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    Remove-InboxRule -Identity $rule.Identity -Confirm:$false
    Write-Host "Deleted rule: $Name" -ForegroundColor Green
}

function Invoke-FoldersOperation {
    Test-GraphConnection

    Write-Host "`n=== Inbox Folders ===" -ForegroundColor Cyan

    $inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox
    $folders = Get-MgUserMailFolderChildFolder -UserId me -MailFolderId $inbox.Id -All | Sort-Object DisplayName

    if ($folders.Count -eq 0) {
        Write-Host "No subfolders under Inbox." -ForegroundColor Yellow
        return
    }

    Write-Host ""
    $folders | ForEach-Object {
        $unread = if ($_.UnreadItemCount -gt 0) { " ($($_.UnreadItemCount) unread)" } else { "" }
        Write-Host "  $($_.DisplayName)$unread - $($_.TotalItemCount) items" -ForegroundColor $(if ($_.UnreadItemCount -gt 0) { "Yellow" } else { "Gray" })
    }

    Write-Host "`nTotal: $($folders.Count) folders" -ForegroundColor Gray
}

# ---------------------------
# NEW OPERATIONS
# ---------------------------

function Invoke-BackupOperation {
    Test-ExchangeConnection

    Write-Host "`n=== Creating Backup ===" -ForegroundColor Cyan

    # Create backups directory
    $backupDir = Join-Path $PSScriptRoot "backups"
    if (-not (Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir | Out-Null
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $backupPath = Join-Path $backupDir "rules-$timestamp.json"

    $rules = Get-RulesList

    if ($rules.Count -eq 0) {
        Write-Host "No rules to backup." -ForegroundColor Yellow
        return
    }

    $exportedRules = $rules | ForEach-Object { Export-RuleToConfig $_ }

    $exportConfig = [ordered]@{
        "_metadata" = [ordered]@{
            description  = "Outlook Rules Backup"
            backupAt     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            ruleCount    = $rules.Count
            source       = "Exchange Online"
        }
        "rules" = $exportedRules
    }

    $exportConfig | ConvertTo-Json -Depth 10 | Set-Content $backupPath

    Write-Host "Backed up $($rules.Count) rules to:" -ForegroundColor Green
    Write-Host "  $backupPath" -ForegroundColor Cyan

    # List recent backups
    $backups = Get-ChildItem $backupDir -Filter "rules-*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 5
    Write-Host "`nRecent backups:" -ForegroundColor Yellow
    $backups | ForEach-Object {
        Write-Host "  $($_.Name) - $($_.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))"
    }
}

function Invoke-ImportOperation {
    param([string]$Path, [switch]$ForceFlag)

    Test-ExchangeConnection

    Write-Host "`n=== Importing Rules ===" -ForegroundColor Cyan

    if (-not (Test-Path $Path)) {
        Write-Host "ERROR: File not found: $Path" -ForegroundColor Red
        exit 1
    }

    $importData = Get-Content $Path -Raw | ConvertFrom-Json

    if (-not $importData.rules) {
        Write-Host "ERROR: Invalid backup file - no rules found" -ForegroundColor Red
        exit 1
    }

    $ruleCount = $importData.rules.Count
    Write-Host "Found $ruleCount rules in backup file" -ForegroundColor Gray

    if ($importData._metadata) {
        Write-Host "Backup created: $($importData._metadata.backupAt)" -ForegroundColor Gray
    }

    if (-not $ForceFlag) {
        Write-Host "`nThis will create/update $ruleCount rules." -ForegroundColor Yellow
        $confirm = Read-Host "Continue? (y/N)"
        if ($confirm -ne 'y') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    foreach ($rule in $importData.rules) {
        $existing = Get-InboxRule -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $rule.name }

        # Build parameters from imported rule
        $params = @{}
        if ($rule.conditions.from) { $params["From"] = $rule.conditions.from }
        if ($rule.conditions.subjectContainsWords) { $params["SubjectContainsWords"] = $rule.conditions.subjectContainsWords }
        if ($rule.conditions.bodyContainsWords) { $params["BodyContainsWords"] = $rule.conditions.bodyContainsWords }
        if ($rule.conditions.senderDomainIs) { $params["SenderDomainIs"] = $rule.conditions.senderDomainIs }
        if ($rule.actions.moveToFolder) { $params["MoveToFolder"] = $rule.actions.moveToFolder }
        if ($rule.actions.deleteMessage) { $params["DeleteMessage"] = $true }
        if ($rule.actions.markAsRead) { $params["MarkAsRead"] = $true }
        if ($rule.actions.markImportance) { $params["MarkImportance"] = $rule.actions.markImportance }
        if ($rule.actions.assignCategories) { $params["ApplyCategory"] = $rule.actions.assignCategories }
        if ($rule.actions.flagMessage) { $params["FlagMessage"] = $true }
        if ($rule.actions.stopProcessingRules) { $params["StopProcessingRules"] = $true }

        if ($existing) {
            Write-Host "  Updating: $($rule.name)" -ForegroundColor Yellow
            Set-InboxRule -Identity $existing.Identity @params -Priority $rule.priority -Enabled:$rule.enabled
        } else {
            Write-Host "  Creating: $($rule.name)" -ForegroundColor Green
            New-InboxRule -Name $rule.name @params -Priority $rule.priority -Enabled:$rule.enabled
        }
    }

    Write-Host "`nImport complete." -ForegroundColor Green
}

function Invoke-PullOperation {
    param([string]$Path, [switch]$ForceFlag)

    Test-ExchangeConnection

    Write-Host "`n=== Pulling Deployed Rules to Config ===" -ForegroundColor Cyan

    $rules = Get-RulesList

    if ($rules.Count -eq 0) {
        Write-Host "No rules deployed to pull." -ForegroundColor Yellow
        return
    }

    if (-not $ForceFlag) {
        Write-Host "`nThis will overwrite $Path with $($rules.Count) deployed rules."
        $confirm = Read-Host "Continue? (y/N)"
        if ($confirm -ne 'y') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    $exportedRules = $rules | ForEach-Object { Export-RuleToConfig $_ }

    # Create full config structure
    $config = [ordered]@{
        "`$schema" = "./rules-schema.json"
        "_metadata" = [ordered]@{
            description  = "Outlook Inbox Rules Configuration"
            version      = "1.0.0"
            lastModified = (Get-Date).ToString("yyyy-MM-dd")
            pulledFrom   = "Exchange Online"
        }
        "settings" = [ordered]@{
            noiseAction = "Archive"
            categories = [ordered]@{
                action = "Action Required"
                metrics = "Metrics"
            }
        }
        "folders" = @()
        "senderLists" = [ordered]@{}
        "keywordLists" = [ordered]@{}
        "rules" = $exportedRules
    }

    # Extract unique folders from rules
    $folders = @()
    foreach ($rule in $exportedRules) {
        if ($rule.actions.moveToFolder -and $rule.actions.moveToFolder -like "Inbox\*") {
            $folderName = $rule.actions.moveToFolder -replace "^Inbox\\", ""
            if ($folderName -and $folderName -notin $folders) {
                $folders += $folderName
            }
        }
    }
    $config.folders = $folders | ForEach-Object {
        [ordered]@{ name = $_; description = "Auto-detected from rules"; parent = "Inbox" }
    }

    $config | ConvertTo-Json -Depth 10 | Set-Content $Path

    Write-Host "Pulled $($rules.Count) rules to: $Path" -ForegroundColor Green
    Write-Host "`nNote: Review and organize senderLists/keywordLists manually for better maintainability." -ForegroundColor Yellow
}

function Invoke-EnableAllOperation {
    param([switch]$ForceFlag)

    Test-ExchangeConnection

    $rules = Get-RulesList | Where-Object { -not $_.Enabled }

    if ($rules.Count -eq 0) {
        Write-Host "All rules are already enabled." -ForegroundColor Green
        return
    }

    Write-Host "`n=== Enable All Rules ===" -ForegroundColor Cyan
    Write-Host "Found $($rules.Count) disabled rules." -ForegroundColor Yellow

    if (-not $ForceFlag) {
        $confirm = Read-Host "Enable all? (y/N)"
        if ($confirm -ne 'y') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    foreach ($rule in $rules) {
        Set-InboxRule -Identity $rule.Identity -Enabled:$true
        Write-Host "  Enabled: $($rule.Name)" -ForegroundColor Green
    }

    Write-Host "`nEnabled $($rules.Count) rules." -ForegroundColor Green
}

function Invoke-DisableAllOperation {
    param([switch]$ForceFlag)

    Test-ExchangeConnection

    $rules = Get-RulesList | Where-Object { $_.Enabled }

    if ($rules.Count -eq 0) {
        Write-Host "All rules are already disabled." -ForegroundColor Yellow
        return
    }

    Write-Host "`n=== Disable All Rules ===" -ForegroundColor Cyan
    Write-Host "Found $($rules.Count) enabled rules." -ForegroundColor Yellow

    if (-not $ForceFlag) {
        $confirm = Read-Host "Disable all? (y/N)"
        if ($confirm -ne 'y') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    foreach ($rule in $rules) {
        Set-InboxRule -Identity $rule.Identity -Enabled:$false
        Write-Host "  Disabled: $($rule.Name)" -ForegroundColor Yellow
    }

    Write-Host "`nDisabled $($rules.Count) rules." -ForegroundColor Yellow
}

function Invoke-DeleteAllOperation {
    param([switch]$ForceFlag)

    Test-ExchangeConnection

    $rules = Get-RulesList

    if ($rules.Count -eq 0) {
        Write-Host "No rules to delete." -ForegroundColor Yellow
        return
    }

    Write-Host "`n=== DELETE ALL RULES ===" -ForegroundColor Red
    Write-Host "WARNING: This will permanently delete ALL $($rules.Count) inbox rules!" -ForegroundColor Red

    $rules | ForEach-Object {
        Write-Host "  - $($_.Name)" -ForegroundColor Gray
    }

    if (-not $ForceFlag) {
        Write-Host "`nType 'DELETE' to confirm:" -ForegroundColor Red
        $confirm = Read-Host
        if ($confirm -ne 'DELETE') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    # Backup first
    Write-Host "`nCreating backup before deletion..." -ForegroundColor Yellow
    Invoke-BackupOperation

    foreach ($rule in $rules) {
        Remove-InboxRule -Identity $rule.Identity -Confirm:$false
        Write-Host "  Deleted: $($rule.Name)" -ForegroundColor Red
    }

    Write-Host "`nDeleted $($rules.Count) rules." -ForegroundColor Red
}

function Invoke-StatsOperation {
    Test-GraphConnection

    Write-Host "`n=== Mailbox Statistics ===" -ForegroundColor Cyan

    # Get well-known folders
    $folders = @(
        @{ Name = "Inbox"; Id = "Inbox" },
        @{ Name = "Sent Items"; Id = "SentItems" },
        @{ Name = "Drafts"; Id = "Drafts" },
        @{ Name = "Deleted Items"; Id = "DeletedItems" },
        @{ Name = "Junk Email"; Id = "JunkEmail" },
        @{ Name = "Archive"; Id = "Archive" }
    )

    Write-Host "`nMain Folders:" -ForegroundColor Yellow
    $totalItems = 0
    $totalUnread = 0

    foreach ($f in $folders) {
        try {
            $folder = Get-MgUserMailFolder -UserId me -MailFolderId $f.Id -ErrorAction SilentlyContinue
            if ($folder) {
                $unread = if ($folder.UnreadItemCount -gt 0) { " ($($folder.UnreadItemCount) unread)" } else { "" }
                $color = if ($folder.UnreadItemCount -gt 0) { "Yellow" } else { "Gray" }
                Write-Host ("  {0,-20} {1,8} items{2}" -f $f.Name, $folder.TotalItemCount, $unread) -ForegroundColor $color
                $totalItems += $folder.TotalItemCount
                $totalUnread += $folder.UnreadItemCount
            }
        } catch {
            # Folder may not exist
        }
    }

    # Get inbox subfolders
    Write-Host "`nInbox Subfolders:" -ForegroundColor Yellow
    $inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox
    $subfolders = Get-MgUserMailFolderChildFolder -UserId me -MailFolderId $inbox.Id -All | Sort-Object TotalItemCount -Descending

    if ($subfolders.Count -eq 0) {
        Write-Host "  (none)" -ForegroundColor Gray
    } else {
        foreach ($sf in $subfolders) {
            $unread = if ($sf.UnreadItemCount -gt 0) { " ($($sf.UnreadItemCount) unread)" } else { "" }
            $color = if ($sf.UnreadItemCount -gt 0) { "Yellow" } else { "Gray" }
            Write-Host ("  {0,-20} {1,8} items{2}" -f $sf.DisplayName, $sf.TotalItemCount, $unread) -ForegroundColor $color
            $totalItems += $sf.TotalItemCount
            $totalUnread += $sf.UnreadItemCount
        }
    }

    Write-Host "`nSummary:" -ForegroundColor Cyan
    Write-Host "  Total items:  $totalItems"
    Write-Host "  Total unread: $totalUnread"
    Write-Host "  Inbox subfolders: $($subfolders.Count)"

    # Rule stats
    Test-ExchangeConnection
    $rules = Get-InboxRule -ErrorAction SilentlyContinue
    $enabledRules = ($rules | Where-Object { $_.Enabled }).Count
    Write-Host "  Inbox rules:  $($rules.Count) ($enabledRules enabled)"
}

function Invoke-ValidateOperation {
    Test-ExchangeConnection

    Write-Host "`n=== Validating Rules ===" -ForegroundColor Cyan

    $rules = Get-RulesList
    $issues = @()

    if ($rules.Count -eq 0) {
        Write-Host "No rules to validate." -ForegroundColor Yellow
        return
    }

    Write-Host "Checking $($rules.Count) rules..." -ForegroundColor Gray

    foreach ($rule in $rules) {
        # Check for disabled rules
        if (-not $rule.Enabled) {
            $issues += [PSCustomObject]@{
                Rule = $rule.Name
                Type = "Warning"
                Issue = "Rule is disabled"
            }
        }

        # Check for rules without conditions
        $hasCondition = $rule.From -or $rule.SubjectContainsWords -or $rule.BodyContainsWords -or
                       $rule.SenderDomainIs -or $rule.FromAddressContainsWords -or
                       $rule.SubjectOrBodyContainsWords -or $rule.HasAttachment

        if (-not $hasCondition) {
            $issues += [PSCustomObject]@{
                Rule = $rule.Name
                Type = "Error"
                Issue = "No conditions defined - rule may match all emails!"
            }
        }

        # Check for rules without actions
        $hasAction = $rule.MoveToFolder -or $rule.DeleteMessage -or $rule.MarkAsRead -or
                    $rule.ApplyCategory -or $rule.ForwardTo -or $rule.RedirectTo -or
                    $rule.MarkImportance -or $rule.FlagMessage

        if (-not $hasAction) {
            $issues += [PSCustomObject]@{
                Rule = $rule.Name
                Type = "Warning"
                Issue = "No actions defined"
            }
        }

        # Check for delete without stop processing
        if ($rule.DeleteMessage -and -not $rule.StopProcessingRules) {
            $issues += [PSCustomObject]@{
                Rule = $rule.Name
                Type = "Info"
                Issue = "Delete action without StopProcessingRules"
            }
        }

        # Check for duplicate priorities
        $samePriority = $rules | Where-Object { $_.Priority -eq $rule.Priority -and $_.Name -ne $rule.Name }
        if ($samePriority) {
            $issues += [PSCustomObject]@{
                Rule = $rule.Name
                Type = "Warning"
                Issue = "Shares priority $($rule.Priority) with: $($samePriority.Name -join ', ')"
            }
        }

        # Check for missing target folder
        if ($rule.MoveToFolder -and $rule.MoveToFolder -like "Inbox\*") {
            $folderName = $rule.MoveToFolder -replace "^Inbox\\", ""
            try {
                $inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox -ErrorAction SilentlyContinue
                if ($inbox) {
                    $targetFolder = Get-MgUserMailFolderChildFolder -UserId me -MailFolderId $inbox.Id -All |
                        Where-Object { $_.DisplayName -eq $folderName }
                    if (-not $targetFolder) {
                        $issues += [PSCustomObject]@{
                            Rule = $rule.Name
                            Type = "Error"
                            Issue = "Target folder '$folderName' does not exist"
                        }
                    }
                }
            } catch {
                # Skip folder check if Graph not connected
            }
        }
    }

    # Display results
    if ($issues.Count -eq 0) {
        Write-Host "`nNo issues found. All rules look good!" -ForegroundColor Green
    } else {
        Write-Host "`nFound $($issues.Count) issue(s):" -ForegroundColor Yellow

        $errors = $issues | Where-Object { $_.Type -eq "Error" }
        $warnings = $issues | Where-Object { $_.Type -eq "Warning" }
        $infos = $issues | Where-Object { $_.Type -eq "Info" }

        if ($errors.Count -gt 0) {
            Write-Host "`nErrors:" -ForegroundColor Red
            $errors | ForEach-Object {
                Write-Host "  [$($_.Rule)] $($_.Issue)" -ForegroundColor Red
            }
        }

        if ($warnings.Count -gt 0) {
            Write-Host "`nWarnings:" -ForegroundColor Yellow
            $warnings | ForEach-Object {
                Write-Host "  [$($_.Rule)] $($_.Issue)" -ForegroundColor Yellow
            }
        }

        if ($infos.Count -gt 0) {
            Write-Host "`nInfo:" -ForegroundColor Gray
            $infos | ForEach-Object {
                Write-Host "  [$($_.Rule)] $($_.Issue)" -ForegroundColor Gray
            }
        }
    }

    Write-Host "`nSummary: $($errors.Count) errors, $($warnings.Count) warnings, $($infos.Count) info" -ForegroundColor Cyan
}

function Invoke-CategoriesOperation {
    Test-ExchangeConnection

    Write-Host "`n=== Outlook Categories ===" -ForegroundColor Cyan

    # Get categories used in rules
    $rules = Get-RulesList
    $usedCategories = @()
    foreach ($rule in $rules) {
        if ($rule.ApplyCategory) {
            $usedCategories += $rule.ApplyCategory
        }
    }
    $usedCategories = $usedCategories | Select-Object -Unique

    Write-Host "`nCategories used in rules:" -ForegroundColor Yellow
    if ($usedCategories.Count -eq 0) {
        Write-Host "  (none)" -ForegroundColor Gray
    } else {
        $usedCategories | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Cyan
        }
    }

    Write-Host "`nNote: To manage Outlook categories, use Outlook client or OWA." -ForegroundColor Gray
    Write-Host "Categories > Manage Categories (right-click on any email)" -ForegroundColor Gray
}

# ---------------------------
# MAIN
# ---------------------------

switch ($Operation) {
    # Rule operations
    "List"       { Invoke-ListOperation }
    "Show"       { Invoke-ShowOperation -Name $RuleName }
    "Export"     { Invoke-ExportOperation -Path $ExportPath }
    "Backup"     { Invoke-BackupOperation }
    "Import"     { Invoke-ImportOperation -Path $ExportPath -ForceFlag:$Force }
    "Compare"    { Invoke-CompareOperation -Path $ConfigPath }
    "Deploy"     { Invoke-DeployOperation -Path $ConfigPath -ForceFlag:$Force }
    "Pull"       { Invoke-PullOperation -Path $ConfigPath -ForceFlag:$Force }
    "Enable"     { Invoke-EnableOperation -Name $RuleName -ForceFlag:$Force }
    "Disable"    { Invoke-DisableOperation -Name $RuleName -ForceFlag:$Force }
    "EnableAll"  { Invoke-EnableAllOperation -ForceFlag:$Force }
    "DisableAll" { Invoke-DisableAllOperation -ForceFlag:$Force }
    "Delete"     { Invoke-DeleteOperation -Name $RuleName -ForceFlag:$Force }
    "DeleteAll"  { Invoke-DeleteAllOperation -ForceFlag:$Force }

    # Folder operations
    "Folders"    { Invoke-FoldersOperation }
    "Stats"      { Invoke-StatsOperation }

    # Utility operations
    "Validate"   { Invoke-ValidateOperation }
    "Categories" { Invoke-CategoriesOperation }
}
