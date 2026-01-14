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

    Mailbox Settings Operations:
    - OutOfOffice: View/set Out-of-Office auto-reply settings
    - Forwarding:  View/set mailbox forwarding
    - JunkMail:    View/set junk mail (safe/blocked senders) settings

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

.PARAMETER ConfigProfile
    Profile name for multi-account support. Uses .env.{profile} and rules-config.{profile}.json.
    Example: -ConfigProfile personal uses rules-config.personal.json
    Example: -ConfigProfile work uses rules-config.work.json

.PARAMETER Force
    Skip confirmation prompts for destructive operations

.PARAMETER EnableAuditLog
    Enable audit logging to ./logs/ directory (requires SecurityHelpers module)

.PARAMETER OOOEnabled
    Enable or disable Out-of-Office auto-reply (use with -Operation OutOfOffice)

.PARAMETER OOOInternal
    Internal auto-reply message (use with -Operation OutOfOffice)

.PARAMETER OOOExternal
    External auto-reply message (use with -Operation OutOfOffice)

.PARAMETER OOOStartDate
    Out-of-Office start date (use with -Operation OutOfOffice)

.PARAMETER OOOEndDate
    Out-of-Office end date (use with -Operation OutOfOffice)

.PARAMETER ForwardingAddress
    Email address to forward mail to (use with -Operation Forwarding)

.PARAMETER ForwardingEnabled
    Enable/disable forwarding (use with -Operation Forwarding)

.PARAMETER DeliverToMailbox
    Keep a copy in mailbox when forwarding (use with -Operation Forwarding)

.NOTES
    Use -Verbose for detailed diagnostic output when troubleshooting.
    Example: .\Manage-OutlookRules.ps1 -Operation Deploy -Verbose

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
    .\Manage-OutlookRules.ps1 -Operation OutOfOffice
    # View current Out-of-Office settings

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation OutOfOffice -OOOEnabled $true -OOOInternal "I'm out of the office"
    # Enable Out-of-Office with a message

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation Forwarding
    # View current forwarding settings

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation Forwarding -ForwardingAddress "backup@example.com" -ForwardingEnabled $true
    # Enable forwarding to another address

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation JunkMail
    # View junk mail settings including safe/blocked senders

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation Deploy -Verbose
    # Deploy with verbose output for troubleshooting

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation Deploy -EnableAuditLog
    # Deploy with audit logging enabled

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation List -ConfigProfile personal
    # List rules for personal email account (uses .env.personal)

.EXAMPLE
    .\Manage-OutlookRules.ps1 -Operation Deploy -ConfigProfile work
    # Deploy rules for work email using rules-config.work.json
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet(
        "List", "Show", "Export", "Backup", "Import", "Compare", "Deploy", "Pull",
        "Enable", "Disable", "EnableAll", "DisableAll", "Delete", "DeleteAll",
        "Folders", "Stats", "Validate", "Categories",
        "OutOfOffice", "Forwarding", "JunkMail"
    )]
    [string]$Operation,

    [Parameter(Mandatory = $false)]
    [string]$RuleName,

    [Parameter(Mandatory = $false)]
    [string]$ConfigProfile,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath,

    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$EnableAuditLog,

    # Out-of-Office parameters
    [Parameter(Mandatory = $false)]
    [Nullable[bool]]$OOOEnabled,

    [Parameter(Mandatory = $false)]
    [string]$OOOInternal,

    [Parameter(Mandatory = $false)]
    [string]$OOOExternal,

    [Parameter(Mandatory = $false)]
    [datetime]$OOOStartDate,

    [Parameter(Mandatory = $false)]
    [datetime]$OOOEndDate,

    # Forwarding parameters
    [Parameter(Mandatory = $false)]
    [string]$ForwardingAddress,

    [Parameter(Mandatory = $false)]
    [Nullable[bool]]$ForwardingEnabled,

    [Parameter(Mandatory = $false)]
    [Nullable[bool]]$DeliverToMailbox,

    # Junk mail parameters
    [Parameter(Mandatory = $false)]
    [string[]]$SafeSenders,

    [Parameter(Mandatory = $false)]
    [string[]]$BlockedSenders,

    [Parameter(Mandatory = $false)]
    [string[]]$SafeRecipients
)

# ---------------------------
# PROFILE RESOLUTION
# ---------------------------

# Set default paths based on ConfigProfile
if ($ConfigProfile) {
    if (-not $ConfigPath) {
        $ConfigPath = Join-Path $PSScriptRoot "rules-config.$ConfigProfile.json"
    }
    if (-not $ExportPath) {
        $ExportPath = Join-Path $PSScriptRoot "exported-rules.$ConfigProfile.json"
    }
    $script:ProfileDisplay = "[$ConfigProfile]"
    Write-Host "Using profile: $ConfigProfile" -ForegroundColor Cyan
} else {
    if (-not $ConfigPath) {
        $ConfigPath = Join-Path $PSScriptRoot "rules-config.json"
    }
    if (-not $ExportPath) {
        $ExportPath = Join-Path $PSScriptRoot "exported-rules.json"
    }
    $script:ProfileDisplay = ""
}

# ---------------------------
# SECURITY MODULE
# ---------------------------

# Import security helpers module
$securityModule = Join-Path $PSScriptRoot "scripts\SecurityHelpers.psm1"
if (Test-Path $securityModule) {
    Import-Module $securityModule -Force -ErrorAction SilentlyContinue
    $script:SecurityModuleLoaded = $true
} else {
    $script:SecurityModuleLoaded = $false
    Write-Verbose "Security module not found at $securityModule - some validations will be skipped"
}

# ---------------------------
# AUDIT LOGGING
# ---------------------------

function Write-OperationLog {
    param(
        [string]$Operation,
        [string]$RuleName,
        [ValidateSet('Success', 'Failure', 'Warning')]
        [string]$Result,
        [string]$Details
    )

    if ($script:EnableAuditLog -and $script:SecurityModuleLoaded) {
        try {
            $logDir = Join-Path $PSScriptRoot "logs"
            # SECURITY: Redact sensitive data (email addresses, GUIDs) from logs
            $safeDetails = if ($Details -and (Get-Command 'Protect-SensitiveLogData' -ErrorAction SilentlyContinue)) {
                Protect-SensitiveLogData -Data $Details
            } else {
                $Details
            }
            Write-AuditLog -Operation $Operation -RuleName $RuleName -Result $Result -Details $safeDetails -LogDirectory $logDir | Out-Null
            Write-Verbose "Audit log: $Operation - $Result"
        } catch {
            Write-Verbose "Failed to write audit log: $($_.Exception.Message)"
        }
    }
}

# ---------------------------
# HELPER FUNCTIONS
# ---------------------------

function Test-ExchangeConnection {
    Write-Verbose "Checking Exchange Online connection..."
    $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*ExchangeOnline*" }
    if (-not $conn) {
        Write-Host "ERROR: Not connected to Exchange Online." -ForegroundColor Red
        $connectCmd = if ($ConfigProfile) { ".\Connect-OutlookRulesApp.ps1 -ConfigProfile $ConfigProfile" } else { ".\Connect-OutlookRulesApp.ps1" }
        Write-Host "Run: $connectCmd" -ForegroundColor Yellow
        exit 1
    }
    Write-Verbose "Exchange Online connected: $($conn.UserPrincipalName)"
}

function Test-GraphConnection {
    Write-Verbose "Checking Microsoft Graph connection..."
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Host "ERROR: Not connected to Microsoft Graph." -ForegroundColor Red
        $connectCmd = if ($ConfigProfile) { ".\Connect-OutlookRulesApp.ps1 -ConfigProfile $ConfigProfile" } else { ".\Connect-OutlookRulesApp.ps1" }
        Write-Host "Run: $connectCmd" -ForegroundColor Yellow
        exit 1
    }
    Write-Verbose "Microsoft Graph connected: $($ctx.Account) with scopes: $($ctx.Scopes -join ', ')"
}

function Get-RulesList {
    Write-Verbose "Retrieving inbox rules from Exchange Online..."
    $rules = Get-InboxRule -ErrorAction Stop | Sort-Object Priority
    Write-Verbose "Retrieved $($rules.Count) rules"
    return $rules
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
    param(
        [string]$Path,
        [switch]$SkipValidation
    )

    # Security: Validate path is within allowed directory
    if ($script:SecurityModuleLoaded) {
        try {
            $Path = Resolve-SafePath -Path $Path -BaseDirectory $PSScriptRoot
            Write-Verbose "Validated config path: $Path"
        } catch {
            Write-Host "SECURITY ERROR: $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    }

    if (-not (Test-Path $Path)) {
        Write-Host "ERROR: Config file not found: $Path" -ForegroundColor Red
        exit 1
    }

    $config = Get-Content $Path -Raw | ConvertFrom-Json

    # Validate configuration if security module is loaded
    if ($script:SecurityModuleLoaded -and -not $SkipValidation) {
        Write-Verbose "Validating configuration schema..."

        # Schema validation
        $schemaResult = Test-ConfigSchema -Config $config
        if (-not $schemaResult.Valid) {
            Write-Host "`nConfiguration Errors:" -ForegroundColor Red
            $schemaResult.Errors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
            exit 1
        }
        if ($schemaResult.Warnings.Count -gt 0) {
            Write-Host "`nConfiguration Warnings:" -ForegroundColor Yellow
            $schemaResult.Warnings | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
        }

        # Email validation
        $emailResult = Test-ConfigEmails -Config $config
        if (-not $emailResult.Valid) {
            Write-Host "`nEmail Validation Errors:" -ForegroundColor Red
            $emailResult.Errors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
            exit 1
        }

        Write-Verbose "Configuration validation passed"
    }

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

    # ===================
    # PROCESS CONDITIONS
    # ===================

    # Sender conditions
    if ($Rule.conditions.from) {
        $params["From"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.from
    }
    if ($Rule.conditions.fromAddressContainsWords) {
        $params["FromAddressContainsWords"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.fromAddressContainsWords
    }
    if ($Rule.conditions.senderDomainIs) {
        $domains = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.senderDomainIs
        # SECURITY: Validate domain format
        if ($script:SecurityModuleLoaded) {
            foreach ($domain in $domains) {
                if (-not (Test-ValidDomain -Domain $domain)) {
                    Write-Verbose "WARNING: Domain format may be invalid: $domain"
                }
            }
        }
        $params["SenderDomainIs"] = $domains
    }

    # Recipient conditions
    if ($Rule.conditions.sentTo) {
        $params["SentTo"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.sentTo
    }
    if ($Rule.conditions.recipientAddressContainsWords) {
        $params["RecipientAddressContainsWords"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.recipientAddressContainsWords
    }
    if ($null -ne $Rule.conditions.myNameInToBox) {
        $params["MyNameInToBox"] = $Rule.conditions.myNameInToBox
    }
    if ($null -ne $Rule.conditions.myNameInCcBox) {
        $params["MyNameInCcBox"] = $Rule.conditions.myNameInCcBox
    }
    if ($null -ne $Rule.conditions.myNameInToOrCcBox) {
        $params["MyNameInToOrCcBox"] = $Rule.conditions.myNameInToOrCcBox
    }
    if ($null -ne $Rule.conditions.myNameNotInToBox) {
        $params["MyNameNotInToBox"] = $Rule.conditions.myNameNotInToBox
    }
    if ($null -ne $Rule.conditions.sentOnlyToMe) {
        $params["SentOnlyToMe"] = $Rule.conditions.sentOnlyToMe
    }

    # Content conditions
    if ($Rule.conditions.subjectContainsWords) {
        $params["SubjectContainsWords"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.subjectContainsWords
    }
    if ($Rule.conditions.bodyContainsWords) {
        $params["BodyContainsWords"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.bodyContainsWords
    }
    if ($Rule.conditions.subjectOrBodyContainsWords) {
        $params["SubjectOrBodyContainsWords"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.subjectOrBodyContainsWords
    }
    if ($Rule.conditions.headerContainsWords) {
        $params["HeaderContainsWords"] = Resolve-ConfigReferences -Config $Config -Value $Rule.conditions.headerContainsWords
    }

    # Message property conditions
    if ($null -ne $Rule.conditions.hasAttachment) {
        $params["HasAttachment"] = $Rule.conditions.hasAttachment
    }
    if ($Rule.conditions.messageTypeMatches) {
        # Valid values: AutomaticReply, AutomaticForward, Encrypted, Calendaring, CalendaringResponse,
        # PermissionControlled, Voicemail, Signed, ApprovalRequest, ReadReceipt, NonDeliveryReport
        $params["MessageTypeMatches"] = $Rule.conditions.messageTypeMatches
    }
    if ($Rule.conditions.withImportance) {
        # Valid values: High, Normal, Low
        $params["WithImportance"] = $Rule.conditions.withImportance
    }
    if ($Rule.conditions.withSensitivity) {
        # Valid values: Normal, Personal, Private, CompanyConfidential
        $params["WithSensitivity"] = $Rule.conditions.withSensitivity
    }
    if ($Rule.conditions.flaggedForAction) {
        # Valid values: Any, Call, DoNotForward, FollowUp, ForYourInformation, Forward, NoResponseNecessary,
        # Read, Reply, ReplyToAll, Review
        $params["FlaggedForAction"] = $Rule.conditions.flaggedForAction
    }
    if ($Rule.conditions.hasClassification) {
        $params["HasClassification"] = $Rule.conditions.hasClassification
    }

    # Size conditions
    if ($Rule.conditions.withinSizeRangeMaximum) {
        $params["WithinSizeRangeMaximum"] = $Rule.conditions.withinSizeRangeMaximum
    }
    if ($Rule.conditions.withinSizeRangeMinimum) {
        $params["WithinSizeRangeMinimum"] = $Rule.conditions.withinSizeRangeMinimum
    }

    # Date conditions
    if ($Rule.conditions.receivedAfterDate) {
        $params["ReceivedAfterDate"] = [datetime]$Rule.conditions.receivedAfterDate
    }
    if ($Rule.conditions.receivedBeforeDate) {
        $params["ReceivedBeforeDate"] = [datetime]$Rule.conditions.receivedBeforeDate
    }

    # ===================
    # PROCESS ACTIONS
    # ===================

    # Move/Copy actions
    if ($Rule.actions.moveToFolder) {
        # Handle noise action override
        if ($Rule.id -eq "rule-99" -and $Config.settings.noiseAction -eq "Delete") {
            $params["DeleteMessage"] = $true
            $params["StopProcessingRules"] = $true
        } else {
            $params["MoveToFolder"] = $Rule.actions.moveToFolder
        }
    }
    if ($Rule.actions.copyToFolder) {
        $params["CopyToFolder"] = $Rule.actions.copyToFolder
    }
    if ($Rule.actions.deleteMessage) {
        $params["DeleteMessage"] = $Rule.actions.deleteMessage
    }
    if ($Rule.actions.softDeleteMessage) {
        $params["SoftDeleteMessage"] = $Rule.actions.softDeleteMessage
    }

    # Mark actions
    if ($Rule.actions.markAsRead) {
        $params["MarkAsRead"] = $Rule.actions.markAsRead
    }
    if ($Rule.actions.markImportance) {
        $params["MarkImportance"] = $Rule.actions.markImportance
    }
    if ($Rule.actions.flagMessage) {
        $params["FlagMessage"] = $Rule.actions.flagMessage
    }
    if ($null -ne $Rule.actions.pinMessage) {
        $params["PinMessage"] = $Rule.actions.pinMessage
    }

    # Category actions
    if ($Rule.actions.assignCategories) {
        $categories = $Rule.actions.assignCategories | ForEach-Object {
            Resolve-ConfigReferences -Config $Config -Value $_
        }
        $params["ApplyCategory"] = $categories
    }
    if ($Rule.actions.applySystemCategory) {
        # Valid values: NotDefined, Bills, Commerce, Entertainment, Health, Insurance, LiveView,
        # Lodging, Package, Shipping, Shopping, Travel, Flight, Money, HomeAutomation
        $params["ApplySystemCategory"] = $Rule.actions.applySystemCategory
    }
    if ($Rule.actions.deleteSystemCategory) {
        $params["DeleteSystemCategory"] = $Rule.actions.deleteSystemCategory
    }

    # Forward/Redirect actions (use with caution - security implications)
    # SECURITY: Validate all email addresses in forwarding actions
    if ($Rule.actions.forwardTo) {
        Write-Verbose "SECURITY: Rule '$($Rule.name)' contains forwardTo action"
        $forwardAddresses = Resolve-ConfigReferences -Config $Config -Value $Rule.actions.forwardTo
        if ($script:SecurityModuleLoaded) {
            foreach ($addr in $forwardAddresses) {
                if (-not (Test-ValidEmail -Email $addr)) {
                    Write-Host "SECURITY ERROR: Invalid email in forwardTo action: $addr" -ForegroundColor Red
                    throw "Invalid email address in forwardTo action: $addr"
                }
            }
        }
        $params["ForwardTo"] = $forwardAddresses
    }
    if ($Rule.actions.redirectTo) {
        Write-Verbose "SECURITY: Rule '$($Rule.name)' contains redirectTo action"
        $redirectAddresses = Resolve-ConfigReferences -Config $Config -Value $Rule.actions.redirectTo
        if ($script:SecurityModuleLoaded) {
            foreach ($addr in $redirectAddresses) {
                if (-not (Test-ValidEmail -Email $addr)) {
                    Write-Host "SECURITY ERROR: Invalid email in redirectTo action: $addr" -ForegroundColor Red
                    throw "Invalid email address in redirectTo action: $addr"
                }
            }
        }
        $params["RedirectTo"] = $redirectAddresses
    }
    if ($Rule.actions.forwardAsAttachmentTo) {
        Write-Verbose "SECURITY: Rule '$($Rule.name)' contains forwardAsAttachmentTo action"
        $attachmentAddresses = Resolve-ConfigReferences -Config $Config -Value $Rule.actions.forwardAsAttachmentTo
        if ($script:SecurityModuleLoaded) {
            foreach ($addr in $attachmentAddresses) {
                if (-not (Test-ValidEmail -Email $addr)) {
                    Write-Host "SECURITY ERROR: Invalid email in forwardAsAttachmentTo action: $addr" -ForegroundColor Red
                    throw "Invalid email address in forwardAsAttachmentTo action: $addr"
                }
            }
        }
        $params["ForwardAsAttachmentTo"] = $attachmentAddresses
    }

    # Processing control
    if ($Rule.actions.stopProcessingRules) {
        $params["StopProcessingRules"] = $Rule.actions.stopProcessingRules
    }

    return $params
}

function Export-RuleToConfig {
    param($Rule)

    $conditions = @{}
    $actions = @{}

    # ===================
    # MAP CONDITIONS
    # ===================

    # Sender conditions
    if ($Rule.From) { $conditions["from"] = @($Rule.From) }
    if ($Rule.FromAddressContainsWords) { $conditions["fromAddressContainsWords"] = @($Rule.FromAddressContainsWords) }
    if ($Rule.SenderDomainIs) { $conditions["senderDomainIs"] = @($Rule.SenderDomainIs) }

    # Recipient conditions
    if ($Rule.SentTo) { $conditions["sentTo"] = @($Rule.SentTo) }
    if ($Rule.RecipientAddressContainsWords) { $conditions["recipientAddressContainsWords"] = @($Rule.RecipientAddressContainsWords) }
    if ($Rule.MyNameInToBox) { $conditions["myNameInToBox"] = $Rule.MyNameInToBox }
    if ($Rule.MyNameInCcBox) { $conditions["myNameInCcBox"] = $Rule.MyNameInCcBox }
    if ($Rule.MyNameInToOrCcBox) { $conditions["myNameInToOrCcBox"] = $Rule.MyNameInToOrCcBox }
    if ($Rule.MyNameNotInToBox) { $conditions["myNameNotInToBox"] = $Rule.MyNameNotInToBox }
    if ($Rule.SentOnlyToMe) { $conditions["sentOnlyToMe"] = $Rule.SentOnlyToMe }

    # Content conditions
    if ($Rule.SubjectContainsWords) { $conditions["subjectContainsWords"] = @($Rule.SubjectContainsWords) }
    if ($Rule.BodyContainsWords) { $conditions["bodyContainsWords"] = @($Rule.BodyContainsWords) }
    if ($Rule.SubjectOrBodyContainsWords) { $conditions["subjectOrBodyContainsWords"] = @($Rule.SubjectOrBodyContainsWords) }
    if ($Rule.HeaderContainsWords) { $conditions["headerContainsWords"] = @($Rule.HeaderContainsWords) }

    # Message property conditions
    if ($Rule.HasAttachment) { $conditions["hasAttachment"] = $Rule.HasAttachment }
    if ($Rule.MessageTypeMatches) { $conditions["messageTypeMatches"] = $Rule.MessageTypeMatches }
    if ($Rule.WithImportance) { $conditions["withImportance"] = $Rule.WithImportance }
    if ($Rule.WithSensitivity) { $conditions["withSensitivity"] = $Rule.WithSensitivity }
    if ($Rule.FlaggedForAction) { $conditions["flaggedForAction"] = $Rule.FlaggedForAction }
    if ($Rule.HasClassification) { $conditions["hasClassification"] = $Rule.HasClassification }

    # Size conditions
    if ($Rule.WithinSizeRangeMaximum) { $conditions["withinSizeRangeMaximum"] = $Rule.WithinSizeRangeMaximum }
    if ($Rule.WithinSizeRangeMinimum) { $conditions["withinSizeRangeMinimum"] = $Rule.WithinSizeRangeMinimum }

    # Date conditions
    if ($Rule.ReceivedAfterDate) { $conditions["receivedAfterDate"] = $Rule.ReceivedAfterDate.ToString("yyyy-MM-dd") }
    if ($Rule.ReceivedBeforeDate) { $conditions["receivedBeforeDate"] = $Rule.ReceivedBeforeDate.ToString("yyyy-MM-dd") }

    # ===================
    # MAP ACTIONS
    # ===================

    # Move/Copy actions
    if ($Rule.MoveToFolder) { $actions["moveToFolder"] = $Rule.MoveToFolder }
    if ($Rule.CopyToFolder) { $actions["copyToFolder"] = $Rule.CopyToFolder }
    if ($Rule.DeleteMessage) { $actions["deleteMessage"] = $true }
    if ($Rule.SoftDeleteMessage) { $actions["softDeleteMessage"] = $true }

    # Mark actions
    if ($Rule.MarkAsRead) { $actions["markAsRead"] = $true }
    if ($Rule.MarkImportance) { $actions["markImportance"] = $Rule.MarkImportance }
    if ($Rule.FlagMessage) { $actions["flagMessage"] = $true }
    if ($Rule.PinMessage) { $actions["pinMessage"] = $true }

    # Category actions
    if ($Rule.ApplyCategory) { $actions["assignCategories"] = @($Rule.ApplyCategory) }
    if ($Rule.ApplySystemCategory) { $actions["applySystemCategory"] = $Rule.ApplySystemCategory }
    if ($Rule.DeleteSystemCategory) { $actions["deleteSystemCategory"] = $Rule.DeleteSystemCategory }

    # Forward/Redirect actions
    if ($Rule.ForwardTo) { $actions["forwardTo"] = @($Rule.ForwardTo) }
    if ($Rule.RedirectTo) { $actions["redirectTo"] = @($Rule.RedirectTo) }
    if ($Rule.ForwardAsAttachmentTo) { $actions["forwardAsAttachmentTo"] = @($Rule.ForwardAsAttachmentTo) }

    # Processing control
    if ($Rule.StopProcessingRules) { $actions["stopProcessingRules"] = $true }

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
    Write-Verbose "Export destination: $Path"

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
    Write-Verbose "Config path: $Path"

    $config = Read-Config -Path $Path
    Write-Verbose "Config loaded: $($config.rules.Count) rules defined"
    $deployedRules = Get-RulesList
    Write-Verbose "Deployed rules retrieved: $($deployedRules.Count) rules"

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
    Write-Verbose "Fetching Inbox folder ID..."
    $inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox
    Write-Verbose "Inbox ID: $($inbox.Id)"
    Write-Verbose "Processing $($config.folders.Count) folders from config..."

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

    $successCount = 0
    $failureCount = 0

    foreach ($rule in $config.rules) {
        if (-not $rule.enabled) {
            Write-Host "  Skipping (disabled in config): $($rule.name)" -ForegroundColor Gray
            Write-Verbose "Skipped disabled rule: $($rule.name)"
            continue
        }

        $params = Convert-ConfigRuleToParams -Config $config -Rule $rule

        $existing = Get-InboxRule -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $rule.name }

        try {
            if ($existing) {
                Write-Host "  Updating: $($rule.name)" -ForegroundColor Yellow
                Set-InboxRule -Identity $existing.Identity @params -Priority $rule.priority -Enabled:$true -ErrorAction Stop
                Write-OperationLog -Operation "Deploy-Update" -RuleName $rule.name -Result "Success" -Details "Updated existing rule"
            } else {
                Write-Host "  Creating: $($rule.name)" -ForegroundColor Green
                New-InboxRule -Name $rule.name @params -Priority $rule.priority -Enabled:$true -ErrorAction Stop
                Write-OperationLog -Operation "Deploy-Create" -RuleName $rule.name -Result "Success" -Details "Created new rule"
            }
            $successCount++
        } catch {
            Write-Host "  FAILED: $($rule.name) - $($_.Exception.Message)" -ForegroundColor Red
            Write-OperationLog -Operation "Deploy" -RuleName $rule.name -Result "Failure" -Details $_.Exception.Message
            $failureCount++
        }
    }

    Write-Host "`nDeployment complete. Success: $successCount, Failed: $failureCount" -ForegroundColor $(if ($failureCount -gt 0) { "Yellow" } else { "Green" })
    Write-OperationLog -Operation "Deploy-Summary" -RuleName $null -Result $(if ($failureCount -gt 0) { "Warning" } else { "Success" }) -Details "Deployed $successCount rules, $failureCount failures"
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
    Write-Verbose "Backup directory: $backupDir"
    if (-not (Test-Path $backupDir)) {
        Write-Verbose "Creating backup directory..."
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
    Write-Verbose "Import file path: $Path"

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

    Write-OperationLog -Operation "DeleteAll-Start" -RuleName $null -Result "Warning" -Details "Starting deletion of $($rules.Count) rules"

    $deletedCount = 0
    foreach ($rule in $rules) {
        try {
            Remove-InboxRule -Identity $rule.Identity -Confirm:$false -ErrorAction Stop
            Write-Host "  Deleted: $($rule.Name)" -ForegroundColor Red
            Write-OperationLog -Operation "Delete" -RuleName $rule.Name -Result "Success" -Details "Rule deleted"
            $deletedCount++
        } catch {
            Write-Host "  FAILED to delete: $($rule.Name) - $($_.Exception.Message)" -ForegroundColor Yellow
            Write-OperationLog -Operation "Delete" -RuleName $rule.Name -Result "Failure" -Details $_.Exception.Message
        }
    }

    Write-Host "`nDeleted $deletedCount of $($rules.Count) rules." -ForegroundColor Red
    Write-OperationLog -Operation "DeleteAll-Complete" -RuleName $null -Result "Success" -Details "Deleted $deletedCount rules"
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
    Write-Verbose "Starting rule validation..."

    $rules = Get-RulesList
    $issues = @()

    if ($rules.Count -eq 0) {
        Write-Host "No rules to validate." -ForegroundColor Yellow
        return
    }

    Write-Host "Checking $($rules.Count) rules..." -ForegroundColor Gray
    Write-Verbose "Validation checks: disabled rules, missing conditions, missing actions, duplicate priorities, folder existence"

    foreach ($rule in $rules) {
        Write-Verbose "Validating rule: $($rule.Name)"
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
# MAILBOX SETTINGS OPERATIONS
# ---------------------------

function Invoke-OutOfOfficeOperation {
    param(
        [Nullable[bool]]$Enabled,
        [string]$InternalMessage,
        [string]$ExternalMessage,
        [datetime]$StartDate,
        [datetime]$EndDate,
        [switch]$ForceFlag
    )

    Test-ExchangeConnection

    # Check if we're viewing or setting
    $isSettingOOO = $null -ne $Enabled -or $InternalMessage -or $ExternalMessage -or $StartDate -or $EndDate

    if (-not $isSettingOOO) {
        # View current settings
        Write-Host "`n=== Out-of-Office Settings ===" -ForegroundColor Cyan
        Write-Verbose "Retrieving Out-of-Office configuration..."

        try {
            $ooo = Get-MailboxAutoReplyConfiguration -Identity (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName -ErrorAction Stop

            Write-Host "`nStatus:" -ForegroundColor Yellow
            $statusColor = switch ($ooo.AutoReplyState) {
                "Enabled" { "Green" }
                "Scheduled" { "Yellow" }
                default { "Gray" }
            }
            Write-Host "  Auto-Reply State: $($ooo.AutoReplyState)" -ForegroundColor $statusColor

            if ($ooo.AutoReplyState -eq "Scheduled") {
                Write-Host "  Start Time:       $($ooo.StartTime)" -ForegroundColor Gray
                Write-Host "  End Time:         $($ooo.EndTime)" -ForegroundColor Gray
            }

            Write-Host "`nInternal Message:" -ForegroundColor Yellow
            if ($ooo.InternalMessage) {
                # Strip HTML for display
                $internalText = $ooo.InternalMessage -replace '<[^>]+>', '' -replace '&nbsp;', ' '
                Write-Host "  $internalText" -ForegroundColor Gray
            } else {
                Write-Host "  (not set)" -ForegroundColor Gray
            }

            Write-Host "`nExternal Message:" -ForegroundColor Yellow
            Write-Host "  External Audience: $($ooo.ExternalAudience)" -ForegroundColor Gray
            if ($ooo.ExternalMessage) {
                $externalText = $ooo.ExternalMessage -replace '<[^>]+>', '' -replace '&nbsp;', ' '
                Write-Host "  $externalText" -ForegroundColor Gray
            } else {
                Write-Host "  (not set)" -ForegroundColor Gray
            }

            Write-Host "`nTo modify, use parameters:" -ForegroundColor Cyan
            Write-Host "  -OOOEnabled `$true/`$false" -ForegroundColor Gray
            Write-Host "  -OOOInternal 'message'" -ForegroundColor Gray
            Write-Host "  -OOOExternal 'message'" -ForegroundColor Gray
            Write-Host "  -OOOStartDate '2024-01-15'" -ForegroundColor Gray
            Write-Host "  -OOOEndDate '2024-01-20'" -ForegroundColor Gray
        } catch {
            Write-Host "ERROR: Failed to retrieve Out-of-Office settings: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        # Set OOO settings
        Write-Host "`n=== Setting Out-of-Office ===" -ForegroundColor Cyan

        $setParams = @{}
        $identity = (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName

        if ($null -ne $Enabled) {
            if ($Enabled -and $StartDate -and $EndDate) {
                $setParams["AutoReplyState"] = "Scheduled"
                $setParams["StartTime"] = $StartDate
                $setParams["EndTime"] = $EndDate
            } elseif ($Enabled) {
                $setParams["AutoReplyState"] = "Enabled"
            } else {
                $setParams["AutoReplyState"] = "Disabled"
            }
        }

        if ($InternalMessage) {
            # SECURITY: Sanitize HTML to prevent script injection
            if ($script:SecurityModuleLoaded -and (Get-Command 'ConvertTo-SafeText' -ErrorAction SilentlyContinue)) {
                $safeInternal = ConvertTo-SafeText -Text $InternalMessage -AllowBasicFormatting
                Write-Verbose "Sanitized internal message for security"
            } else {
                $safeInternal = $InternalMessage
            }
            $setParams["InternalMessage"] = $safeInternal
        }

        if ($ExternalMessage) {
            # SECURITY: Sanitize HTML to prevent script injection
            if ($script:SecurityModuleLoaded -and (Get-Command 'ConvertTo-SafeText' -ErrorAction SilentlyContinue)) {
                $safeExternal = ConvertTo-SafeText -Text $ExternalMessage -AllowBasicFormatting
                Write-Verbose "Sanitized external message for security"
            } else {
                $safeExternal = $ExternalMessage
            }
            $setParams["ExternalMessage"] = $safeExternal
            $setParams["ExternalAudience"] = "All"  # Send to all external senders
        }

        if (-not $ForceFlag -and $setParams.Count -gt 0) {
            Write-Host "This will update Out-of-Office settings:" -ForegroundColor Yellow
            $setParams.GetEnumerator() | ForEach-Object {
                Write-Host "  $($_.Key): $($_.Value)" -ForegroundColor Gray
            }
            $confirm = Read-Host "Continue? (y/N)"
            if ($confirm -ne 'y') {
                Write-Host "Cancelled." -ForegroundColor Yellow
                return
            }
        }

        try {
            Set-MailboxAutoReplyConfiguration -Identity $identity @setParams -ErrorAction Stop
            Write-Host "Out-of-Office settings updated successfully." -ForegroundColor Green
            Write-OperationLog -Operation "OutOfOffice-Set" -RuleName $null -Result "Success" -Details ($setParams | ConvertTo-Json -Compress)
        } catch {
            Write-Host "ERROR: Failed to set Out-of-Office: $($_.Exception.Message)" -ForegroundColor Red
            Write-OperationLog -Operation "OutOfOffice-Set" -RuleName $null -Result "Failure" -Details $_.Exception.Message
        }
    }
}

function Invoke-ForwardingOperation {
    param(
        [string]$Address,
        [Nullable[bool]]$Enabled,
        [Nullable[bool]]$KeepCopy,
        [switch]$ForceFlag
    )

    Test-ExchangeConnection

    # Check if we're viewing or setting
    $isSettingForwarding = $Address -or $null -ne $Enabled -or $null -ne $KeepCopy

    if (-not $isSettingForwarding) {
        # View current settings
        Write-Host "`n=== Mailbox Forwarding Settings ===" -ForegroundColor Cyan
        Write-Verbose "Retrieving mailbox forwarding configuration..."

        try {
            $identity = (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName
            $mailbox = Get-Mailbox -Identity $identity -ErrorAction Stop

            Write-Host "`nForwarding Status:" -ForegroundColor Yellow

            if ($mailbox.ForwardingSmtpAddress) {
                Write-Host "  Forwarding To:     $($mailbox.ForwardingSmtpAddress)" -ForegroundColor Green
                $deliverStatus = if ($mailbox.DeliverToMailboxAndForward) { "Yes" } else { "No" }
                Write-Host "  Keep Copy:         $deliverStatus" -ForegroundColor Gray
            } elseif ($mailbox.ForwardingAddress) {
                Write-Host "  Forwarding To:     $($mailbox.ForwardingAddress)" -ForegroundColor Green
                $deliverStatus = if ($mailbox.DeliverToMailboxAndForward) { "Yes" } else { "No" }
                Write-Host "  Keep Copy:         $deliverStatus" -ForegroundColor Gray
            } else {
                Write-Host "  Forwarding:        Disabled" -ForegroundColor Gray
            }

            Write-Host "`nTo modify, use parameters:" -ForegroundColor Cyan
            Write-Host "  -ForwardingAddress 'user@example.com'" -ForegroundColor Gray
            Write-Host "  -ForwardingEnabled `$true/`$false" -ForegroundColor Gray
            Write-Host "  -DeliverToMailbox `$true  # Keep a copy" -ForegroundColor Gray
        } catch {
            Write-Host "ERROR: Failed to retrieve forwarding settings: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        # Set forwarding
        Write-Host "`n=== Setting Mailbox Forwarding ===" -ForegroundColor Cyan

        $identity = (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName
        $setParams = @{}

        if ($null -ne $Enabled -and -not $Enabled) {
            # Disable forwarding
            $setParams["ForwardingSmtpAddress"] = $null
            $setParams["ForwardingAddress"] = $null
            Write-Host "Disabling forwarding..." -ForegroundColor Yellow
        } elseif ($Address) {
            # SECURITY: Validate email address format before setting
            if ($script:SecurityModuleLoaded -and (Get-Command 'Test-ValidEmail' -ErrorAction SilentlyContinue)) {
                if (-not (Test-ValidEmail -Email $Address)) {
                    Write-Host "SECURITY ERROR: Invalid email address format: $Address" -ForegroundColor Red
                    Write-Host "Email must be in format: user@domain.com" -ForegroundColor Yellow
                    Write-OperationLog -Operation "Forwarding-Set" -RuleName $null -Result "Failure" -Details "Invalid email format: validation blocked"
                    return
                }
            }
            $setParams["ForwardingSmtpAddress"] = "smtp:$Address"
            Write-Host "Setting forwarding to: $Address" -ForegroundColor Yellow
        }

        if ($null -ne $KeepCopy) {
            $setParams["DeliverToMailboxAndForward"] = $KeepCopy
        }

        if (-not $ForceFlag -and $setParams.Count -gt 0) {
            Write-Host "`nWARNING: Forwarding emails has security implications." -ForegroundColor Red
            Write-Host "Settings to apply:" -ForegroundColor Yellow
            $setParams.GetEnumerator() | ForEach-Object {
                Write-Host "  $($_.Key): $($_.Value)" -ForegroundColor Gray
            }
            $confirm = Read-Host "Continue? (y/N)"
            if ($confirm -ne 'y') {
                Write-Host "Cancelled." -ForegroundColor Yellow
                return
            }
        }

        try {
            Set-Mailbox -Identity $identity @setParams -ErrorAction Stop
            Write-Host "Forwarding settings updated successfully." -ForegroundColor Green
            Write-OperationLog -Operation "Forwarding-Set" -RuleName $null -Result "Success" -Details ($setParams | ConvertTo-Json -Compress)
        } catch {
            Write-Host "ERROR: Failed to set forwarding: $($_.Exception.Message)" -ForegroundColor Red
            Write-OperationLog -Operation "Forwarding-Set" -RuleName $null -Result "Failure" -Details $_.Exception.Message
        }
    }
}

function Invoke-JunkMailOperation {
    param(
        [string[]]$AddSafeSenders,
        [string[]]$AddBlockedSenders,
        [string[]]$AddSafeRecipients,
        [switch]$ForceFlag
    )

    Test-ExchangeConnection

    # Check if we're viewing or setting
    $isSettingJunk = $AddSafeSenders -or $AddBlockedSenders -or $AddSafeRecipients

    if (-not $isSettingJunk) {
        # View current settings
        Write-Host "`n=== Junk Mail Settings ===" -ForegroundColor Cyan
        Write-Verbose "Retrieving junk mail configuration..."

        try {
            $identity = (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName
            $junk = Get-MailboxJunkEmailConfiguration -Identity $identity -ErrorAction Stop

            Write-Host "`nJunk Filter Status:" -ForegroundColor Yellow
            $enabledStatus = if ($junk.Enabled) { "Enabled" } else { "Disabled" }
            Write-Host "  Junk Filter:         $enabledStatus" -ForegroundColor $(if ($junk.Enabled) { "Green" } else { "Gray" })
            Write-Host "  Trust Contacts:      $($junk.ContactsTrusted)" -ForegroundColor Gray
            Write-Host "  Trust Safe Lists:    $($junk.TrustedListsOnly)" -ForegroundColor Gray

            Write-Host "`nSafe Senders ($($junk.TrustedSendersAndDomains.Count)):" -ForegroundColor Yellow
            if ($junk.TrustedSendersAndDomains.Count -eq 0) {
                Write-Host "  (none)" -ForegroundColor Gray
            } else {
                $junk.TrustedSendersAndDomains | Select-Object -First 10 | ForEach-Object {
                    Write-Host "  - $_" -ForegroundColor Gray
                }
                if ($junk.TrustedSendersAndDomains.Count -gt 10) {
                    Write-Host "  ... and $($junk.TrustedSendersAndDomains.Count - 10) more" -ForegroundColor Gray
                }
            }

            Write-Host "`nBlocked Senders ($($junk.BlockedSendersAndDomains.Count)):" -ForegroundColor Yellow
            if ($junk.BlockedSendersAndDomains.Count -eq 0) {
                Write-Host "  (none)" -ForegroundColor Gray
            } else {
                $junk.BlockedSendersAndDomains | Select-Object -First 10 | ForEach-Object {
                    Write-Host "  - $_" -ForegroundColor Red
                }
                if ($junk.BlockedSendersAndDomains.Count -gt 10) {
                    Write-Host "  ... and $($junk.BlockedSendersAndDomains.Count - 10) more" -ForegroundColor Gray
                }
            }

            Write-Host "`nSafe Recipients ($($junk.TrustedRecipientsAndDomains.Count)):" -ForegroundColor Yellow
            if ($junk.TrustedRecipientsAndDomains.Count -eq 0) {
                Write-Host "  (none)" -ForegroundColor Gray
            } else {
                $junk.TrustedRecipientsAndDomains | Select-Object -First 10 | ForEach-Object {
                    Write-Host "  - $_" -ForegroundColor Gray
                }
                if ($junk.TrustedRecipientsAndDomains.Count -gt 10) {
                    Write-Host "  ... and $($junk.TrustedRecipientsAndDomains.Count - 10) more" -ForegroundColor Gray
                }
            }

            Write-Host "`nTo modify, use parameters:" -ForegroundColor Cyan
            Write-Host "  -SafeSenders 'user@example.com','domain.com'" -ForegroundColor Gray
            Write-Host "  -BlockedSenders 'spam@example.com'" -ForegroundColor Gray
            Write-Host "  -SafeRecipients 'list@example.com'" -ForegroundColor Gray
        } catch {
            Write-Host "ERROR: Failed to retrieve junk mail settings: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        # Set junk mail settings
        Write-Host "`n=== Updating Junk Mail Settings ===" -ForegroundColor Cyan

        $identity = (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName

        try {
            $current = Get-MailboxJunkEmailConfiguration -Identity $identity -ErrorAction Stop
            $setParams = @{}

            if ($AddSafeSenders) {
                # SECURITY: Validate email/domain format
                if ($script:SecurityModuleLoaded) {
                    foreach ($entry in $AddSafeSenders) {
                        $isValidEmail = Test-ValidEmail -Email $entry
                        $isValidDomain = Test-ValidDomain -Domain $entry
                        if (-not $isValidEmail -and -not $isValidDomain) {
                            Write-Host "WARNING: Invalid format skipped: $entry" -ForegroundColor Yellow
                            $AddSafeSenders = $AddSafeSenders | Where-Object { $_ -ne $entry }
                        }
                    }
                }
                if ($AddSafeSenders.Count -gt 0) {
                    $newSafe = @($current.TrustedSendersAndDomains) + $AddSafeSenders | Select-Object -Unique
                    $setParams["TrustedSendersAndDomains"] = $newSafe
                    Write-Host "Adding to safe senders: $($AddSafeSenders -join ', ')" -ForegroundColor Green
                }
            }

            if ($AddBlockedSenders) {
                # SECURITY: Validate email/domain format
                if ($script:SecurityModuleLoaded) {
                    foreach ($entry in $AddBlockedSenders) {
                        $isValidEmail = Test-ValidEmail -Email $entry
                        $isValidDomain = Test-ValidDomain -Domain $entry
                        if (-not $isValidEmail -and -not $isValidDomain) {
                            Write-Host "WARNING: Invalid format skipped: $entry" -ForegroundColor Yellow
                            $AddBlockedSenders = $AddBlockedSenders | Where-Object { $_ -ne $entry }
                        }
                    }
                }
                if ($AddBlockedSenders.Count -gt 0) {
                    $newBlocked = @($current.BlockedSendersAndDomains) + $AddBlockedSenders | Select-Object -Unique
                    $setParams["BlockedSendersAndDomains"] = $newBlocked
                    Write-Host "Adding to blocked senders: $($AddBlockedSenders -join ', ')" -ForegroundColor Red
                }
            }

            if ($AddSafeRecipients) {
                # SECURITY: Validate email/domain format
                if ($script:SecurityModuleLoaded) {
                    foreach ($entry in $AddSafeRecipients) {
                        $isValidEmail = Test-ValidEmail -Email $entry
                        $isValidDomain = Test-ValidDomain -Domain $entry
                        if (-not $isValidEmail -and -not $isValidDomain) {
                            Write-Host "WARNING: Invalid format skipped: $entry" -ForegroundColor Yellow
                            $AddSafeRecipients = $AddSafeRecipients | Where-Object { $_ -ne $entry }
                        }
                    }
                }
                if ($AddSafeRecipients.Count -gt 0) {
                    $newRecipients = @($current.TrustedRecipientsAndDomains) + $AddSafeRecipients | Select-Object -Unique
                    $setParams["TrustedRecipientsAndDomains"] = $newRecipients
                    Write-Host "Adding to safe recipients: $($AddSafeRecipients -join ', ')" -ForegroundColor Green
                }
            }

            if (-not $ForceFlag -and $setParams.Count -gt 0) {
                $confirm = Read-Host "Continue? (y/N)"
                if ($confirm -ne 'y') {
                    Write-Host "Cancelled." -ForegroundColor Yellow
                    return
                }
            }

            Set-MailboxJunkEmailConfiguration -Identity $identity @setParams -ErrorAction Stop
            Write-Host "Junk mail settings updated successfully." -ForegroundColor Green
            Write-OperationLog -Operation "JunkMail-Set" -RuleName $null -Result "Success" -Details "Updated junk mail lists"
        } catch {
            Write-Host "ERROR: Failed to set junk mail settings: $($_.Exception.Message)" -ForegroundColor Red
            Write-OperationLog -Operation "JunkMail-Set" -RuleName $null -Result "Failure" -Details $_.Exception.Message
        }
    }
}

# ---------------------------
# MAIN
# ---------------------------

Write-Verbose "Operation: $Operation"
Write-Verbose "ConfigPath: $ConfigPath"
Write-Verbose "ExportPath: $ExportPath"
Write-Verbose "RuleName: $RuleName"
Write-Verbose "Force: $Force"
Write-Verbose "EnableAuditLog: $EnableAuditLog"
Write-Verbose "SecurityModuleLoaded: $($script:SecurityModuleLoaded)"

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

    # Mailbox settings operations
    "OutOfOffice" {
        Invoke-OutOfOfficeOperation -Enabled $OOOEnabled -InternalMessage $OOOInternal `
            -ExternalMessage $OOOExternal -StartDate $OOOStartDate -EndDate $OOOEndDate -ForceFlag:$Force
    }
    "Forwarding" {
        Invoke-ForwardingOperation -Address $ForwardingAddress -Enabled $ForwardingEnabled `
            -KeepCopy $DeliverToMailbox -ForceFlag:$Force
    }
    "JunkMail" {
        Invoke-JunkMailOperation -AddSafeSenders $SafeSenders -AddBlockedSenders $BlockedSenders `
            -AddSafeRecipients $SafeRecipients -ForceFlag:$Force
    }

    # Utility operations
    "Validate"   { Invoke-ValidateOperation }
    "Categories" { Invoke-CategoriesOperation }
}
