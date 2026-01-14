
<#
.SYNOPSIS
  Creates Inbox subfolders (via Microsoft Graph) and server-side Inbox rules (via Exchange Online)
  tailored for priority senders, action items, metrics/Connect, leadership comms, alerts, and noise.

.NOTES
  Requires: Microsoft.Graph, ExchangeOnlineManagement
  Sign-in: Connect-MgGraph -Scopes "Mail.ReadWrite"; Connect-ExchangeOnline
#>

# ---------------------------
# CONFIG: adjust to your preferences
# ---------------------------
$NoiseAction = "Archive"   # "Archive" or "Delete"
$Category_Action = "Action Required"
$Category_Metrics = "Metrics"

# Priority senders (high signal sources)
# NOTE: Replace these example addresses with your actual priority senders
# For production use, configure senders in rules-config.json (gitignored)
$PrioritySenders = @(
    "manager@example.com",              # Your manager
    "skip.level@example.com",           # Skip-level manager
    "colleague1@example.com",           # Key colleague
    "colleague2@example.com",           # Key colleague
    "external.contact@contoso.com"      # External contact
)

# Keywords per rule
$Keywords_Action = @("Action", "Approval", "Response Needed", "Due", "Deadline", "Sign-off", "Decision Required")
$Keywords_Leadership = @("Leadership", "Executive", "LT", "Review", "Staff Meeting")
$Keywords_Metrics = @("Connect", "ACR", "Performance", "Impact", "KPI", "Attainment", "ADS", "Azure Consumption", "QBR", "Scorecard")
$Keywords_Alerts = @("Alert", "Notification", "Digest")

# Noise domains (expand as needed)
$NoiseDomains = @(
    "news.microsoft.com",
    "events.microsoft.com",
    "linkedin.com",
    "updates.mail.*",
    "notifications.*",
    "mailer.*"
)

# Target folders to ensure under Inbox
$TargetFolders = @("Priority", "Action Required", "Metrics", "Leadership", "Alerts", "Low Priority")

# ---------------------------
# GRAPH: ensure folders exist under Inbox
# ---------------------------
function Ensure-InboxFolder {
    param(
        [Parameter(Mandatory = $true)][string]$FolderName,
        [Parameter(Mandatory = $true)][string]$InboxId
    )
    $existing = Get-MgUserMailFolderChildFolder -UserId me -MailFolderId $InboxId -All |
    Where-Object { $_.DisplayName -eq $FolderName }
    if (-not $existing) {
        Write-Host "Creating folder '$FolderName' under Inbox..." -ForegroundColor Cyan
        New-MgUserMailFolderChildFolder -UserId me -MailFolderId $InboxId -DisplayName $FolderName | Out-Null
    }
    else {
        Write-Host "Folder '$FolderName' already exists." -ForegroundColor Green
    }
}

# Connect if not already
try {
    if (-not (Get-MgContext)) { Connect-MgGraph -Scopes "Mail.ReadWrite" -NoWelcome }
}
catch { throw "Graph connection failed: $($_.Exception.Message)" }

$inbox = Get-MgUserMailFolder -UserId me -MailFolderId Inbox
if (-not $inbox) { throw "Unable to resolve Inbox via Graph." }

foreach ($f in $TargetFolders) { Ensure-InboxFolder -FolderName $f -InboxId $inbox.Id }

# ---------------------------
# EXCHANGE ONLINE: create/update rules
# ---------------------------
try {
    # Connect if not already
    $exoConn = Get-ConnectionInformation | Where-Object { $_.Name -like "*ExchangeOnline*" }
    if (-not $exoConn) { Connect-ExchangeOnline }
}
catch { throw "Exchange Online connection failed: $($_.Exception.Message)" }

function Ensure-Rule {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [hashtable]$CreateParams,
        [hashtable]$UpdateParams,
        [int]$Priority = 1
    )
    $rule = Get-InboxRule -Mailbox $env:USERNAME -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $Name }
    if ($rule) {
        Write-Host "Updating rule: $Name" -ForegroundColor Yellow
        Set-InboxRule -Identity $rule.Identity @UpdateParams -Priority $Priority -Enabled:$true
    }
    else {
        Write-Host "Creating rule: $Name" -ForegroundColor Cyan
        New-InboxRule -Name $Name @CreateParams -Priority $Priority -Enabled:$true
    }
}

# Rule 01: Priority senders -> Inbox\Priority (High importance, stop processing)
$rule01ParamsCreate = @{
    From                = $PrioritySenders
    MoveToFolder        = "Inbox\Priority"
    MarkImportance      = "High"
    StopProcessingRules = $true
}
$rule01ParamsUpdate = $rule01ParamsCreate
Ensure-Rule -Name "01 - Priority Senders" -CreateParams $rule01ParamsCreate -UpdateParams $rule01ParamsUpdate -Priority 1

# Rule 02: Action Required -> Inbox\Action Required (Category, High importance, Flag)
$rule02ParamsCreate = @{
    SubjectContainsWords = $Keywords_Action
    AssignCategories     = @($Category_Action)
    MoveToFolder         = "Inbox\Action Required"
    MarkImportance       = "High"
    FlagMessage          = $true
}
$rule02ParamsUpdate = $rule02ParamsCreate
Ensure-Rule -Name "02 - Action Required" -CreateParams $rule02ParamsCreate -UpdateParams $rule02ParamsUpdate -Priority 2

# Rule 03: Connect & Metrics -> Inbox\Metrics (Category, Flag)
$rule03ParamsCreate = @{
    SubjectContainsWords = $Keywords_Metrics
    BodyContainsWords    = $Keywords_Metrics
    AssignCategories     = @($Category_Metrics)
    MoveToFolder         = "Inbox\Metrics"
    FlagMessage          = $true
}
$rule03ParamsUpdate = $rule03ParamsCreate
Ensure-Rule -Name "03 - Connect & Metrics" -CreateParams $rule03ParamsCreate -UpdateParams $rule03ParamsUpdate -Priority 3

# Rule 04: Leadership & Exec -> Inbox\Leadership (High importance)
$rule04ParamsCreate = @{
    SubjectContainsWords = $Keywords_Leadership
    MoveToFolder         = "Inbox\Leadership"
    MarkImportance       = "High"
}
$rule04ParamsUpdate = $rule04ParamsCreate
Ensure-Rule -Name "04 - Leadership & Exec" -CreateParams $rule04ParamsCreate -UpdateParams $rule04ParamsUpdate -Priority 4

# Rule 05: Alerts & Notifications -> Inbox\Alerts (Mark as read)
$rule05ParamsCreate = @{
    SubjectContainsWords = $Keywords_Alerts
    MoveToFolder         = "Inbox\Alerts"
    MarkAsRead           = $true
}
$rule05ParamsUpdate = $rule05ParamsCreate
Ensure-Rule -Name "05 - Alerts & Notifications" -CreateParams $rule05ParamsCreate -UpdateParams $rule05ParamsUpdate -Priority 5

# Rule 99: Noise filter -> Inbox\Low Priority (or Delete)
if ($NoiseAction -eq "Delete") {
    $rule99ParamsCreate = @{
        SenderDomainIs      = $NoiseDomains
        DeleteMessage       = $true
        StopProcessingRules = $true
    }
    $rule99ParamsUpdate = $rule99ParamsCreate
}
else {
    $rule99ParamsCreate = @{
        SenderDomainIs = $NoiseDomains
        MoveToFolder   = "Inbox\Low Priority"
        MarkAsRead     = $true
    }
    $rule99ParamsUpdate = $rule99ParamsCreate
}
Ensure-Rule -Name "99 - Noise Filter" -CreateParams $rule99ParamsCreate -UpdateParams $rule99ParamsUpdate -Priority 99

Write-Host "`nAll done. New folders and rules are in place." -ForegroundColor Green
