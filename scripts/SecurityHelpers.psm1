<#
.SYNOPSIS
    Security helper functions for Outlook Rules Manager.

.DESCRIPTION
    Provides:
    - Path traversal protection
    - Email address validation
    - Configuration schema validation
    - Audit logging
#>

# ============================================
# PATH SECURITY
# ============================================

function Resolve-SafePath {
    <#
    .SYNOPSIS
        Resolves a path and ensures it's within the allowed directory.
    .DESCRIPTION
        Prevents path traversal attacks by validating paths are within
        the script root directory or a specified base directory.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [string]$BaseDirectory = $PSScriptRoot
    )

    try {
        # Resolve to absolute path
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

        # Normalize paths for comparison
        $normalizedBase = [System.IO.Path]::GetFullPath($BaseDirectory).TrimEnd('\', '/')
        $normalizedPath = [System.IO.Path]::GetFullPath($resolvedPath).TrimEnd('\', '/')

        # Check if path is within base directory
        if (-not $normalizedPath.StartsWith($normalizedBase, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Security Error: Path '$Path' is outside the allowed directory '$BaseDirectory'"
        }

        return $normalizedPath
    }
    catch {
        throw "Security Error: Invalid path '$Path' - $($_.Exception.Message)"
    }
}

# ============================================
# INPUT VALIDATION
# ============================================

function Test-ValidEmail {
    <#
    .SYNOPSIS
        Validates an email address format.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email
    )

    # RFC 5322 simplified pattern
    $pattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return $Email -match $pattern
}

function Test-ValidDomain {
    <#
    .SYNOPSIS
        Validates a domain pattern (supports wildcards).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    # Allow wildcard patterns like "*.example.com" or "newsletter.*"
    $pattern = '^(\*\.)?[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?)*(\.\*)?$'
    return $Domain -match $pattern
}

function Test-ConfigEmails {
    <#
    .SYNOPSIS
        Validates all email addresses in a configuration object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    $errors = @()

    # Check senderLists for email addresses
    if ($Config.senderLists) {
        foreach ($listName in $Config.senderLists.PSObject.Properties.Name) {
            $list = $Config.senderLists.$listName

            # Check addresses array
            if ($list.addresses) {
                foreach ($email in $list.addresses) {
                    if (-not (Test-ValidEmail -Email $email)) {
                        $errors += "Invalid email in senderLists.$listName.addresses: $email"
                    }
                }
            }

            # Check domains array
            if ($list.domains) {
                foreach ($domain in $list.domains) {
                    if (-not (Test-ValidDomain -Domain $domain)) {
                        $errors += "Invalid domain in senderLists.$listName.domains: $domain"
                    }
                }
            }
        }
    }

    if ($errors.Count -gt 0) {
        return [PSCustomObject]@{
            Valid = $false
            Errors = $errors
        }
    }

    return [PSCustomObject]@{
        Valid = $true
        Errors = @()
    }
}

# ============================================
# SCHEMA VALIDATION
# ============================================

function Test-ConfigSchema {
    <#
    .SYNOPSIS
        Validates the structure of a rules configuration object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    $errors = @()
    $warnings = @()

    # Required top-level fields
    $requiredFields = @('settings', 'folders', 'senderLists', 'keywordLists', 'rules')
    foreach ($field in $requiredFields) {
        if (-not $Config.$field) {
            $errors += "Missing required field: $field"
        }
    }

    # Validate settings
    if ($Config.settings) {
        if ($Config.settings.noiseAction -and $Config.settings.noiseAction -notin @('Archive', 'Delete')) {
            $errors += "settings.noiseAction must be 'Archive' or 'Delete'"
        }
    }

    # Validate folders
    if ($Config.folders) {
        $folderNames = @()
        foreach ($folder in $Config.folders) {
            if (-not $folder.name) {
                $errors += "Folder missing required 'name' field"
            } else {
                if ($folder.name -in $folderNames) {
                    $errors += "Duplicate folder name: $($folder.name)"
                }
                $folderNames += $folder.name
            }
        }
    }

    # Validate rules
    if ($Config.rules) {
        $ruleIds = @()
        $rulePriorities = @()

        foreach ($rule in $Config.rules) {
            # Check required fields
            if (-not $rule.id) {
                $errors += "Rule missing required 'id' field"
            } elseif ($rule.id -in $ruleIds) {
                $errors += "Duplicate rule id: $($rule.id)"
            } else {
                $ruleIds += $rule.id
            }

            if (-not $rule.name) {
                $errors += "Rule '$($rule.id)' missing required 'name' field"
            }

            if (-not $rule.priority) {
                $errors += "Rule '$($rule.id)' missing required 'priority' field"
            } elseif ($rule.priority -in $rulePriorities) {
                $warnings += "Duplicate rule priority $($rule.priority) in rule '$($rule.id)'"
            } else {
                $rulePriorities += $rule.priority
            }

            # Check conditions
            if (-not $rule.conditions -or ($rule.conditions.PSObject.Properties.Count -eq 0)) {
                $warnings += "Rule '$($rule.id)' has no conditions (will match all emails)"
            }

            # Check actions
            if (-not $rule.actions -or ($rule.actions.PSObject.Properties.Count -eq 0)) {
                $errors += "Rule '$($rule.id)' has no actions"
            }

            # Validate folder references
            if ($rule.actions.moveToFolder) {
                $folderName = $rule.actions.moveToFolder -replace '^Inbox\\', ''
                if ($Config.folders -and $folderName -notin $Config.folders.name) {
                    $warnings += "Rule '$($rule.id)' references non-existent folder: $folderName"
                }
            }

            # Security warning for forwarding rules
            if ($rule.actions.forwardTo -or $rule.actions.redirectTo) {
                $warnings += "SECURITY: Rule '$($rule.id)' forwards/redirects email - verify this is intentional"
            }
        }
    }

    return [PSCustomObject]@{
        Valid = ($errors.Count -eq 0)
        Errors = $errors
        Warnings = $warnings
    }
}

# ============================================
# AUDIT LOGGING
# ============================================

function Write-AuditLog {
    <#
    .SYNOPSIS
        Writes an audit log entry for security-relevant operations.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [Parameter(Mandatory = $false)]
        [string]$RuleName,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Success', 'Failure', 'Warning')]
        [string]$Result,

        [Parameter(Mandatory = $false)]
        [string]$Details,

        [Parameter(Mandatory = $false)]
        [string]$LogDirectory
    )

    if (-not $LogDirectory) {
        $LogDirectory = Join-Path $PSScriptRoot "..\logs"
    }

    # Create logs directory if it doesn't exist
    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    $logEntry = [PSCustomObject]@{
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        Operation = $Operation
        RuleName = $RuleName
        Result = $Result
        Details = $Details
        User = $env:USERNAME
        Computer = $env:COMPUTERNAME
        ProcessId = $PID
    }

    $logFile = Join-Path $LogDirectory "audit-$(Get-Date -Format 'yyyy-MM-dd').json"

    # Append as JSON lines format
    $logEntry | ConvertTo-Json -Compress | Add-Content -Path $logFile

    return $logEntry
}

function Get-AuditLogs {
    <#
    .SYNOPSIS
        Retrieves audit log entries.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [datetime]$StartDate = (Get-Date).AddDays(-7),

        [Parameter(Mandatory = $false)]
        [datetime]$EndDate = (Get-Date),

        [Parameter(Mandatory = $false)]
        [string]$Operation,

        [Parameter(Mandatory = $false)]
        [string]$LogDirectory
    )

    if (-not $LogDirectory) {
        $LogDirectory = Join-Path $PSScriptRoot "..\logs"
    }

    if (-not (Test-Path $LogDirectory)) {
        return @()
    }

    $logs = @()
    $logFiles = Get-ChildItem -Path $LogDirectory -Filter "audit-*.json"

    foreach ($file in $logFiles) {
        # Parse date from filename
        if ($file.Name -match 'audit-(\d{4}-\d{2}-\d{2})\.json') {
            $fileDate = [datetime]::ParseExact($matches[1], 'yyyy-MM-dd', $null)
            if ($fileDate -ge $StartDate.Date -and $fileDate -le $EndDate.Date) {
                $content = Get-Content $file.FullName
                foreach ($line in $content) {
                    if ($line.Trim()) {
                        $entry = $line | ConvertFrom-Json
                        if (-not $Operation -or $entry.Operation -eq $Operation) {
                            $logs += $entry
                        }
                    }
                }
            }
        }
    }

    return $logs | Sort-Object Timestamp
}

# ============================================
# SENSITIVE DATA DETECTION
# ============================================

function Test-SensitiveData {
    <#
    .SYNOPSIS
        Checks content for sensitive data patterns.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Content,

        [Parameter(Mandatory = $false)]
        [string[]]$BlockedDomains = @(
            'microsoft.com',
            'outlook.com',
            'hotmail.com',
            'gmail.com',
            'live.com'
        )
    )

    $findings = @()

    # Check for blocked email domains
    foreach ($domain in $BlockedDomains) {
        $pattern = "@$([regex]::Escape($domain))"
        if ($Content -match $pattern) {
            $findings += [PSCustomObject]@{
                Type = 'BlockedEmailDomain'
                Pattern = "@$domain"
                Severity = 'High'
            }
        }
    }

    # Check for Azure GUIDs (excluding placeholders and known Microsoft GUIDs)
    $guidPattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
    $allowedGuids = @(
        '00000000-0000-0000-0000-000000000000',  # Placeholder
        '00000003-0000-0000-c000-000000000000',  # Microsoft Graph
        'e383f46e-2787-4529-855e-0e479a3ffac0',  # Mail.ReadWrite
        '570282fd-fa5c-430d-a7fd-fc8dc98a9dca',  # Mail.Read
        'e1fe6dd8-ba31-4d61-89e7-88639da4683d'   # User.Read
    )

    $matches = [regex]::Matches($Content, $guidPattern)
    foreach ($match in $matches) {
        if ($match.Value -notin $allowedGuids) {
            $findings += [PSCustomObject]@{
                Type = 'AzureGUID'
                Pattern = $match.Value.Substring(0, 8) + '...'
                Severity = 'Medium'
            }
        }
    }

    # Check for potential secrets
    $secretPatterns = @{
        'ClientSecret' = 'client[_-]?secret\s*[=:]\s*["\x27][A-Za-z0-9~._-]{20,}["\x27]'
        'ApiKey' = 'api[_-]?key\s*[=:]\s*["\x27][A-Za-z0-9_-]{20,}["\x27]'
        'Password' = 'password\s*[=:]\s*["\x27][^\x27"]{8,}["\x27]'
        'BearerToken' = 'bearer\s+[A-Za-z0-9_\-\.=]{50,}'
    }

    foreach ($patternName in $secretPatterns.Keys) {
        if ($Content -match $secretPatterns[$patternName]) {
            $findings += [PSCustomObject]@{
                Type = $patternName
                Pattern = '[REDACTED]'
                Severity = 'Critical'
            }
        }
    }

    return [PSCustomObject]@{
        HasSensitiveData = ($findings.Count -gt 0)
        Findings = $findings
    }
}

# ============================================
# EXPORTS
# ============================================

Export-ModuleMember -Function @(
    'Resolve-SafePath',
    'Test-ValidEmail',
    'Test-ValidDomain',
    'Test-ConfigEmails',
    'Test-ConfigSchema',
    'Write-AuditLog',
    'Get-AuditLogs',
    'Test-SensitiveData'
)
