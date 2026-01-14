<#
.SYNOPSIS
    Pre-commit security check for Outlook Rules Manager
    Run this BEFORE committing to catch credentials, secrets, and PII

.DESCRIPTION
    Scans staged and modified files for:
    - Email addresses (real domains)
    - Azure GUIDs (Tenant/Client IDs)
    - Secrets and credentials
    - Files that should not be committed

.EXAMPLE
    .\scripts\Check-BeforeCommit.ps1

.EXAMPLE
    # Run with verbose output
    .\scripts\Check-BeforeCommit.ps1 -Verbose
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Continue"
$script:IssuesFound = 0

# Colors for output
function Write-Success { param($Message) Write-Host "✅ $Message" -ForegroundColor Green }
function Write-Warn { param($Message) Write-Host "⚠️ $Message" -ForegroundColor Yellow }
function Write-Fail { param($Message) Write-Host "❌ $Message" -ForegroundColor Red; $script:IssuesFound++ }
function Write-Info { param($Message) Write-Host "🔍 $Message" -ForegroundColor Cyan }

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  🔒 PRE-COMMIT SECURITY CHECK - Outlook Rules Manager" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# ============================================
# CHECK 1: Blocked Files
# ============================================
Write-Info "Checking for blocked files..."

$blockedFiles = @(
    ".env",
    "app-config.json",
    "rules-config.json",
    "demo.env"
)

$stagedFiles = git diff --cached --name-only 2>$null

foreach ($blocked in $blockedFiles) {
    if ($stagedFiles -contains $blocked) {
        Write-Fail "BLOCKED FILE STAGED: $blocked"
        Write-Host "   Run: git reset HEAD $blocked" -ForegroundColor Gray
    }
}

# Check if blocked files exist in repo
foreach ($blocked in $blockedFiles) {
    $tracked = git ls-files $blocked 2>$null
    if ($tracked) {
        Write-Fail "BLOCKED FILE IN REPO: $blocked"
        Write-Host "   Run: git rm --cached $blocked" -ForegroundColor Gray
    }
}

if ($script:IssuesFound -eq 0) {
    Write-Success "No blocked files detected"
}

# ============================================
# CHECK 2: Email Addresses
# ============================================
Write-Host ""
Write-Info "Checking for email addresses..."

$blockedDomains = @(
    "microsoft.com",
    "limitlessdata.ai",
    "housegarofalo.com",
    "outlook.com",
    "hotmail.com",
    "gmail.com",
    "live.com"
)

$allowedDomains = @(
    "example.com",
    "company.com",
    "domain.com",
    "contoso.com",
    "fabrikam.com"
)

$emailIssues = 0
$filesToCheck = Get-ChildItem -Path . -Include *.ps1,*.json -Recurse | Where-Object { $_.Name -notmatch "example" }

foreach ($file in $filesToCheck) {
    $content = Get-Content $file.FullName -Raw -ErrorAction SilentlyContinue
    if (-not $content) { continue }

    foreach ($domain in $blockedDomains) {
        $pattern = "@$domain"
        if ($content -match [regex]::Escape($pattern)) {
            Write-Fail "EMAIL FOUND in $($file.Name): @$domain"
            $emailIssues++
        }
    }
}

if ($emailIssues -eq 0) {
    Write-Success "No blocked email domains found"
}

# ============================================
# CHECK 3: Azure GUIDs
# ============================================
Write-Host ""
Write-Info "Checking for Azure GUIDs..."

$guidPattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
$placeholderGuid = '00000000-0000-0000-0000-000000000000'
$guidIssues = 0

foreach ($file in $filesToCheck) {
    $content = Get-Content $file.FullName -Raw -ErrorAction SilentlyContinue
    if (-not $content) { continue }

    $matches = [regex]::Matches($content, $guidPattern)
    foreach ($match in $matches) {
        if ($match.Value -ne $placeholderGuid) {
            # Check if it's in a variable assignment (likely intentional)
            $line = ($content.Substring(0, $match.Index) -split "`n")[-1]
            if ($line -match '\$\w+\s*=' -or $line -match 'ClientId|TenantId') {
                Write-Warn "GUID in $($file.Name): $($match.Value.Substring(0,8))... (may be config)"
            } else {
                Write-Fail "GUID FOUND in $($file.Name): $($match.Value.Substring(0,8))..."
                $guidIssues++
            }
        }
    }
}

if ($guidIssues -eq 0) {
    Write-Success "No exposed Azure GUIDs found"
}

# ============================================
# CHECK 4: Secrets Patterns
# ============================================
Write-Host ""
Write-Info "Checking for secret patterns..."

$secretPatterns = @{
    "Client Secret" = 'client.*secret\s*=\s*["''][A-Za-z0-9~._-]{20,}["'']'
    "API Key" = 'api[-_]?key\s*=\s*["''][A-Za-z0-9_-]{20,}["'']'
    "Password" = 'password\s*=\s*["''][^"'']+["'']'
    "Bearer Token" = 'bearer\s+[A-Za-z0-9_\-\.=]{20,}'
}

$secretIssues = 0

foreach ($file in $filesToCheck) {
    $content = Get-Content $file.FullName -Raw -ErrorAction SilentlyContinue
    if (-not $content) { continue }

    foreach ($patternName in $secretPatterns.Keys) {
        if ($content -match $secretPatterns[$patternName]) {
            Write-Fail "$patternName FOUND in $($file.Name)"
            $secretIssues++
        }
    }
}

if ($secretIssues -eq 0) {
    Write-Success "No secret patterns found"
}

# ============================================
# CHECK 5: Sensitive Data in Staged Changes
# ============================================
Write-Host ""
Write-Info "Checking staged changes..."

$stagedDiff = git diff --cached 2>$null

if ($stagedDiff) {
    $diffIssues = 0

    foreach ($domain in $blockedDomains) {
        if ($stagedDiff -match [regex]::Escape("@$domain")) {
            Write-Fail "STAGED CHANGE contains @$domain"
            $diffIssues++
        }
    }

    # Check for GUIDs in staged changes
    if ($stagedDiff -match $guidPattern) {
        $guidMatches = [regex]::Matches($stagedDiff, $guidPattern) | Where-Object { $_.Value -ne $placeholderGuid }
        if ($guidMatches.Count -gt 0) {
            Write-Warn "STAGED CHANGE contains $($guidMatches.Count) GUID(s) - verify these are not real IDs"
        }
    }

    if ($diffIssues -eq 0) {
        Write-Success "Staged changes look clean"
    }
} else {
    Write-Host "   No staged changes to check" -ForegroundColor Gray
}

# ============================================
# SUMMARY
# ============================================
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan

if ($script:IssuesFound -gt 0) {
    Write-Host ""
    Write-Fail "FOUND $script:IssuesFound ISSUE(S) - Please fix before committing!"
    Write-Host ""
    Write-Host "Tips:" -ForegroundColor Yellow
    Write-Host "  • Use .env files for credentials (gitignored)" -ForegroundColor Gray
    Write-Host "  • Use example.com for sample emails" -ForegroundColor Gray
    Write-Host "  • Use 00000000-0000-0000-0000-000000000000 for placeholder GUIDs" -ForegroundColor Gray
    Write-Host ""
    exit 1
} else {
    Write-Host ""
    Write-Success "ALL CHECKS PASSED - Safe to commit!"
    Write-Host ""
    exit 0
}
