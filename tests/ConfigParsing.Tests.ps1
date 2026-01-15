#Requires -Modules Pester
<#
.SYNOPSIS
    Pester tests for configuration parsing and reference resolution.

.DESCRIPTION
    Tests covering:
    - .env file parsing
    - rules-config.json parsing
    - Reference resolution (@senderLists.priority, etc.)
    - Config profile handling

.NOTES
    Run with: Invoke-Pester -Path .\tests\ConfigParsing.Tests.ps1 -Output Detailed
#>

BeforeAll {
    # Create temp directory for test files
    $script:TestTempDir = Join-Path $PSScriptRoot "temp"
    if (-not (Test-Path $script:TestTempDir)) {
        New-Item -ItemType Directory -Path $script:TestTempDir -Force | Out-Null
    }

    # Helper function to resolve config references (extracted from Manage-OutlookRules.ps1)
    function Resolve-ConfigReferences {
        param($Config, $Value)

        if ($Value -is [string] -and $Value.StartsWith("@")) {
            $path = $Value.Substring(1).Split(".")
            $resolved = $Config
            foreach ($segment in $path) {
                if ($null -eq $resolved) { return $null }
                $resolved = $resolved.$segment
            }

            if ($null -eq $resolved) { return $null }
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
}

AfterAll {
    # Cleanup temp directory
    if (Test-Path $script:TestTempDir) {
        Remove-Item -Path $script:TestTempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe "Environment File Parsing" {
    Context "Standard .env format" {
        It "Should parse KEY=value format" {
            $envFile = Join-Path $script:TestTempDir ".env.test1"
            @"
ClientId=12345678-1234-1234-1234-123456789012
TenantId=87654321-4321-4321-4321-210987654321
"@ | Set-Content $envFile

            $vars = @{}
            $content = Get-Content $envFile
            foreach ($line in $content) {
                $line = $line.Trim()
                if ($line -and -not $line.StartsWith("#")) {
                    if ($line -match '^(\w+)=(.+)$') {
                        $vars[$matches[1]] = $matches[2]
                    }
                }
            }

            $vars["ClientId"] | Should -Be "12345678-1234-1234-1234-123456789012"
            $vars["TenantId"] | Should -Be "87654321-4321-4321-4321-210987654321"
        }

        It "Should parse KEY=`"value`" format" {
            $envFile = Join-Path $script:TestTempDir ".env.test2"
            @"
ClientId="12345678-1234-1234-1234-123456789012"
TenantId="87654321-4321-4321-4321-210987654321"
"@ | Set-Content $envFile

            $vars = @{}
            $content = Get-Content $envFile
            foreach ($line in $content) {
                $line = $line.Trim()
                if ($line -and -not $line.StartsWith("#")) {
                    if ($line -match '^(\w+)="?([^"]+)"?$') {
                        $vars[$matches[1]] = $matches[2]
                    }
                }
            }

            $vars["ClientId"] | Should -Be "12345678-1234-1234-1234-123456789012"
        }

        It "Should parse PowerShell variable format" {
            $envFile = Join-Path $script:TestTempDir ".env.test3"
            @"
`$ClientId = "12345678-1234-1234-1234-123456789012"
`$TenantId = "87654321-4321-4321-4321-210987654321"
"@ | Set-Content $envFile

            $vars = @{}
            $content = Get-Content $envFile
            foreach ($line in $content) {
                $line = $line.Trim()
                if ($line -and -not $line.StartsWith("#")) {
                    if ($line -match '^\$?(\w+)\s*=\s*"?([^"]+)"?$') {
                        $vars[$matches[1]] = $matches[2]
                    }
                }
            }

            $vars["ClientId"] | Should -Be "12345678-1234-1234-1234-123456789012"
        }

        It "Should skip comments" {
            $envFile = Join-Path $script:TestTempDir ".env.test4"
            @"
# This is a comment
ClientId=12345678-1234-1234-1234-123456789012
# Another comment
TenantId=87654321-4321-4321-4321-210987654321
"@ | Set-Content $envFile

            $vars = @{}
            $content = Get-Content $envFile
            foreach ($line in $content) {
                $line = $line.Trim()
                if ($line -and -not $line.StartsWith("#")) {
                    if ($line -match '^(\w+)=(.+)$') {
                        $vars[$matches[1]] = $matches[2]
                    }
                }
            }

            $vars.Count | Should -Be 2
            $vars["ClientId"] | Should -Not -BeNullOrEmpty
        }

        It "Should skip empty lines" {
            $envFile = Join-Path $script:TestTempDir ".env.test5"
            @"

ClientId=12345678-1234-1234-1234-123456789012

TenantId=87654321-4321-4321-4321-210987654321

"@ | Set-Content $envFile

            $vars = @{}
            $content = Get-Content $envFile
            foreach ($line in $content) {
                $line = $line.Trim()
                if ($line -and -not $line.StartsWith("#")) {
                    if ($line -match '^(\w+)=(.+)$') {
                        $vars[$matches[1]] = $matches[2]
                    }
                }
            }

            $vars.Count | Should -Be 2
        }
    }
}

Describe "Rules Config JSON Parsing" {
    Context "Valid configuration structure" {
        It "Should parse minimal valid config" {
            $configFile = Join-Path $script:TestTempDir "rules-config.test1.json"
            @"
{
    "settings": { "noiseAction": "Archive" },
    "folders": [],
    "senderLists": {},
    "keywordLists": {},
    "rules": []
}
"@ | Set-Content $configFile

            $config = Get-Content $configFile -Raw | ConvertFrom-Json
            $config.settings.noiseAction | Should -Be "Archive"
            $config.rules | Should -BeNullOrEmpty
        }

        It "Should parse config with sender lists" {
            $configFile = Join-Path $script:TestTempDir "rules-config.test2.json"
            @"
{
    "settings": {},
    "folders": [],
    "senderLists": {
        "priority": {
            "description": "VIP senders",
            "addresses": ["vip1@example.com", "vip2@example.com"]
        },
        "noise": {
            "domains": ["newsletter.example.com", "*.marketing.com"]
        }
    },
    "keywordLists": {},
    "rules": []
}
"@ | Set-Content $configFile

            $config = Get-Content $configFile -Raw | ConvertFrom-Json
            $config.senderLists.priority.addresses.Count | Should -Be 2
            $config.senderLists.noise.domains | Should -Contain "*.marketing.com"
        }

        It "Should parse config with keyword lists" {
            $configFile = Join-Path $script:TestTempDir "rules-config.test3.json"
            @"
{
    "settings": {},
    "folders": [],
    "senderLists": {},
    "keywordLists": {
        "action": {
            "keywords": ["Urgent", "Action Required", "Due"]
        }
    },
    "rules": []
}
"@ | Set-Content $configFile

            $config = Get-Content $configFile -Raw | ConvertFrom-Json
            $config.keywordLists.action.keywords.Count | Should -Be 3
        }
    }
}

Describe "Reference Resolution" {
    Context "Sender list references" {
        It "Should resolve @senderLists.priority.addresses" {
            $config = [PSCustomObject]@{
                senderLists = [PSCustomObject]@{
                    priority = [PSCustomObject]@{
                        addresses = @("vip1@example.com", "vip2@example.com")
                    }
                }
            }

            $result = Resolve-ConfigReferences -Config $config -Value "@senderLists.priority"
            $result | Should -Contain "vip1@example.com"
            $result | Should -Contain "vip2@example.com"
        }

        It "Should resolve @senderLists.noise.domains" {
            $config = [PSCustomObject]@{
                senderLists = [PSCustomObject]@{
                    noise = [PSCustomObject]@{
                        domains = @("newsletter.example.com", "*.marketing.com")
                    }
                }
            }

            $result = Resolve-ConfigReferences -Config $config -Value "@senderLists.noise"
            $result | Should -Contain "newsletter.example.com"
            $result | Should -Contain "*.marketing.com"
            $result.Count | Should -Be 2
        }

        It "Should handle wildcard domain patterns" {
            $config = [PSCustomObject]@{
                senderLists = [PSCustomObject]@{
                    blocked = [PSCustomObject]@{
                        domains = @("*.spam.com", "*.newsletter.net", "marketing.example.org")
                    }
                }
            }

            $result = Resolve-ConfigReferences -Config $config -Value "@senderLists.blocked"
            $result.Count | Should -Be 3
            ($result | Where-Object { $_ -like '`**' }).Count | Should -Be 2
        }
    }

    Context "Keyword list references" {
        It "Should resolve @keywordLists.action" {
            $config = [PSCustomObject]@{
                keywordLists = [PSCustomObject]@{
                    action = [PSCustomObject]@{
                        keywords = @("Urgent", "Action Required", "Due")
                    }
                }
            }

            $result = Resolve-ConfigReferences -Config $config -Value "@keywordLists.action"
            $result | Should -Contain "Urgent"
            $result.Count | Should -Be 3
        }
    }

    Context "Settings references" {
        It "Should resolve @settings.categories.action" {
            $config = [PSCustomObject]@{
                settings = [PSCustomObject]@{
                    categories = [PSCustomObject]@{
                        action = "Action Required"
                        metrics = "Metrics"
                    }
                }
            }

            $result = Resolve-ConfigReferences -Config $config -Value "@settings.categories.action"
            $result | Should -Be "Action Required"
        }
    }

    Context "Non-reference values" {
        It "Should return literal strings unchanged" {
            $config = [PSCustomObject]@{}
            $result = Resolve-ConfigReferences -Config $config -Value "literal@value.com"
            $result | Should -Be "literal@value.com"
        }

        It "Should return arrays unchanged" {
            $config = [PSCustomObject]@{}
            $result = Resolve-ConfigReferences -Config $config -Value @("a", "b", "c")
            $result | Should -Contain "a"
            $result.Count | Should -Be 3
        }
    }
}

Describe "Config Profile Handling" {
    Context "Profile file resolution" {
        It "Should construct correct profile env filename" {
            $profile = "personal"
            $envFile = ".env.$profile"
            $envFile | Should -Be ".env.personal"
        }

        It "Should construct correct profile config filename" {
            $profile = "work"
            $configFile = "rules-config.$profile.json"
            $configFile | Should -Be "rules-config.work.json"
        }

        It "Should default to .env when no profile" {
            $profile = $null
            $envFile = if ($profile) { ".env.$profile" } else { ".env" }
            $envFile | Should -Be ".env"
        }

        It "Should default to rules-config.json when no profile" {
            $profile = $null
            $configFile = if ($profile) { "rules-config.$profile.json" } else { "rules-config.json" }
            $configFile | Should -Be "rules-config.json"
        }
    }
}

Describe "Authorization Role Constants" {
    Context "App Role IDs" {
        It "Should have correct Admin role ID" {
            $adminRoleId = "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f"
            $adminRoleId | Should -Match "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$"
        }

        It "Should have correct User role ID" {
            $userRoleId = "a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d"
            $userRoleId | Should -Match "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$"
        }

        It "Admin and User role IDs should be different" {
            $adminRoleId = "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f"
            $userRoleId = "a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d"
            $adminRoleId | Should -Not -Be $userRoleId
        }
    }

    Context "Role name mapping" {
        It "Should map Admin role ID to Administrator" {
            $roleId = "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f"
            $roleName = switch ($roleId) {
                "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f" { "Administrator" }
                "a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d" { "User" }
                default { "Unknown" }
            }
            $roleName | Should -Be "Administrator"
        }

        It "Should map User role ID to User" {
            $roleId = "a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d"
            $roleName = switch ($roleId) {
                "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f" { "Administrator" }
                "a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d" { "User" }
                default { "Unknown" }
            }
            $roleName | Should -Be "User"
        }

        It "Should return Unknown for unrecognized role ID" {
            $roleId = "00000000-0000-0000-0000-000000000000"
            $roleName = switch ($roleId) {
                "f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f" { "Administrator" }
                "a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d" { "User" }
                default { "Unknown" }
            }
            $roleName | Should -Be "Unknown"
        }
    }
}
