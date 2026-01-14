#Requires -Modules Pester
<#
.SYNOPSIS
    Pester tests for SecurityHelpers module.

.DESCRIPTION
    Comprehensive tests covering:
    - Path traversal protection
    - Email/domain validation
    - Configuration schema validation
    - Audit logging
    - HTML sanitization
    - Sensitive data detection

.NOTES
    Run with: Invoke-Pester -Path .\tests\SecurityHelpers.Tests.ps1 -Output Detailed
#>

BeforeAll {
    # Import the module under test (now in src/modules/)
    $modulePath = Join-Path $PSScriptRoot "..\src\modules\SecurityHelpers.psm1"
    Import-Module $modulePath -Force

    # Create temp directory for tests
    $script:TestTempDir = Join-Path $PSScriptRoot "temp"
    if (-not (Test-Path $script:TestTempDir)) {
        New-Item -ItemType Directory -Path $script:TestTempDir -Force | Out-Null
    }
}

AfterAll {
    # Cleanup temp directory
    if (Test-Path $script:TestTempDir) {
        Remove-Item -Path $script:TestTempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe "Resolve-SafePath" {
    Context "Valid paths within base directory" {
        It "Should resolve a simple relative path" {
            $basePath = $PSScriptRoot
            $result = Resolve-SafePath -Path ".\test.json" -BaseDirectory $basePath
            $result | Should -BeLike "*test.json"
        }

        It "Should resolve an absolute path within base" {
            $basePath = $PSScriptRoot
            $testPath = Join-Path $basePath "subfolder\file.txt"
            $result = Resolve-SafePath -Path $testPath -BaseDirectory $basePath
            $result | Should -BeLike "*subfolder*file.txt"
        }
    }

    Context "Path traversal attacks" {
        It "Should block simple parent directory traversal" {
            $basePath = Join-Path $PSScriptRoot "temp"
            New-Item -ItemType Directory -Path $basePath -Force | Out-Null

            { Resolve-SafePath -Path "..\..\..\etc\passwd" -BaseDirectory $basePath } |
                Should -Throw "*outside the allowed directory*"
        }

        It "Should block encoded path traversal" {
            $basePath = Join-Path $PSScriptRoot "temp"
            { Resolve-SafePath -Path "..%2F..%2Fetc%2Fpasswd" -BaseDirectory $basePath } |
                Should -Throw
        }

        It "Should block absolute path outside base" {
            $basePath = Join-Path $PSScriptRoot "temp"
            { Resolve-SafePath -Path "C:\Windows\System32\config" -BaseDirectory $basePath } |
                Should -Throw "*outside the allowed directory*"
        }

        It "Should block double-dot sequences with different separators" {
            $basePath = Join-Path $PSScriptRoot "temp"
            { Resolve-SafePath -Path "..\..\..\..\Windows" -BaseDirectory $basePath } |
                Should -Throw
        }
    }
}

Describe "Test-ValidEmail" {
    Context "Valid email addresses" {
        It "Should accept standard email format" {
            Test-ValidEmail -Email "user@example.com" | Should -BeTrue
        }

        It "Should accept email with subdomain" {
            Test-ValidEmail -Email "user@mail.example.com" | Should -BeTrue
        }

        It "Should accept email with plus addressing" {
            Test-ValidEmail -Email "user+tag@example.com" | Should -BeTrue
        }

        It "Should accept email with dots in local part" {
            Test-ValidEmail -Email "first.last@example.com" | Should -BeTrue
        }

        It "Should accept email with numbers" {
            Test-ValidEmail -Email "user123@example456.com" | Should -BeTrue
        }

        It "Should accept email with long TLD" {
            Test-ValidEmail -Email "user@example.museum" | Should -BeTrue
        }
    }

    Context "Invalid email addresses" {
        It "Should reject email without @" {
            Test-ValidEmail -Email "userexample.com" | Should -BeFalse
        }

        It "Should reject email without domain" {
            Test-ValidEmail -Email "user@" | Should -BeFalse
        }

        It "Should reject email without local part" {
            Test-ValidEmail -Email "@example.com" | Should -BeFalse
        }

        It "Should reject email without TLD" {
            Test-ValidEmail -Email "user@example" | Should -BeFalse
        }

        It "Should reject email with spaces" {
            Test-ValidEmail -Email "user @example.com" | Should -BeFalse
        }

        It "Should reject email with multiple @" {
            Test-ValidEmail -Email "user@@example.com" | Should -BeFalse
        }

        It "Should reject empty string" {
            Test-ValidEmail -Email "" | Should -BeFalse
        }
    }
}

Describe "Test-ValidDomain" {
    Context "Valid domain patterns" {
        It "Should accept simple domain" {
            Test-ValidDomain -Domain "example.com" | Should -BeTrue
        }

        It "Should accept subdomain" {
            Test-ValidDomain -Domain "mail.example.com" | Should -BeTrue
        }

        It "Should accept wildcard prefix" {
            Test-ValidDomain -Domain "*.example.com" | Should -BeTrue
        }

        It "Should accept wildcard suffix" {
            Test-ValidDomain -Domain "newsletter.*" | Should -BeTrue
        }

        It "Should accept domain with hyphens" {
            Test-ValidDomain -Domain "my-example.co.uk" | Should -BeTrue
        }

        It "Should accept domain with numbers" {
            Test-ValidDomain -Domain "example123.com" | Should -BeTrue
        }
    }

    Context "Invalid domain patterns" {
        It "Should reject domain starting with hyphen" {
            Test-ValidDomain -Domain "-example.com" | Should -BeFalse
        }

        It "Should reject domain with invalid characters" {
            Test-ValidDomain -Domain "exam_ple.com" | Should -BeFalse
        }

        It "Should reject double dots" {
            Test-ValidDomain -Domain "example..com" | Should -BeFalse
        }

        It "Should reject empty string" {
            Test-ValidDomain -Domain "" | Should -BeFalse
        }
    }
}

Describe "Test-ConfigSchema" {
    Context "Valid configuration" {
        It "Should validate minimal valid config" {
            $config = [PSCustomObject]@{
                settings = @{ noiseAction = "Archive" }
                folders = @(@{ name = "Test" })
                senderLists = @{}
                keywordLists = @{}
                rules = @(@{
                    id = "rule-1"
                    name = "Test Rule"
                    priority = 1
                    conditions = @{ from = "test@example.com" }
                    actions = @{ markAsRead = $true }
                })
            }

            $result = Test-ConfigSchema -Config $config
            $result.Valid | Should -BeTrue
            $result.Errors.Count | Should -Be 0
        }
    }

    Context "Missing required fields" {
        It "Should report missing settings" {
            $config = [PSCustomObject]@{
                folders = @()
                senderLists = @{}
                keywordLists = @{}
                rules = @()
            }

            $result = Test-ConfigSchema -Config $config
            $result.Errors | Should -Contain "Missing required field: settings"
        }

        It "Should report missing rules" {
            $config = [PSCustomObject]@{
                settings = @{}
                folders = @()
                senderLists = @{}
                keywordLists = @{}
            }

            $result = Test-ConfigSchema -Config $config
            $result.Errors | Should -Contain "Missing required field: rules"
        }
    }

    Context "Invalid settings" {
        It "Should reject invalid noiseAction value" {
            $config = [PSCustomObject]@{
                settings = @{ noiseAction = "Invalid" }
                folders = @()
                senderLists = @{}
                keywordLists = @{}
                rules = @()
            }

            $result = Test-ConfigSchema -Config $config
            $result.Errors | Should -Contain "settings.noiseAction must be 'Archive' or 'Delete'"
        }
    }

    Context "Rule validation" {
        It "Should detect duplicate rule IDs" {
            $config = [PSCustomObject]@{
                settings = @{}
                folders = @()
                senderLists = @{}
                keywordLists = @{}
                rules = @(
                    @{ id = "rule-1"; name = "Rule 1"; priority = 1; conditions = @{}; actions = @{ markAsRead = $true } }
                    @{ id = "rule-1"; name = "Rule 2"; priority = 2; conditions = @{}; actions = @{ markAsRead = $true } }
                )
            }

            $result = Test-ConfigSchema -Config $config
            $result.Errors | Should -Contain "Duplicate rule id: rule-1"
        }

        It "Should warn about duplicate priorities" {
            $config = [PSCustomObject]@{
                settings = @{}
                folders = @()
                senderLists = @{}
                keywordLists = @{}
                rules = @(
                    @{ id = "rule-1"; name = "Rule 1"; priority = 1; conditions = @{ from = "a@b.com" }; actions = @{ markAsRead = $true } }
                    @{ id = "rule-2"; name = "Rule 2"; priority = 1; conditions = @{ from = "c@d.com" }; actions = @{ markAsRead = $true } }
                )
            }

            $result = Test-ConfigSchema -Config $config
            $result.Warnings | Should -Match "Duplicate rule priority"
        }

        It "Should warn about forwarding rules" {
            $config = [PSCustomObject]@{
                settings = @{}
                folders = @()
                senderLists = @{}
                keywordLists = @{}
                rules = @(@{
                    id = "rule-1"
                    name = "Forward Rule"
                    priority = 1
                    conditions = @{ from = "test@example.com" }
                    actions = @{ forwardTo = "other@example.com" }
                })
            }

            $result = Test-ConfigSchema -Config $config
            $result.Warnings | Should -Match "SECURITY.*forwards/redirects"
        }
    }

    Context "Folder validation" {
        It "Should detect duplicate folder names" {
            $config = [PSCustomObject]@{
                settings = @{}
                folders = @(
                    @{ name = "Priority" }
                    @{ name = "Priority" }
                )
                senderLists = @{}
                keywordLists = @{}
                rules = @()
            }

            $result = Test-ConfigSchema -Config $config
            $result.Errors | Should -Contain "Duplicate folder name: Priority"
        }
    }
}

Describe "Test-ConfigEmails" {
    Context "Valid email addresses in config" {
        It "Should pass with valid emails" {
            $config = [PSCustomObject]@{
                senderLists = [PSCustomObject]@{
                    priority = [PSCustomObject]@{
                        addresses = @("user@example.com", "admin@example.org")
                    }
                }
            }

            $result = Test-ConfigEmails -Config $config
            $result.Valid | Should -BeTrue
        }

        It "Should pass with valid domains" {
            $config = [PSCustomObject]@{
                senderLists = [PSCustomObject]@{
                    noise = [PSCustomObject]@{
                        domains = @("newsletter.example.com", "*.marketing.com")
                    }
                }
            }

            $result = Test-ConfigEmails -Config $config
            $result.Valid | Should -BeTrue
        }
    }

    Context "Invalid email addresses in config" {
        It "Should fail with invalid email" {
            $config = [PSCustomObject]@{
                senderLists = [PSCustomObject]@{
                    priority = [PSCustomObject]@{
                        addresses = @("invalid-email", "user@example.com")
                    }
                }
            }

            $result = Test-ConfigEmails -Config $config
            $result.Valid | Should -BeFalse
            $result.Errors | Should -Match "Invalid email"
        }

        It "Should fail with invalid domain" {
            $config = [PSCustomObject]@{
                senderLists = [PSCustomObject]@{
                    noise = [PSCustomObject]@{
                        domains = @("--invalid..com")
                    }
                }
            }

            $result = Test-ConfigEmails -Config $config
            $result.Valid | Should -BeFalse
            $result.Errors | Should -Match "Invalid domain"
        }
    }
}

Describe "ConvertTo-SafeText" {
    Context "Script injection prevention" {
        It "Should remove script tags" {
            $text = 'Hello <script>alert("XSS")</script> World'
            $result = ConvertTo-SafeText -Text $text
            $result | Should -Not -Match "<script>"
            $result | Should -Not -Match "alert"
        }

        It "Should remove event handlers" {
            $text = '<div onclick="evil()">Click me</div>'
            $result = ConvertTo-SafeText -Text $text
            $result | Should -Not -Match "onclick"
            $result | Should -Not -Match "evil"
        }

        It "Should remove javascript: URLs" {
            $text = '<a href="javascript:alert(1)">Link</a>'
            $result = ConvertTo-SafeText -Text $text
            $result | Should -Not -Match "javascript:"
        }

        It "Should remove data: URLs" {
            $text = '<img src="data:image/svg+xml,...">'
            $result = ConvertTo-SafeText -Text $text
            $result | Should -Not -Match "data:"
        }

        It "Should remove style tags" {
            $text = '<style>.evil { display: none; }</style>Content'
            $result = ConvertTo-SafeText -Text $text
            $result | Should -Not -Match "<style>"
            $result | Should -Not -Match "evil"
        }
    }

    Context "Basic formatting allowed" {
        It "Should allow basic HTML when flag set" {
            $text = '<p>Paragraph</p><b>Bold</b>'
            $result = ConvertTo-SafeText -Text $text -AllowBasicFormatting
            # Tags should be preserved
            $result | Should -Match "Paragraph"
            $result | Should -Match "Bold"
        }
    }

    Context "Edge cases" {
        It "Should handle empty string" {
            $result = ConvertTo-SafeText -Text ""
            $result | Should -Be ""
        }

        It "Should handle null-like empty string" {
            $result = ConvertTo-SafeText -Text "   "
            $result | Should -Be ""
        }

        It "Should escape remaining angle brackets" {
            $text = "3 < 5 and 10 > 7"
            $result = ConvertTo-SafeText -Text $text
            $result | Should -Match "&lt;"
            $result | Should -Match "&gt;"
        }
    }
}

Describe "Protect-SensitiveLogData" {
    Context "Email redaction" {
        It "Should redact email addresses" {
            $data = "User email is john.doe@example.com"
            $result = Protect-SensitiveLogData -Data $data
            $result | Should -Not -Match "john.doe"
            $result | Should -Match "\*\*\*@example.com"
        }

        It "Should preserve domain for context" {
            $data = "Blocked sender@malicious.com from sending"
            $result = Protect-SensitiveLogData -Data $data
            $result | Should -Match "malicious.com"
        }
    }

    Context "GUID redaction" {
        It "Should partially redact GUIDs" {
            $data = "Resource ID: 12345678-1234-1234-1234-123456789012"
            $result = Protect-SensitiveLogData -Data $data
            $result | Should -Match "1234\*\*\*\*"
            $result | Should -Not -Match "123456789012"
        }
    }

    Context "Secret redaction" {
        It "Should redact password values" {
            $data = 'password = "MySecretPassword123"'
            $result = Protect-SensitiveLogData -Data $data
            $result | Should -Match "\[REDACTED\]"
            $result | Should -Not -Match "MySecretPassword123"
        }

        It "Should redact token values" {
            $data = 'token: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9"'
            $result = Protect-SensitiveLogData -Data $data
            $result | Should -Match "\[REDACTED\]"
        }
    }

    Context "Edge cases" {
        It "Should handle empty string" {
            $result = Protect-SensitiveLogData -Data ""
            $result | Should -Be ""
        }
    }
}

Describe "Test-SensitiveData" {
    Context "Email domain detection" {
        It "Should detect blocked email domains" {
            $content = "Forward to user@microsoft.com"
            $result = Test-SensitiveData -Content $content
            $result.HasSensitiveData | Should -BeTrue
            $result.Findings.Type | Should -Contain "BlockedEmailDomain"
        }

        It "Should allow non-blocked domains" {
            $content = "Forward to user@example.com"
            $result = Test-SensitiveData -Content $content
            $result.Findings | Where-Object { $_.Type -eq "BlockedEmailDomain" } | Should -BeNullOrEmpty
        }
    }

    Context "GUID detection" {
        It "Should detect Azure GUIDs" {
            $content = "Client ID: 12345678-1234-1234-1234-123456789012"
            $result = Test-SensitiveData -Content $content
            $result.HasSensitiveData | Should -BeTrue
            $result.Findings.Type | Should -Contain "AzureGUID"
        }

        It "Should allow known Microsoft GUIDs" {
            $content = "Graph ID: 00000003-0000-0000-c000-000000000000"
            $result = Test-SensitiveData -Content $content
            $result.Findings | Where-Object { $_.Type -eq "AzureGUID" } | Should -BeNullOrEmpty
        }
    }

    Context "Secret detection" {
        It "Should detect client secrets" {
            $content = 'client_secret = "AbCdEfGhIjKlMnOpQrStUvWxYz12345678"'
            $result = Test-SensitiveData -Content $content
            $result.HasSensitiveData | Should -BeTrue
            $result.Findings.Type | Should -Contain "ClientSecret"
        }

        It "Should detect API keys" {
            # Using clearly fake test pattern (not a real key format)
            $content = 'api-key: "test_fake_key_1234567890abcdefghij"'
            $result = Test-SensitiveData -Content $content
            $result.HasSensitiveData | Should -BeTrue
            $result.Findings.Type | Should -Contain "ApiKey"
        }

        It "Should detect bearer tokens" {
            $longToken = "A" * 60
            $content = "Authorization: Bearer $longToken"
            $result = Test-SensitiveData -Content $content
            $result.HasSensitiveData | Should -BeTrue
            $result.Findings.Type | Should -Contain "BearerToken"
        }
    }
}

Describe "Write-AuditLog and Get-AuditLogs" {
    BeforeAll {
        $script:TestLogDir = Join-Path $script:TestTempDir "logs"
        if (Test-Path $script:TestLogDir) {
            Remove-Item -Path $script:TestLogDir -Recurse -Force
        }
    }

    AfterAll {
        if (Test-Path $script:TestLogDir) {
            Remove-Item -Path $script:TestLogDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Context "Writing audit logs" {
        It "Should create log directory if missing" {
            $logDir = Join-Path $script:TestTempDir "newlogs"
            Write-AuditLog -Operation "Test" -Result "Success" -LogDirectory $logDir
            Test-Path $logDir | Should -BeTrue
            Remove-Item -Path $logDir -Recurse -Force
        }

        It "Should write log entry with all fields" {
            $entry = Write-AuditLog -Operation "Deploy" -RuleName "Test Rule" -Result "Success" -Details "Test details" -LogDirectory $script:TestLogDir
            $entry.Operation | Should -Be "Deploy"
            $entry.RuleName | Should -Be "Test Rule"
            $entry.Result | Should -Be "Success"
            $entry.Details | Should -Be "Test details"
            $entry.Timestamp | Should -Not -BeNullOrEmpty
        }

        It "Should create JSON log file" {
            Write-AuditLog -Operation "Test" -Result "Success" -LogDirectory $script:TestLogDir
            $logFile = Get-ChildItem -Path $script:TestLogDir -Filter "audit-*.json" | Select-Object -First 1
            $logFile | Should -Not -BeNullOrEmpty
        }
    }

    Context "Reading audit logs" {
        BeforeEach {
            # Clear and create fresh log
            if (Test-Path $script:TestLogDir) {
                Remove-Item -Path $script:TestLogDir -Recurse -Force
            }
            Write-AuditLog -Operation "Op1" -Result "Success" -LogDirectory $script:TestLogDir
            Write-AuditLog -Operation "Op2" -Result "Failure" -LogDirectory $script:TestLogDir
            Write-AuditLog -Operation "Op1" -Result "Warning" -LogDirectory $script:TestLogDir
        }

        It "Should retrieve all logs" {
            $logs = Get-AuditLogs -LogDirectory $script:TestLogDir
            $logs.Count | Should -BeGreaterOrEqual 3
        }

        It "Should filter by operation" {
            $logs = Get-AuditLogs -Operation "Op1" -LogDirectory $script:TestLogDir
            $logs | ForEach-Object { $_.Operation | Should -Be "Op1" }
        }

        It "Should return empty for non-existent directory" {
            $logs = Get-AuditLogs -LogDirectory "C:\NonExistent\Path"
            $logs | Should -BeNullOrEmpty
        }
    }
}
