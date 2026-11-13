<#
.SYNOPSIS
    Validates .env configuration and exercises the deployment scripts without
    changing Outlook or the registry.

.DESCRIPTION
    This is a safe smoke test:
      - Loads .env through DeploymentHelpers.ps1
      - Validates required values and derived settings (values are never printed)
      - Renders ABConfiguration.xml into a temporary directory
      - Runs Deploy-LdapAB.ps1 and Remove-LdapAB.ps1 in -DryRun mode
      - Parses every PowerShell script for syntax errors
      - Checks that secrets are not tracked by git
#>

[CmdletBinding()]
param(
    [string]$EnvFile = ''
)

$ErrorActionPreference = 'Stop'

$script:FailureCount = 0
$script:CheckCount = 0

function Write-Check {
    param(
        [string]$Name,
        [bool]$Passed,
        [string]$Detail = ''
    )

    $script:CheckCount++
    if ($Passed) {
        Write-Host "[PASS] $Name" -ForegroundColor Green
    } else {
        $script:FailureCount++
        if ($Detail) {
            Write-Host "[FAIL] $Name - $Detail" -ForegroundColor Red
        } else {
            Write-Host "[FAIL] $Name" -ForegroundColor Red
        }
    }
}

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

. (Join-Path $PSScriptRoot 'DeploymentHelpers.ps1')

$repoRoot = Split-Path -Parent $PSScriptRoot
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("LdapABTest_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

try {
    Write-Info 'Checking .env'

    $resolvedEnvFile = $EnvFile
    if ([string]::IsNullOrWhiteSpace($resolvedEnvFile)) {
        $resolvedEnvFile = Find-DotEnvFile -StartPath $PSScriptRoot
    }

    $envExists = [bool]($resolvedEnvFile -and (Test-Path -LiteralPath $resolvedEnvFile -PathType Leaf))
    Write-Check '.env file exists' $envExists 'Copy .env.example to .env and fill in the values.'

    if (-not $envExists) {
        throw 'Cannot continue without a .env file.'
    }

    try {
        $settings = Get-DeploymentSettings -EnvFile $resolvedEnvFile -SearchRoot $PSScriptRoot
        Write-Check '.env loads and required values are present' $true
    } catch {
        Write-Check '.env loads and required values are present' $false $_.Exception.Message
        throw
    }

    Write-Check 'DNS domain resolved' (-not [string]::IsNullOrWhiteSpace($settings.DomainFqdn))
    Write-Check 'NetBIOS domain resolved' (-not [string]::IsNullOrWhiteSpace($settings.NetBiosDomain))
    Write-Check 'NETLOGON path resolved' (-not [string]::IsNullOrWhiteSpace($settings.NetLogonPath))
    Write-Check 'LDAP port is valid' ($settings.LdapPort -ge 1 -and $settings.LdapPort -le 65535)
    Write-Check 'TARGET_OU looks like a distinguished name' ($settings.TargetOU -match '^\s*(OU|CN|DC)=')

    Write-Info 'Rendering client configuration'

    $renderedConfig = Join-Path $tempRoot 'ABConfiguration.xml'
    New-RenderedConfiguration `
        -TemplatePath (Join-Path $PSScriptRoot 'ABConfiguration.xml') `
        -DestinationPath $renderedConfig `
        -Settings $settings

    [xml]$renderedXml = Get-Content -LiteralPath $renderedConfig -ErrorAction Stop

    Write-Check 'Rendered ServerName matches .env' ($renderedXml.ABConfiguration.ServerName -eq $settings.LdapServer)
    Write-Check 'Rendered ServerPort matches .env' ($renderedXml.ABConfiguration.ServerPort -eq [string]$settings.LdapPort)
    Write-Check 'Rendered UseSSL matches .env' ($renderedXml.ABConfiguration.UseSSL -eq ([bool]$settings.UseSsl).ToString())
    Write-Check 'Client config uses SPA (no embedded credentials)' (
        [string]::IsNullOrEmpty($renderedXml.ABConfiguration.Username) -and
        [string]::IsNullOrEmpty($renderedXml.ABConfiguration.Password) -and
        $renderedXml.ABConfiguration.RequireSecurePasswordAuth -eq 'True'
    )

    Write-Info 'Checking PowerShell syntax'

    $scripts = @(
        'DeploymentHelpers.ps1',
        'Deploy-All.ps1',
        'Configure-LdapAB-GPO.ps1',
        'Deploy-LdapAB.ps1',
        'Remove-LdapAB.ps1',
        'Test-Configuration.ps1'
    )

    foreach ($name in $scripts) {
        $path = Join-Path $PSScriptRoot $name
        $tokens = $null
        $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors) | Out-Null
        $parseErrors = @($errors)
        $detail = ($parseErrors | ForEach-Object { $_.Message }) -join '; '
        Write-Check "Syntax: $name" ($parseErrors.Count -eq 0) $detail
    }

    Write-Info 'Dry-running client deployment'

    $deployScript = Join-Path $PSScriptRoot 'Deploy-LdapAB.ps1'
    $deployLog = Join-Path $tempRoot 'deploy.log'
    $deployOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $deployScript `
        -ConfigPath $renderedConfig -LogPath $deployLog -RegistryRoot 'LdapABDeployTest' -DryRun 2>&1 | Out-String
    $deployExit = $LASTEXITCODE
    Write-Check 'Deploy-LdapAB.ps1 dry run exits 0' ($deployExit -eq 0) "Exit code $deployExit"
    if ($deployExit -ne 0) {
        Write-Host $deployOutput -ForegroundColor DarkGray
    }

    Write-Info 'Dry-running client removal'

    $removeScript = Join-Path $PSScriptRoot 'Remove-LdapAB.ps1'
    $removeLog = Join-Path $tempRoot 'remove.log'
    $removeOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $removeScript `
        -ConfigPath $renderedConfig -LogPath $removeLog -RegistryRoot 'LdapABDeployTest' -DryRun 2>&1 | Out-String
    $removeExit = $LASTEXITCODE
    Write-Check 'Remove-LdapAB.ps1 dry run exits 0' ($removeExit -eq 0) "Exit code $removeExit"
    if ($removeExit -ne 0) {
        Write-Host $removeOutput -ForegroundColor DarkGray
    }

    Write-Info 'Validating deployment package generation'

    $deployAllScript = Join-Path $PSScriptRoot 'Deploy-All.ps1'
    $deployAllOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $deployAllScript `
        -SkipConnectivityTest -SkipGPO -SkipCopy 2>&1 | Out-String
    $deployAllExit = $LASTEXITCODE
    Write-Check 'Deploy-All.ps1 -SkipCopy exits 0' ($deployAllExit -eq 0) "Exit code $deployAllExit"
    if ($deployAllExit -ne 0) {
        Write-Host 'Deploy-All.ps1 failed during package generation.' -ForegroundColor DarkGray
    }

    Write-Info 'Checking repository hygiene'

    if (Get-Command git -ErrorAction SilentlyContinue) {
        Push-Location $repoRoot
        $previousErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $trackedEnv = @(& git ls-files .env 2>$null)
            Write-Check '.env is not tracked by git' ($trackedEnv.Count -eq 0)

            & git check-ignore -q .env 2>$null | Out-Null
            Write-Check '.env is ignored by git' ($LASTEXITCODE -eq 0)

            $remoteUrl = & git config --get remote.origin.url 2>$null
            Write-Check 'Git remote URL has no embedded credentials' (-not ($remoteUrl -match '://[^/@]+@'))

            $trackedFiles = @(& git ls-files)
            $textExtensions = @(
                '.ps1', '.xml', '.md', '.txt', '.example', '.json', '.yml', '.yaml',
                '.config', '.cpp', '.h'
            )
            $tokenPattern = 'ghp_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,}'
            $ipPattern = '\b(?:\d{1,3}\.){3}\d{1,3}\b'
            $passwordElementPattern = '<Pass' + 'word>\s*[^<\s]'
            $secretHits = @()

            foreach ($relativePath in $trackedFiles) {
                $extension = [System.IO.Path]::GetExtension($relativePath)
                if ($relativePath -eq '.env.example' -or $relativePath -eq '.gitignore') {
                    $extension = '.example'
                }
                if ($textExtensions -notcontains $extension) { continue }

                $fullPath = Join-Path $repoRoot $relativePath
                if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) { continue }

                $content = Get-Content -LiteralPath $fullPath -Raw -ErrorAction SilentlyContinue
                if (-not $content) { continue }

                if ($content -match $tokenPattern) {
                    $secretHits += "$relativePath (token pattern)"
                }
                if ($content -match $ipPattern) {
                    $secretHits += "$relativePath (IP address literal)"
                }
                if ($extension -eq '.xml' -and $content -match $passwordElementPattern) {
                    $secretHits += "$relativePath (non-empty Password element)"
                }
            }

            $uniqueHits = @($secretHits | Select-Object -Unique)
            Write-Check 'Tracked text files contain no token/IP/password literals' ($uniqueHits.Count -eq 0) ($uniqueHits -join '; ')
        } finally {
            $ErrorActionPreference = $previousErrorActionPreference
            Pop-Location
        }
    } else {
        Write-Check 'git available for repository hygiene checks' $false 'git was not found.'
    }

    if ($script:FailureCount -eq 0) {
        Write-Host "`nAll $script:CheckCount checks passed." -ForegroundColor Green
        $exitCode = 0
    } else {
        Write-Host "`n$script:FailureCount of $script:CheckCount checks failed." -ForegroundColor Red
        $exitCode = 1
    }
} catch {
    Write-Host "[FAIL] $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

exit $exitCode
