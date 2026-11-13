<#
.SYNOPSIS
    Complete automation for Outlook LDAP Address Book GPO deployment.

.DESCRIPTION
    Loads deployment values from .env (or explicit parameters), renders the
    client configuration, copies only the client-side files to NETLOGON, and
    optionally creates and links the GPO.

    The .env file is read locally and is never copied to the domain controller.

.PARAMETER EnvFile
    Path to the .env file. Defaults to the nearest .env in the project tree.

.PARAMETER DomainController
    Domain controller server name. Defaults to DC_SERVER in .env.

.PARAMETER TargetOU
    Target OU for GPO linking. Defaults to TARGET_OU in .env.

.PARAMETER LDAPServer
    LDAP server host name or address. Defaults to LDAP_SERVER in .env.

.PARAMETER LDAPPort
    LDAP port. Defaults to LDAP_PORT in .env.

.PARAMETER UseSSL
    true or false. Defaults to USE_SSL in .env.

.PARAMETER SkipGPO
    Skip GPO creation and linking.

.PARAMETER SkipConnectivityTest
    Skip the domain controller connectivity check.

.PARAMETER SkipCopy
    Render and validate the deployment package without copying anything to
    NETLOGON. GPO configuration is also skipped.

.PARAMETER AdminUser
    Account used for SMB mapping to NETLOGON. Defaults to USERNAME in .env.

.PARAMETER AdminPassword
    Password used for SMB mapping. Defaults to PASSWORD in .env. If empty,
    the current Windows credentials are used.
#>

[CmdletBinding()]
param(
    [string]$EnvFile = '',
    [string]$DomainController = '',
    [string]$TargetOU = '',
    [string]$LDAPServer = '',
    [int]$LDAPPort = 0,
    [string]$UseSSL = '',
    [switch]$SkipGPO,
    [switch]$SkipConnectivityTest,
    [switch]$SkipCopy,
    [string]$AdminUser = '',
    [string]$AdminPassword = ''
)

$ErrorActionPreference = 'Stop'

function Write-Step {
    param([string]$Message)
    Write-Host "`n=== $Message ===" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Message)
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Message)
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Write-Fail {
    param([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

. (Join-Path $PSScriptRoot 'DeploymentHelpers.ps1')

try {
    $settings = Get-DeploymentSettings -EnvFile $EnvFile -SearchRoot $PSScriptRoot
} catch {
    Write-Fail $_.Exception.Message
    exit 1
}

# Explicit parameters override .env values.
if ($PSBoundParameters.ContainsKey('DomainController') -and $DomainController) {
    $settings.DomainController = $DomainController
    $settings.NetLogonPath = "\\$DomainController\NETLOGON"
    $settings.ScriptPath = "$($settings.NetLogonPath)\LdapABDeploy\Deploy-LdapAB.ps1"
}
if ($PSBoundParameters.ContainsKey('TargetOU') -and $TargetOU) {
    $settings.TargetOU = $TargetOU
}
if ($PSBoundParameters.ContainsKey('LDAPServer') -and $LDAPServer) {
    $settings.LdapServer = $LDAPServer
}
if ($PSBoundParameters.ContainsKey('LDAPPort') -and $LDAPPort -gt 0) {
    $settings.LdapPort = $LDAPPort
}
if ($PSBoundParameters.ContainsKey('UseSSL') -and $UseSSL) {
    switch -Regex ($UseSSL.Trim()) {
        '^(?i:true|1|yes)$'  { $settings.UseSsl = $true; break }
        '^(?i:false|0|no)$'  { $settings.UseSsl = $false; break }
        default {
            Write-Fail 'UseSSL must be true or false.'
            exit 1
        }
    }
}
if ($PSBoundParameters.ContainsKey('AdminUser') -and $AdminUser) {
    $settings.AdminUser = $AdminUser
}
if ($PSBoundParameters.ContainsKey('AdminPassword') -and $AdminPassword) {
    $settings.AdminPassword = $AdminPassword
}

Write-Host @"
+---------------------------------------------------------------+
|  Outlook LDAP Address Book GPO Deployment                     |
+---------------------------------------------------------------+
"@ -ForegroundColor Cyan

$startTime = Get-Date
$deploySource = $PSScriptRoot
$uncDestinationPath = "\\$($settings.DomainController)\NETLOGON\LdapABDeploy"
$destinationPath = $uncDestinationPath
$stagingPath = Join-Path ([System.IO.Path]::GetTempPath()) ("LdapABDeploy_" + [guid]::NewGuid().ToString('N'))
$useSmbMap = $false
$driveLetter = $null

Write-Step '1. Environment Validation'

$missingFiles = @()
foreach ($file in (Get-ClientDeploymentFiles)) {
    $path = Join-Path $deploySource $file
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $missingFiles += $file
    }
}
if ($missingFiles.Count -gt 0) {
    Write-Fail ("Missing deployment files: " + ($missingFiles -join ', '))
    exit 1
}
Write-Ok 'Deployment package files are present'

if ($SkipConnectivityTest) {
    Write-Warn 'Domain controller connectivity check skipped'
} else {
    try {
        if (Test-Connection -ComputerName $settings.DomainController -Count 2 -Quiet) {
            Write-Ok "Domain controller reachable: $($settings.DomainController)"
        } else {
            Write-Fail "Cannot reach domain controller: $($settings.DomainController)"
            exit 1
        }
    } catch {
        Write-Fail "Connection test failed: $($_.Exception.Message)"
        exit 1
    }
}

Write-Step '2. Prepare Deployment Package'

try {
    New-DeploymentPackage `
        -SourcePath $deploySource `
        -DestinationPath $stagingPath `
        -Settings $settings | Out-Null
    Write-Ok "Rendered package for $($settings.LdapServer):$($settings.LdapPort) (SSL: $($settings.UseSsl))"
} catch {
    Write-Fail $_.Exception.Message
    Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

Write-Step '3. Copy Files to NETLOGON'

if ($SkipCopy) {
    Write-Warn 'NETLOGON copy skipped (-SkipCopy)'
    Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
} else {
    try {
        if (-not [string]::IsNullOrWhiteSpace($settings.AdminUser) -and -not [string]::IsNullOrWhiteSpace($settings.AdminPassword)) {
            for ($i = 25; $i -ge 0; $i--) {
                $letter = [char](65 + $i)
                $candidate = "$letter`:"
                if (Test-Path -LiteralPath $candidate) { continue }

                try {
                    New-SmbMapping `
                        -LocalPath $candidate `
                        -RemotePath "\\$($settings.DomainController)\NETLOGON" `
                        -UserName $settings.AdminUser `
                        -Password $settings.AdminPassword `
                        -Persistent:$false `
                        -ErrorAction Stop | Out-Null
                    $driveLetter = $candidate
                    $useSmbMap = $true
                    break
                } catch {
                    continue
                }
            }
        }

        if ($useSmbMap) {
            $destinationPath = "$driveLetter\LdapABDeploy"
        }

        New-Item -ItemType Directory -Path $destinationPath -Force -ErrorAction Stop | Out-Null
        Copy-Item -Path (Join-Path $stagingPath '*') -Destination $destinationPath -Force -ErrorAction Stop
        Write-Ok "Client files copied to $destinationPath"
    } catch {
        Write-Fail "Failed to copy files: $($_.Exception.Message)"
        Write-Warn 'Run with -AdminUser and -AdminPassword, or copy the staged files manually.'
        exit 1
    } finally {
        if ($useSmbMap -and $driveLetter) {
            Remove-SmbMapping -LocalPath $driveLetter -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Step '4. GPO Configuration'

$scriptFullPath = "$($settings.NetLogonPath)\LdapABDeploy\Deploy-LdapAB.ps1"

if ($SkipGPO -or $SkipCopy) {
    Write-Warn 'GPO creation skipped'
} elseif ([string]::IsNullOrWhiteSpace($settings.TargetOU)) {
    Write-Warn 'TARGET_OU is empty - GPO not created or linked'
} elseif (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
    Write-Warn 'GroupPolicy module not available'
    Write-Warn 'Run Configure-LdapAB-GPO.ps1 on a domain controller with RSAT installed'
} else {
    try {
        Import-Module GroupPolicy -ErrorAction Stop

        $gpo = Get-GPO -Name $settings.GpoName -ErrorAction SilentlyContinue
        if (-not $gpo) {
            $gpo = New-GPO -Name $settings.GpoName -Comment "Deploy $($settings.DisplayName) LDAP Address Book"
            Write-Ok "Created GPO: $($settings.GpoName)"
        } else {
            Write-Ok "GPO already exists: $($settings.GpoName)"
        }

        $link = Get-GPLink -Name $settings.GpoName -Target $settings.TargetOU -ErrorAction SilentlyContinue
        if (-not $link) {
            New-GPLink -Name $settings.GpoName -Target $settings.TargetOU -Enforced:$false | Out-Null
            Write-Ok "Linked GPO to $($settings.TargetOU)"
        } else {
            Write-Ok "GPO already linked to $($settings.TargetOU)"
        }

        Write-Warn "Enable 'Run logon scripts synchronously' in the GPO"
        Write-Warn "Logon script: $scriptFullPath"
    } catch {
        Write-Fail "GPO configuration failed: $($_.Exception.Message)"
        Write-Warn 'Use Configure-LdapAB-GPO.ps1 on the domain controller, or configure the GPO manually.'
    }
}

Write-Step '5. Deployment Summary'

$duration = (Get-Date) - $startTime
Write-Host "Duration: $($duration.TotalSeconds.ToString('0.00')) seconds" -ForegroundColor Gray
if ($SkipCopy) {
    Write-Host "`nPackage generation validated (NETLOGON copy skipped)" -ForegroundColor Green
} else {
    Write-Host "`nDeployment package: $uncDestinationPath" -ForegroundColor Green
}
Write-Host "Logon script:       $scriptFullPath" -ForegroundColor Green
Write-Host "`nNext Steps:" -ForegroundColor Yellow
Write-Host '1. Create or link the GPO if it was not configured automatically'
Write-Host "2. Set the logon script to: $scriptFullPath"
Write-Host "3. Enable 'Run logon scripts synchronously'"
Write-Host '4. Test on a pilot machine'
Write-Host '5. Monitor logs: %TEMP%\LdapABDeploy.log'
Write-Host "`nRollback: run Remove-LdapAB.ps1 or unlink the GPO" -ForegroundColor Cyan
Write-Host "`nDone!" -ForegroundColor Green
