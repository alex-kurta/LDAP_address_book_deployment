<#
.SYNOPSIS
    Creates and links the Outlook LDAP Address Book GPO and configures its logon script.

.DESCRIPTION
    Run this on a domain controller or a management workstation with RSAT Group
    Policy tools installed. Values default to .env settings and can be
    overridden with parameters.

.PARAMETER EnvFile
    Path to the .env file. Defaults to the nearest .env in the project tree.

.PARAMETER GPOName
    Name of the GPO. Defaults to GPO_NAME in .env.

.PARAMETER TargetOU
    Distinguished name of the OU to link. Defaults to TARGET_OU in .env.

.PARAMETER ScriptPath
    Full UNC path to Deploy-LdapAB.ps1. Defaults to the NETLOGON path.

.PARAMETER DomainController
    Domain controller server name. Defaults to DC_SERVER in .env.

.PARAMETER DomainFqdn
    DNS domain name used for the SYSVOL policy path. Derived from .env when omitted.
#>

[CmdletBinding()]
param(
    [string]$EnvFile = '',
    [string]$GPOName = '',
    [string]$TargetOU = '',
    [string]$ScriptPath = '',
    [string]$DomainController = '',
    [string]$DomainFqdn = ''
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'DeploymentHelpers.ps1')

$settings = $null
$envError = $null
try {
    $settings = Get-DeploymentSettings -EnvFile $EnvFile -SearchRoot $PSScriptRoot
} catch {
    # .env is optional on a domain controller when every required value is
    # supplied as a parameter.
    $envError = $_.Exception.Message
}

if ([string]::IsNullOrWhiteSpace($GPOName)) {
    if ($settings) {
        $GPOName = $settings.GpoName
    } else {
        $GPOName = 'Outlook LDAP Address Book Deployment'
    }
}
if ([string]::IsNullOrWhiteSpace($TargetOU) -and $settings) { $TargetOU = $settings.TargetOU }
if ([string]::IsNullOrWhiteSpace($ScriptPath) -and $settings) { $ScriptPath = $settings.ScriptPath }
if ([string]::IsNullOrWhiteSpace($DomainController) -and $settings) { $DomainController = $settings.DomainController }
if ([string]::IsNullOrWhiteSpace($DomainFqdn) -and $settings) { $DomainFqdn = $settings.DomainFqdn }

if ([string]::IsNullOrWhiteSpace($DomainFqdn) -and -not [string]::IsNullOrWhiteSpace($TargetOU)) {
    $DomainFqdn = Get-DomainFqdnFromDistinguishedName -DistinguishedName $TargetOU
}
if ([string]::IsNullOrWhiteSpace($ScriptPath) -and -not [string]::IsNullOrWhiteSpace($DomainController)) {
    $ScriptPath = "\\$DomainController\NETLOGON\LdapABDeploy\Deploy-LdapAB.ps1"
}

if ([string]::IsNullOrWhiteSpace($TargetOU)) {
    $message = 'TargetOU is required. Pass -TargetOU or provide TARGET_OU in .env.'
    if ($envError) { $message = "$message $envError" }
    Write-Host "[ERROR] $message" -ForegroundColor Red
    exit 1
}
if ([string]::IsNullOrWhiteSpace($DomainController)) {
    $message = 'DomainController is required. Pass -DomainController or provide DC_SERVER in .env.'
    if ($envError) { $message = "$message $envError" }
    Write-Host "[ERROR] $message" -ForegroundColor Red
    exit 1
}
if ([string]::IsNullOrWhiteSpace($DomainFqdn)) {
    $message = 'DomainFqdn is required. Pass -DomainFqdn, use a TARGET_OU with DC= components, or provide DOMAIN in .env.'
    if ($envError) { $message = "$message $envError" }
    Write-Host "[ERROR] $message" -ForegroundColor Red
    exit 1
}

if ($settings) {
    $displayName = $settings.DisplayName
} else {
    $displayName = 'LDAP Address Book'
}

try {
    Import-Module GroupPolicy -ErrorAction Stop

    $gpo = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
    if (-not $gpo) {
        Write-Host "Creating GPO: $GPOName" -ForegroundColor Green
        $gpo = New-GPO -Name $GPOName -Comment "Deploy $displayName LDAP Address Book to Outlook via MAPI"
    } else {
        Write-Host "GPO already exists: $GPOName" -ForegroundColor Yellow
    }

    $link = Get-GPLink -Name $GPOName -Target $TargetOU -ErrorAction SilentlyContinue
    if (-not $link) {
        Write-Host "Linking GPO to $TargetOU" -ForegroundColor Green
        New-GPLink -Name $GPOName -Target $TargetOU -Enforced:$false | Out-Null
    } else {
        Write-Host "GPO already linked to $TargetOU" -ForegroundColor Yellow
    }

    # Configure the user logon script through the GPO's psscripts.ini file.
    $gpoPath = "\\$DomainController\SYSVOL\$DomainFqdn\Policies\$($gpo.Id)\User\Scripts\Logon"
    New-Item -ItemType Directory -Path $gpoPath -Force | Out-Null

    $scriptsIni = Join-Path $gpoPath 'psscripts.ini'
    $iniContent = @"
[Logon]
0CmdLine=$ScriptPath
0Parameters=
"@

    Set-Content -LiteralPath $scriptsIni -Value $iniContent -Encoding ASCII
    Write-Host "Configured logon script: $ScriptPath" -ForegroundColor Green

    Write-Host "`nGPO Configuration Complete!" -ForegroundColor Green
    Write-Host "GPO:         $GPOName"
    Write-Host "Linked to:   $TargetOU"
    Write-Host "Script:      $ScriptPath"
    Write-Host "`nManually enable: User Configuration > Administrative Templates > System > Scripts > Run logon scripts synchronously" -ForegroundColor Cyan
} catch {
    Write-Host "[ERROR] GPO configuration failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
