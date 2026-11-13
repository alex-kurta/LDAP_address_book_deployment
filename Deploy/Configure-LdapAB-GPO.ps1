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

.PARAMETER LinkAtDomainRoot
    Link the GPO at the domain root, which applies it company-wide. This
    overrides TargetOU.

.PARAMETER ScriptPath
    Full UNC path to Deploy-LdapAB.ps1. Defaults to the NETLOGON path.

.PARAMETER DomainController
    Domain controller server name. Defaults to DC_SERVER in .env.

.PARAMETER DomainFqdn
    DNS domain name used for the SYSVOL policy path. Derived from .env when omitted.

.PARAMETER SkipSynchronousScripts
    Do not enable the Run logon scripts synchronously policy.
#>

[CmdletBinding()]
param(
    [string]$EnvFile = '',
    [string]$GPOName = '',
    [string]$TargetOU = '',
    [string]$ScriptPath = '',
    [string]$DomainController = '',
    [string]$DomainFqdn = '',
    [switch]$LinkAtDomainRoot,
    [switch]$SkipSynchronousScripts
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
if ($LinkAtDomainRoot) {
    if ([string]::IsNullOrWhiteSpace($DomainFqdn)) {
        $message = 'Cannot determine DomainFqdn for -LinkAtDomainRoot.'
        if ($envError) { $message = "$message $envError" }
        Write-Host "[ERROR] $message" -ForegroundColor Red
        exit 1
    }
    $TargetOU = Get-DistinguishedNameFromDomainFqdn -DomainFqdn $DomainFqdn
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

try {
    $result = Invoke-LdapABGpoConfiguration `
        -GPOName $GPOName `
        -TargetOU $TargetOU `
        -ScriptPath $ScriptPath `
        -DomainFqdn $DomainFqdn `
        -DomainController $DomainController `
        -SkipSynchronousScripts:$SkipSynchronousScripts

    if ($result.Created) {
        Write-Host "Created GPO: $($result.GPOName)" -ForegroundColor Green
    } else {
        Write-Host "GPO already exists: $($result.GPOName)" -ForegroundColor Yellow
    }
    if ($result.Linked) {
        Write-Host "Linked GPO to $($result.LinkTarget)" -ForegroundColor Green
    } else {
        Write-Host "GPO already linked to $($result.LinkTarget)" -ForegroundColor Yellow
    }
    Write-Host "Configured logon script: $ScriptPath" -ForegroundColor Green
    if ($result.LegacyScriptIniRemoved) {
        Write-Host 'Removed the obsolete Scripts\Logon\psscripts.ini copy' -ForegroundColor Green
    }
    if ($result.ScriptsExtensionRegistered) {
        Write-Host 'Registered the Group Policy Scripts client-side extension' -ForegroundColor Green
    }
    if ($result.SyncPolicySet) {
        Write-Host 'Enabled: Run logon scripts synchronously' -ForegroundColor Green
    }
    if ($result.UserVersion -gt 0) {
        Write-Host "GPO user version: $($result.UserVersion)" -ForegroundColor Green
    }

    Write-Host "`nGPO Configuration Complete!" -ForegroundColor Green
    Write-Host "GPO:         $($result.GPOName)"
    Write-Host "GPO ID:      $($result.GpoId)"
    Write-Host "Linked to:   $($result.LinkTarget)"
    Write-Host "Script:      $ScriptPath"
} catch {
    Write-Host "[ERROR] GPO configuration failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
