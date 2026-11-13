<#
.SYNOPSIS
    Creates and links the Outlook LDAP Address Book GPO through WinRM.

.DESCRIPTION
    Use this script on a workstation that does not have RSAT or the GroupPolicy
    module. It opens a PowerShell remoting session to the domain controller,
    creates or updates the GPO there, links it, installs the user logon script,
    and enables synchronous logon scripts.

    Deployment credentials are read from .env (USERNAME and PASSWORD) or from
    parameters. They are used only for the WinRM connection and are never
    printed or written to disk.

    Run with -DryRun first to see what would change without modifying AD.

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
    Domain controller used for WinRM. Defaults to DC_SERVER in .env.

.PARAMETER DomainFqdn
    DNS domain name used for the SYSVOL policy path. Derived from .env when
    omitted.

.PARAMETER AdminUser
    Account used for the WinRM connection. Defaults to USERNAME in .env.

.PARAMETER AdminPassword
    Password used for the WinRM connection. Defaults to PASSWORD in .env.

.PARAMETER SkipSynchronousScripts
    Do not enable the Run logon scripts synchronously policy.

.PARAMETER DryRun
    Report what would change without creating, linking, or configuring anything.
#>

[CmdletBinding()]
param(
    [string]$EnvFile = '',
    [string]$GPOName = '',
    [string]$TargetOU = '',
    [string]$ScriptPath = '',
    [string]$DomainController = '',
    [string]$DomainFqdn = '',
    [string]$AdminUser = '',
    [string]$AdminPassword = '',
    [switch]$LinkAtDomainRoot,
    [switch]$SkipSynchronousScripts,
    [switch]$DryRun
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

$settings = $null
$envError = $null
try {
    $settings = Get-DeploymentSettings -EnvFile $EnvFile -SearchRoot $PSScriptRoot
} catch {
    # .env is optional when every GPO value is supplied as a parameter.
    $envError = $_.Exception.Message
}

if ($settings) {
    if ([string]::IsNullOrWhiteSpace($GPOName)) { $GPOName = $settings.GpoName }
    if ([string]::IsNullOrWhiteSpace($TargetOU)) { $TargetOU = $settings.TargetOU }
    if ([string]::IsNullOrWhiteSpace($ScriptPath)) { $ScriptPath = $settings.ScriptPath }
    if ([string]::IsNullOrWhiteSpace($DomainController)) { $DomainController = $settings.DomainController }
    if ([string]::IsNullOrWhiteSpace($DomainFqdn)) { $DomainFqdn = $settings.DomainFqdn }
    if ([string]::IsNullOrWhiteSpace($AdminUser)) { $AdminUser = $settings.AdminUser }
    if ([string]::IsNullOrWhiteSpace($AdminPassword)) { $AdminPassword = $settings.AdminPassword }
}

if ([string]::IsNullOrWhiteSpace($GPOName)) {
    $GPOName = 'Outlook LDAP Address Book Deployment'
}

if ($LinkAtDomainRoot) {
    if ([string]::IsNullOrWhiteSpace($DomainFqdn) -and -not [string]::IsNullOrWhiteSpace($TargetOU)) {
        $DomainFqdn = Get-DomainFqdnFromDistinguishedName -DistinguishedName $TargetOU
    }
    if ([string]::IsNullOrWhiteSpace($DomainFqdn)) {
        $message = 'Cannot determine the domain DNS name for -LinkAtDomainRoot.'
        if ($envError) { $message = "$message $envError" }
        Write-Fail $message
        exit 1
    }
    $TargetOU = Get-DistinguishedNameFromDomainFqdn -DomainFqdn $DomainFqdn
} elseif ([string]::IsNullOrWhiteSpace($TargetOU)) {
    $message = 'TargetOU is required. Pass -TargetOU, use -LinkAtDomainRoot, or provide TARGET_OU in .env.'
    if ($envError) { $message = "$message $envError" }
    Write-Fail $message
    exit 1
}

if ([string]::IsNullOrWhiteSpace($DomainFqdn)) {
    $DomainFqdn = Get-DomainFqdnFromDistinguishedName -DistinguishedName $TargetOU
}
if ([string]::IsNullOrWhiteSpace($DomainController)) {
    $message = 'DomainController is required. Pass -DomainController or provide DC_SERVER in .env.'
    if ($envError) { $message = "$message $envError" }
    Write-Fail $message
    exit 1
}
if ([string]::IsNullOrWhiteSpace($DomainFqdn)) {
    $message = 'DomainFqdn is required. Pass -DomainFqdn, use a TargetOU with DC= components, or provide DOMAIN in .env.'
    if ($envError) { $message = "$message $envError" }
    Write-Fail $message
    exit 1
}
if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
    $ScriptPath = "\\$DomainController\NETLOGON\LdapABDeploy\Deploy-LdapAB.ps1"
}

$credential = $null
if (-not [string]::IsNullOrWhiteSpace($AdminUser) -and -not [string]::IsNullOrWhiteSpace($AdminPassword)) {
    $securePassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($AdminUser, $securePassword)
} elseif (-not [string]::IsNullOrWhiteSpace($AdminUser) -or -not [string]::IsNullOrWhiteSpace($AdminPassword)) {
    Write-Fail 'AdminUser and AdminPassword must be provided together.'
    exit 1
}

Write-Host @"
+---------------------------------------------------------------+
|  Outlook LDAP Address Book - Remote GPO Configuration         |
+---------------------------------------------------------------+
"@ -ForegroundColor Cyan

Write-Step '1. Target'
Write-Host "Domain controller: $DomainController"
Write-Host "Domain:            $DomainFqdn"
Write-Host "GPO:               $GPOName"
Write-Host "Link target:       $TargetOU"
Write-Host "Logon script:      $ScriptPath"
if ($DryRun) {
    Write-Warn 'Dry run: no Active Directory, SYSVOL, or policy changes will be made'
}

$sharedHelperNames = @(
    'Get-LdapABScriptsExtensionName',
    'Get-LdapABGpoDirectoryEntry',
    'Get-GpoSysvolPolicyPath',
    'Set-LdapABGpoScriptsIni',
    'Register-LdapABGpoScriptsExtension',
    'Update-LdapABGpoVersion'
)
$sharedHelperSource = (@($sharedHelperNames | ForEach-Object {
    $command = Get-Command -Name $_ -CommandType Function -ErrorAction SilentlyContinue
    if (-not $command) {
        throw "Required helper function not found: $_"
    }
    "function $_ {`r`n$($command.Definition)`r`n}"
}) -join "`r`n`r`n")

$remoteBody = {
    param(
        [string]$GPOName,
        [string]$TargetOU,
        [string]$ScriptPath,
        [string]$DomainFqdn,
        [bool]$SkipSynchronousScripts,
        [bool]$DryRun
    )

    # __SHARED_HELPERS__

    $ErrorActionPreference = 'Stop'
    Import-Module GroupPolicy -ErrorAction Stop

    $state = [ordered]@{
        GPOName                    = $GPOName
        GpoId                      = $null
        GpoExisted                 = $false
        GpoCreated                 = $false
        LinkTarget                 = $TargetOU
        LinkExisted                = $false
        LinkCreated                = $false
        ScriptIni                  = $null
        LegacyScriptIniRemoved     = $false
        ScriptsExtensionRegistered = $false
        SyncPolicySet              = $false
        UserVersion                = 0
        DryRun                     = $DryRun
    }

    $gpo = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
    if ($gpo) {
        $state.GpoExisted = $true
        $state.GpoId = [string]$gpo.Id
    } elseif (-not $DryRun) {
        $gpo = New-GPO -Name $GPOName -Comment 'Deploy LDAP Address Book to Outlook via MAPI'
        $state.GpoId = [string]$gpo.Id
        $state.GpoCreated = $true
    }

    try {
        $inheritance = Get-GPInheritance -Target $TargetOU -ErrorAction Stop
    } catch {
        throw "Cannot read GPO links for target '$TargetOU'. Verify that the OU or domain distinguished name exists. $($_.Exception.Message)"
    }

    $existingLink = $null
    if ($gpo) {
        $existingLink = @($inheritance.GpoLinks) |
            Where-Object { [string]$_.GpoId -eq [string]$gpo.Id } |
            Select-Object -First 1
    }
    if ($existingLink) {
        $state.LinkExisted = $true
    } elseif (-not $DryRun) {
        New-GPLink -Name $GPOName -Target $TargetOU -ErrorAction Stop | Out-Null
        $state.LinkCreated = $true
    }

    if ($DryRun) {
        return [pscustomobject]$state
    }

    # On a domain controller this is the local SYSVOL path, which avoids a
    # second network hop inside the WinRM session.
    $policyPath = Get-GpoSysvolPolicyPath -GpoId ([string]$gpo.Id) -DomainFqdn $DomainFqdn

    $scriptsResult = Set-LdapABGpoScriptsIni -PolicyPath $policyPath -ScriptPath $ScriptPath
    $state.ScriptIni = $scriptsResult.ScriptIni
    $state.LegacyScriptIniRemoved = $scriptsResult.LegacyRemoved

    $state.ScriptsExtensionRegistered = Register-LdapABGpoScriptsExtension -GpoId ([string]$gpo.Id)

    if (-not $SkipSynchronousScripts) {
        # User Configuration > Administrative Templates > System > Scripts >
        # Run logon scripts synchronously.
        Set-GPRegistryValue `
            -Name $GPOName `
            -Key 'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System' `
            -ValueName 'RunLogonScriptSync' `
            -Type DWord `
            -Value 1 `
            -ErrorAction Stop | Out-Null
        $state.SyncPolicySet = $true
    }

    # Bump the user version after all policy files are written so clients are
    # guaranteed to process the logon script on their next refresh.
    $state.UserVersion = Update-LdapABGpoVersion `
        -GpoId ([string]$gpo.Id) `
        -PolicyPath $policyPath `
        -GPOName $GPOName

    return [pscustomobject]$state
}

$remoteScript = [scriptblock]::Create(
    $remoteBody.ToString().Replace('# __SHARED_HELPERS__', $sharedHelperSource)
)

Write-Step '2. Connect and configure'

$invokeParameters = @{
    ComputerName = $DomainController
    ScriptBlock  = $remoteScript
    ArgumentList = @(
        $GPOName,
        $TargetOU,
        $ScriptPath,
        $DomainFqdn,
        [bool]$SkipSynchronousScripts,
        [bool]$DryRun
    )
    ErrorAction  = 'Stop'
}
if ($credential) {
    $invokeParameters['Credential'] = $credential
    Write-Host 'Using deployment credentials from .env or parameters.'
} else {
    Write-Warn 'No deployment credentials supplied; using the current Windows account.'
}

try {
    $result = Invoke-Command @invokeParameters
} catch {
    Write-Fail "Remote GPO configuration failed: $($_.Exception.Message)"
    exit 1
}

Write-Step '3. Result'

if ($result.GpoExisted) {
    Write-Ok "GPO exists: $($result.GPOName) ($($result.GpoId))"
} elseif ($result.DryRun) {
    Write-Warn "GPO does not exist and would be created: $($result.GPOName)"
} else {
    Write-Ok "Created GPO: $($result.GPOName) ($($result.GpoId))"
}

if ($result.LinkExisted) {
    Write-Ok "GPO already linked to $($result.LinkTarget)"
} elseif ($result.DryRun) {
    Write-Warn "GPO is not linked and would be linked to $($result.LinkTarget)"
} else {
    Write-Ok "Linked GPO to $($result.LinkTarget)"
}

if ($result.DryRun) {
    Write-Host "`nDry run complete. No changes were made." -ForegroundColor Yellow
} else {
    if ($result.ScriptIni) {
        Write-Ok "Configured logon script in $($result.ScriptIni)"
    }
    if ($result.LegacyScriptIniRemoved) {
        Write-Ok 'Removed the obsolete Scripts\Logon\psscripts.ini copy'
    }
    if ($result.ScriptsExtensionRegistered) {
        Write-Ok 'Registered the Group Policy Scripts client-side extension'
    }
    if ($result.SyncPolicySet) {
        Write-Ok "Enabled 'Run logon scripts synchronously'"
    }
    if ($result.UserVersion -gt 0) {
        Write-Ok "GPO user version: $($result.UserVersion)"
    }
    Write-Host "`nRemote GPO configuration complete." -ForegroundColor Green
    Write-Host 'Allow time for SYSVOL and Group Policy replication before testing clients.' -ForegroundColor Cyan
}
