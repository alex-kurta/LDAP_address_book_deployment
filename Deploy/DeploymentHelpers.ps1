<#
.SYNOPSIS
    Shared helpers for loading .env settings and rendering the client configuration.

.DESCRIPTION
    Dot-source this file from administrative scripts. It never prints secret
    values; validation errors list missing key names only.
#>

function Find-DotEnvFile {
    [CmdletBinding()]
    param(
        [string]$StartPath = (Get-Location).Path
    )

    $directory = [System.IO.Path]::GetFullPath($StartPath)

    while ($directory) {
        $candidate = Join-Path $directory '.env'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }

        $parent = Split-Path -Parent $directory
        if ([string]::IsNullOrEmpty($parent) -or $parent -eq $directory) {
            break
        }
        $directory = $parent
    }

    return $null
}

function Import-DotEnvFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Environment file not found: $Path"
    }

    foreach ($line in (Get-Content -LiteralPath $Path -ErrorAction Stop)) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        if ($trimmed.StartsWith('#')) { continue }

        if ($trimmed -notmatch '^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$') {
            continue
        }

        $key = $matches[1]
        $value = $matches[2].Trim()

        if ($value.Length -ge 2) {
            $first = $value[0]
            $last = $value[$value.Length - 1]
            if (($first -eq '"' -and $last -eq '"') -or ($first -eq "'" -and $last -eq "'")) {
                $value = $value.Substring(1, $value.Length - 2)
            } else {
                # Strip an unquoted inline comment only when it follows whitespace.
                $value = ($value -replace '\s+#.*$', '').Trim()
            }
        }

        [System.Environment]::SetEnvironmentVariable(
            $key,
            $value,
            [System.EnvironmentVariableTarget]::Process
        )
    }
}

function Get-DomainFqdnFromDistinguishedName {
    [CmdletBinding()]
    param(
        [string]$DistinguishedName
    )

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return $null }

    $parts = @()
    foreach ($rdn in ($DistinguishedName -split ',')) {
        $item = $rdn.Trim()
        if ($item -match '^DC=(.+)$') {
            $parts += $matches[1]
        }
    }

    if ($parts.Count -gt 0) {
        return ($parts -join '.')
    }

    return $null
}

function Get-DeploymentSettings {
    [CmdletBinding()]
    param(
        [string]$EnvFile = '',
        [string]$SearchRoot = (Get-Location).Path
    )

    if ([string]::IsNullOrWhiteSpace($EnvFile)) {
        $EnvFile = Find-DotEnvFile -StartPath $SearchRoot
    } elseif (-not [System.IO.Path]::IsPathRooted($EnvFile)) {
        $EnvFile = Join-Path $SearchRoot $EnvFile
    }

    if (-not $EnvFile -or -not (Test-Path -LiteralPath $EnvFile -PathType Leaf)) {
        throw 'No .env file found. Copy .env.example to .env and fill in the values.'
    }

    Import-DotEnvFile -Path $EnvFile

    $required = @(
        'DOMAIN',
        'DC_SERVER',
        'LDAP_SERVER',
        'LDAP_PORT',
        'USE_SSL',
        'LDAP_DISPLAY_NAME',
        'TARGET_OU'
    )

    $missing = @()
    foreach ($key in $required) {
        if ([string]::IsNullOrWhiteSpace([System.Environment]::GetEnvironmentVariable($key, 'Process'))) {
            $missing += $key
        }
    }

    if ($missing.Count -gt 0) {
        throw ("Missing required .env values: " + ($missing -join ', '))
    }

    $domain = $env:DOMAIN.Trim()
    $targetOu = $env:TARGET_OU.Trim()
    $dcServer = $env:DC_SERVER.Trim()

    $domainFqdn = $env:DOMAIN_FQDN
    if ([string]::IsNullOrWhiteSpace($domainFqdn) -and $domain -match '\.') {
        $domainFqdn = $domain
    }
    if ([string]::IsNullOrWhiteSpace($domainFqdn)) {
        $domainFqdn = Get-DomainFqdnFromDistinguishedName -DistinguishedName $targetOu
    }
    if ([string]::IsNullOrWhiteSpace($domainFqdn)) {
        $domainFqdn = $env:USERDNSDOMAIN
    }
    if ([string]::IsNullOrWhiteSpace($domainFqdn)) {
        throw 'Could not determine the DNS domain name. Set DOMAIN to the FQDN or add DOMAIN_FQDN to .env.'
    }

    $netBiosDomain = $env:NETBIOS_DOMAIN
    if ([string]::IsNullOrWhiteSpace($netBiosDomain)) {
        if ($domain -notmatch '\.') {
            $netBiosDomain = $domain
        } elseif ($env:USERDOMAIN) {
            $netBiosDomain = $env:USERDOMAIN
        } else {
            $netBiosDomain = ($domainFqdn -split '\.')[0]
        }
    }

    $netLogonPath = $env:NETLOGON_PATH
    if ([string]::IsNullOrWhiteSpace($netLogonPath)) {
        $netLogonPath = "\\$dcServer\NETLOGON"
    }
    $netLogonPath = $netLogonPath.TrimEnd('\')

    $portValue = 0
    if (-not [int]::TryParse($env:LDAP_PORT, [ref]$portValue) -or $portValue -lt 1 -or $portValue -gt 65535) {
        throw 'LDAP_PORT must be an integer between 1 and 65535.'
    }

    $useSsl = $false
    switch -Regex ($env:USE_SSL.Trim()) {
        '^(?i:true|1|yes)$'  { $useSsl = $true; break }
        '^(?i:false|0|no)$'  { $useSsl = $false; break }
        default { throw 'USE_SSL must be true or false.' }
    }

    $gpoName = $env:GPO_NAME
    if ([string]::IsNullOrWhiteSpace($gpoName)) {
        $gpoName = 'Outlook LDAP Address Book Deployment'
    }

    $adminUser = $env:USERNAME
    if (-not [string]::IsNullOrWhiteSpace($adminUser) -and $adminUser -notmatch '[\\@]') {
        $adminUser = "$netBiosDomain\$adminUser"
    }

    [pscustomobject]@{
        EnvFile          = $EnvFile
        Domain           = $domain
        DomainFqdn       = $domainFqdn
        NetBiosDomain    = $netBiosDomain
        DomainController = $dcServer
        NetLogonPath     = $netLogonPath
        LdapServer       = $env:LDAP_SERVER.Trim()
        LdapPort         = $portValue
        UseSsl           = $useSsl
        DisplayName      = $env:LDAP_DISPLAY_NAME.Trim()
        TargetOU         = $targetOu
        GpoName          = $gpoName
        AdminUser        = $adminUser
        AdminPassword    = $env:PASSWORD
        ScriptPath       = "$netLogonPath\LdapABDeploy\Deploy-LdapAB.ps1"
    }
}

function New-RenderedConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplatePath,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        [Parameter(Mandatory = $true)]
        $Settings
    )

    if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
        throw "Configuration template not found: $TemplatePath"
    }

    [xml]$xml = Get-Content -LiteralPath $TemplatePath -ErrorAction Stop
    if (-not $xml.ABConfiguration) {
        throw "Invalid configuration template: $TemplatePath"
    }

    $xml.ABConfiguration.DisplayName = $Settings.DisplayName
    $xml.ABConfiguration.ServerName = $Settings.LdapServer
    $xml.ABConfiguration.ServerPort = [string]$Settings.LdapPort
    $xml.ABConfiguration.UseSSL = ([bool]$Settings.UseSsl).ToString()

    # The address book authenticates as the signed-in user (SPA). The deployment
    # account and password must never be written into the client configuration.
    $xml.ABConfiguration.Username = ''
    $xml.ABConfiguration.Password = ''
    $xml.ABConfiguration.RequireSecurePasswordAuth = 'True'

    $parent = Split-Path -Parent $DestinationPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $xml.Save($DestinationPath)
}

function Get-ClientDeploymentFiles {
    @(
        'ABConfiguration.xml',
        'Deploy-LdapAB.ps1',
        'Remove-LdapAB.ps1',
        'LdapABManager_x64.exe',
        'LdapABManager_x86.exe'
    )
}

function New-DeploymentPackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        [Parameter(Mandatory = $true)]
        $Settings
    )

    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null

    New-RenderedConfiguration `
        -TemplatePath (Join-Path $SourcePath 'ABConfiguration.xml') `
        -DestinationPath (Join-Path $DestinationPath 'ABConfiguration.xml') `
        -Settings $Settings

    foreach ($file in (Get-ClientDeploymentFiles)) {
        if ($file -eq 'ABConfiguration.xml') { continue }

        $sourceFile = Join-Path $SourcePath $file
        if (-not (Test-Path -LiteralPath $sourceFile -PathType Leaf)) {
            throw "Required deployment file not found: $sourceFile"
        }

        Copy-Item -LiteralPath $sourceFile -Destination (Join-Path $DestinationPath $file) -Force
    }

    return $DestinationPath
}
