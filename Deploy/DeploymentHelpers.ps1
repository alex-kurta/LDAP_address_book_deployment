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

function Get-DistinguishedNameFromDomainFqdn {
    [CmdletBinding()]
    param(
        [string]$DomainFqdn
    )

    if ([string]::IsNullOrWhiteSpace($DomainFqdn)) { return $null }

    $parts = @($DomainFqdn.Split('.') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -eq 0) { return $null }

    return (($parts | ForEach-Object { "DC=$_" }) -join ',')
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

function Get-GpoSysvolPolicyPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GpoId,
        [Parameter(Mandatory = $true)]
        [string]$DomainFqdn,
        [string]$DomainController = ''
    )

    $policyFolder = "{$GpoId}"
    $candidates = @(
        (Join-Path $env:SystemRoot "SYSVOL\sysvol\$DomainFqdn\Policies\$policyFolder")
    )
    if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
        $candidates += "\\$DomainController\SYSVOL\$DomainFqdn\Policies\$policyFolder"
    }
    $candidates += "\\$DomainFqdn\SYSVOL\$DomainFqdn\Policies\$policyFolder"

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Container) {
            return $candidate
        }
    }

    throw "SYSVOL policy folder not found for GPO $GpoId. Checked: $($candidates -join '; ')"
}

function Get-LdapABScriptsExtensionName {
    [CmdletBinding()]
    param()

    # Scripts client-side extension plus the logon/logoff handler.
    return '[{42B5FAAE-6536-11D2-AE5A-0000F87571E3}{40B66650-4972-11D1-A7CA-0000F87571E3}]'
}

function Get-LdapABGpoDirectoryEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GpoId,
        [string]$DomainController = ''
    )

    if ([string]::IsNullOrWhiteSpace($DomainController)) {
        $rootDse = [ADSI]'LDAP://RootDSE'
    } else {
        $rootDse = [ADSI]"LDAP://$DomainController/RootDSE"
    }

    $domainDn = [string]$rootDse.defaultNamingContext
    if ([string]::IsNullOrWhiteSpace($domainDn)) {
        throw 'Cannot determine the domain distinguished name from RootDSE.'
    }

    if ([string]::IsNullOrWhiteSpace($DomainController)) {
        $gpoPath = "LDAP://CN={$GpoId},CN=Policies,CN=System,$domainDn"
    } else {
        $gpoPath = "LDAP://$DomainController/CN={$GpoId},CN=Policies,CN=System,$domainDn"
    }

    return [ADSI]$gpoPath
}

function Set-LdapABGpoScriptsIni {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PolicyPath,
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath
    )

    $scriptsDir = Join-Path $PolicyPath 'User\Scripts'
    New-Item -ItemType Directory -Path $scriptsDir -Force | Out-Null
    $iniPath = Join-Path $scriptsDir 'psscripts.ini'

    $otherScripts = @()
    if (Test-Path -LiteralPath $iniPath -PathType Leaf) {
        $existing = Get-Content -LiteralPath $iniPath -Raw -ErrorAction Stop
        foreach ($match in [regex]::Matches($existing, '(?im)^\s*\d+CmdLine\s*=\s*(.+?)\s*$')) {
            $cmdLine = $match.Groups[1].Value.Trim()
            if ($cmdLine -and $cmdLine -ne $ScriptPath) {
                $otherScripts += $cmdLine
            }
        }
        if ($otherScripts.Count -gt 0) {
            throw "psscripts.ini already contains other logon scripts ($($otherScripts -join '; ')): $iniPath. Refusing to overwrite them."
        }
    }

    # The Scripts CSE reads User\Scripts\psscripts.ini as UTF-16LE with a byte
    # order mark. The file is ignored when it is stored under Scripts\Logon or
    # written as ASCII.
    $iniContent = "[Logon]`r`n0CmdLine=$ScriptPath`r`n0Parameters="
    Set-Content -LiteralPath $iniPath -Value $iniContent -Encoding Unicode -NoNewline -ErrorAction Stop

    # Older versions of this toolkit used the wrong folder. Remove that file
    # when it only points at our script so it cannot confuse later diagnostics.
    $legacyIni = Join-Path $scriptsDir 'Logon\psscripts.ini'
    $legacyRemoved = $false
    if (Test-Path -LiteralPath $legacyIni -PathType Leaf) {
        $legacy = Get-Content -LiteralPath $legacyIni -Raw -ErrorAction Stop
        if ($legacy -match [regex]::Escape($ScriptPath)) {
            Remove-Item -LiteralPath $legacyIni -Force -ErrorAction Stop
            $legacyRemoved = $true
        }
    }

    return [pscustomobject]@{
        ScriptIni     = $iniPath
        LegacyIniPath = $legacyIni
        LegacyRemoved = $legacyRemoved
    }
}

function Register-LdapABGpoScriptsExtension {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GpoId,
        [string]$DomainController = ''
    )

    $extension = Get-LdapABScriptsExtensionName
    $entry = Get-LdapABGpoDirectoryEntry -GpoId $GpoId -DomainController $DomainController
    $current = [string]$entry.Properties['gPCUserExtensionNames'].Value

    if ($current -match '42B5FAAE-6536-11D2-AE5A-0000F87571E3') {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($current)) {
        $entry.Properties['gPCUserExtensionNames'].Value = $extension
    } else {
        $entry.Properties['gPCUserExtensionNames'].Value = $current + $extension
    }
    $entry.CommitChanges()

    return $true
}

function Update-LdapABGpoVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GpoId,
        [Parameter(Mandatory = $true)]
        [string]$PolicyPath,
        [string]$GPOName = '',
        [string]$DomainController = ''
    )

    $entry = Get-LdapABGpoDirectoryEntry -GpoId $GpoId -DomainController $DomainController
    $current = [int]$entry.Properties['versionNumber'].Value
    $userVersion = [int][math]::Floor($current / 65536)
    $computerVersion = $current % 65536
    $userVersion = ($userVersion + 1) % 65536
    $newVersion = ($userVersion * 65536) + $computerVersion

    $entry.Properties['versionNumber'].Value = $newVersion
    $entry.CommitChanges()

    # The Version in gpt.ini must match versionNumber, or clients skip the GPO.
    $gptIniPath = Join-Path $PolicyPath 'gpt.ini'
    if (Test-Path -LiteralPath $gptIniPath -PathType Leaf) {
        $content = Get-Content -LiteralPath $gptIniPath -Raw -ErrorAction Stop
        if ($content -match '(?im)^\s*Version\s*=') {
            $content = [regex]::Replace($content, '(?im)^\s*Version\s*=.*$', "Version=$newVersion")
        } else {
            $content = $content.TrimEnd() + "`r`nVersion=$newVersion`r`n"
        }
    } else {
        if ([string]::IsNullOrWhiteSpace($GPOName)) {
            $GPOName = 'Outlook LDAP Address Book Deployment'
        }
        $content = "[General]`r`nVersion=$newVersion`r`ndisplayName=$GPOName`r`n"
    }
    Set-Content -LiteralPath $gptIniPath -Value $content -Encoding ASCII -NoNewline -ErrorAction Stop

    return $newVersion
}

function Invoke-LdapABGpoConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GPOName,
        [Parameter(Mandatory = $true)]
        [string]$TargetOU,
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,
        [Parameter(Mandatory = $true)]
        [string]$DomainFqdn,
        [string]$DomainController = '',
        [switch]$SkipSynchronousScripts
    )

    Import-Module GroupPolicy -ErrorAction Stop

    $created = $false
    $gpo = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
    if (-not $gpo) {
        $gpo = New-GPO -Name $GPOName -Comment 'Deploy LDAP Address Book to Outlook via MAPI'
        $created = $true
    }

    $linked = $false
    try {
        $inheritance = Get-GPInheritance -Target $TargetOU -ErrorAction Stop
    } catch {
        throw "Cannot read GPO links for target '$TargetOU'. Verify that the OU or domain distinguished name exists. $($_.Exception.Message)"
    }

    $existingLink = @($inheritance.GpoLinks) |
        Where-Object { [string]$_.GpoId -eq [string]$gpo.Id } |
        Select-Object -First 1
    if (-not $existingLink) {
        New-GPLink -Name $GPOName -Target $TargetOU -ErrorAction Stop | Out-Null
        $linked = $true
    }

    $policyPath = Get-GpoSysvolPolicyPath `
        -GpoId ([string]$gpo.Id) `
        -DomainFqdn $DomainFqdn `
        -DomainController $DomainController

    $scriptsResult = Set-LdapABGpoScriptsIni `
        -PolicyPath $policyPath `
        -ScriptPath $ScriptPath

    $scriptsExtensionRegistered = Register-LdapABGpoScriptsExtension `
        -GpoId ([string]$gpo.Id) `
        -DomainController $DomainController

    $syncPolicySet = $false
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
        $syncPolicySet = $true
    }

    # Bump the user version after the script and registry changes are written
    # so clients are guaranteed to process them on their next refresh.
    $userVersion = Update-LdapABGpoVersion `
        -GpoId ([string]$gpo.Id) `
        -PolicyPath $policyPath `
        -GPOName $GPOName `
        -DomainController $DomainController

    return [pscustomobject]@{
        GPOName                    = $GPOName
        GpoId                      = [string]$gpo.Id
        Created                    = $created
        Linked                     = $linked
        LinkTarget                 = $TargetOU
        ScriptIni                  = $scriptsResult.ScriptIni
        LegacyScriptIniRemoved     = $scriptsResult.LegacyRemoved
        ScriptsExtensionRegistered = $scriptsExtensionRegistered
        SyncPolicySet              = $syncPolicySet
        UserVersion                = $userVersion
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
        'Run-LdapAB.cmd',
        'LdapABManager_x64.exe',
        'LdapABManager_x86.exe',
        'SetLdapABDefault_x64.exe',
        'SetLdapABDefault_x86.exe'
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
