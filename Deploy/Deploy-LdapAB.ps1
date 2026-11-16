<#
.SYNOPSIS
    Deploys the LDAP Address Book to the current user's default Outlook profile.

.DESCRIPTION
    Client-side logon script. The deployment package contains a pre-rendered
    ABConfiguration.xml, so this script does not read .env and contains no
    environment-specific credentials.

.PARAMETER ConfigPath
    Path to the rendered ABConfiguration.xml.

.PARAMETER LogPath
    Path to the deployment log.

.PARAMETER RegistryRoot
    Overrides the registry root used for the idempotency marker. Intended for
    testing. Defaults to the current NetBIOS domain name.

.PARAMETER DryRun
    Validates prerequisites and reports what would happen without changing
    Outlook or the registry.
#>

param(
    [string]$ConfigPath = "$PSScriptRoot\ABConfiguration.xml",
    [string]$LogPath = "$env:TEMP\LdapABDeploy.log",
    [string]$RegistryRoot = "",
    [switch]$DryRun
)

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$timestamp [$Level] $Message" | Tee-Object -FilePath $LogPath -Append
}

# Build the idempotency marker path without hardcoding a domain name.
if (-not [string]::IsNullOrWhiteSpace($RegistryRoot)) {
    $regPath = "Registry::HKEY_CURRENT_USER\Software\$RegistryRoot\LdapABDeploy"
} elseif ($env:USERDOMAIN) {
    $regPath = "Registry::HKEY_CURRENT_USER\Software\$($env:USERDOMAIN)\LdapABDeploy"
} elseif ($env:USERDNSDOMAIN) {
    $regPath = "Registry::HKEY_CURRENT_USER\Software\$($env:USERDNSDOMAIN)\LdapABDeploy"
} else {
    $regPath = 'Registry::HKEY_CURRENT_USER\Software\LdapABDeploy'
}

try {
    Write-Log 'Starting LDAP AB deployment'

    # Detect Outlook bitness.
    $outlookPath = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' -ErrorAction SilentlyContinue).'(default)'
    if (-not $outlookPath -or -not (Test-Path -LiteralPath $outlookPath)) {
        $outlookPath = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' -ErrorAction SilentlyContinue).'(default)'
    }

    # Office Click-to-Run records the platform explicitly. If that is not
    # available, inspect the Outlook executable's PE header.
    function Test-OutlookIs64Bit {
        param([string]$Path)

        try {
            $ctr = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop
            if ($ctr.Platform -eq 'x64') { return $true }
            if ($ctr.Platform -eq 'x86') { return $false }
        } catch {}

        try {
            $office = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook' -ErrorAction Stop
            if ($office.Bitness -eq 'x64') { return $true }
            if ($office.Bitness -eq 'x86') { return $false }
        } catch {}

        try {
            $office32 = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Outlook' -ErrorAction Stop
            if ($office32.Bitness -eq 'x64') { return $true }
            if ($office32.Bitness -eq 'x86') { return $false }
        } catch {}

        if ($Path -and (Test-Path -LiteralPath $Path)) {
            try {
                $stream = [System.IO.File]::OpenRead($Path)
                try {
                    $reader = New-Object System.IO.BinaryReader($stream)
                    $stream.Position = 0x3C
                    $peOffset = $reader.ReadInt32()
                    $stream.Position = $peOffset + 4
                    $machine = $reader.ReadUInt16()
                } finally {
                    $stream.Close()
                }
                if ($machine -eq 0x8664) { return $true }
                if ($machine -eq 0x014C) { return $false }
            } catch {}
        }

        return [Environment]::Is64BitOperatingSystem
    }

    $use64 = Test-OutlookIs64Bit -Path $outlookPath

    $exeName = if ($use64) { 'LdapABManager_x64.exe' } else { 'LdapABManager_x86.exe' }
    $exePath = Join-Path $PSScriptRoot $exeName
    $helperName = if ($use64) { 'SetLdapABDefault_x64.exe' } else { 'SetLdapABDefault_x86.exe' }
    $helperPath = Join-Path $PSScriptRoot $helperName

    if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) {
        throw "Executable not found: $exePath"
    }
    if (-not (Test-Path -LiteralPath $helperPath -PathType Leaf)) {
        throw "Default-address-book helper not found: $helperPath"
    }
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        throw "Config not found: $ConfigPath"
    }

    # Validate the configuration before invoking the MAPI tool.
    try {
        [xml]$configXml = Get-Content -LiteralPath $ConfigPath -ErrorAction Stop
    } catch {
        throw "Config is not valid XML: $ConfigPath"
    }
    if (-not $configXml.ABConfiguration) {
        throw "Config is missing the ABConfiguration root element: $ConfigPath"
    }

    $displayName = [string]$configXml.ABConfiguration.DisplayName
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $displayName = 'Company Directory'
    }

    Write-Log "Using $exeName and $helperName (64-bit: $use64)"
    Write-Log "Config: $ConfigPath"

    # Determine whether the address book already exists. This avoids running
    # create on every scheduled-task pass while still repairing a profile that
    # is missing the address book.
    $addressBookExists = $false
    $listResult = & $exePath -m list 2>&1
    $listExit = $LASTEXITCODE
    $listText = ($listResult | Out-String)
    if ($listExit -eq 0 -and $listText -match [regex]::Escape($displayName)) {
        $addressBookExists = $true
    }

    if ($DryRun) {
        if ($addressBookExists) {
            Write-Log "Dry run: '$displayName' already exists; would verify the default address book"
        } else {
            Write-Log "Dry run: would execute $exeName in create mode, then set '$displayName' as default"
        }
        Write-Log 'Dry run completed; no changes made'
        exit 0
    }

    if (-not $addressBookExists) {
        $managerArgs = @(
            '-m', 'create',
            '-f', $ConfigPath,
            '-pm', 'default'
        )

        Write-Log "Executing: $exePath $($managerArgs -join ' ')"
        $result = & $exePath $managerArgs 2>&1
        $exitCode = $LASTEXITCODE
        $resultText = ($result | Out-String)

        Write-Log "Exit code: $exitCode"
        Write-Log "Output: $resultText"

        if ($exitCode -ne 0 -or $resultText -match 'FAILED!') {
            Write-Log 'Deployment failed' 'ERROR'
            if ($exitCode -ne 0) { exit $exitCode }
            exit 1
        }

        # The MAPI tool returns success only if the profile now contains the
        # address book. Confirm it before continuing.
        $listResult = & $exePath -m list 2>&1
        $listText = ($listResult | Out-String)
        Write-Log "List output: $listText"
        if ($listText -notmatch [regex]::Escape($displayName)) {
            Write-Log "Address book '$displayName' was not found after create" 'ERROR'
            exit 1
        }
    } else {
        Write-Log "Address book '$displayName' already exists"
    }

    # Make the LDAP address list the default directory and put it first in the
    # name-resolution search path. This is idempotent and repairs profiles that
    # already had the address book but never selected it as the default.
    $helperArgs = @('-set', $displayName)
    Write-Log "Executing: $helperPath $($helperArgs -join ' ')"
    $helperResult = & $helperPath $helperArgs 2>&1
    $helperExit = $LASTEXITCODE
    $helperText = ($helperResult | Out-String)

    Write-Log "Helper exit code: $helperExit"
    Write-Log "Helper output: $helperText"

    if ($helperExit -ne 0 -or $helperText -notmatch 'VERIFIED default and search order') {
        Write-Log "Could not set '$displayName' as the default address book" 'ERROR'
        if ($helperExit -ne 0) { exit $helperExit }
        exit 1
    }

    # Mark as deployed. A failure here does not undo a successful change.
    try {
        New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $regPath -Name Version -Value '2.0' -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $regPath -Name DefaultAddressBook -Value $displayName -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $regPath -Name DeployedOn -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $regPath -Name DefaultSetOn -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -PropertyType String -Force -ErrorAction Stop | Out-Null
    } catch {
        Write-Log 'Could not write registry marker' 'WARN'
    }

    Write-Log 'LDAP Address Book deployed and set as default'
    exit 0
} catch {
    Write-Log "Exception: $_" 'ERROR'
    exit 1
}
