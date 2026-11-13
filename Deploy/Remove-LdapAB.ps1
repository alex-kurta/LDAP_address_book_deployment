<#
.SYNOPSIS
    Removes the LDAP Address Book from the current user's Outlook profile.

.DESCRIPTION
    Client-side rollback script. It cleans up the deployment registry marker
    and does not contain a hardcoded domain name.

.PARAMETER ConfigPath
    Path to the rendered ABConfiguration.xml.

.PARAMETER LogPath
    Path to the removal log.

.PARAMETER RegistryRoot
    Overrides the registry root used for the idempotency marker. Intended for
    testing. Defaults to the current NetBIOS domain name.

.PARAMETER DryRun
    Validates prerequisites and reports what would happen without changing
    Outlook or the registry.
#>

param(
    [string]$ConfigPath = "$PSScriptRoot\ABConfiguration.xml",
    [string]$LogPath = "$env:TEMP\LdapABRemove.log",
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
    Write-Log 'Starting LDAP AB removal'

    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        throw "Config not found: $ConfigPath"
    }

    # Detect Outlook bitness.
    $outlookPath = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' -ErrorAction SilentlyContinue).'(default)'
    if (-not $outlookPath -or -not (Test-Path -LiteralPath $outlookPath)) {
        $outlookPath = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' -ErrorAction SilentlyContinue).'(default)'
    }

    $use64 = $false
    if ($outlookPath -and (Test-Path -LiteralPath $outlookPath)) {
        try {
            $use64 = [System.Reflection.AssemblyName]::GetAssemblyName($outlookPath).ProcessorArchitecture -eq 'Amd64'
        } catch {
            $outlookProc = Get-Process OUTLOOK -ErrorAction SilentlyContinue
            if ($outlookProc) {
                $use64 = $outlookProc[0].StartInfo.EnvironmentVariables['PROCESSOR_ARCHITECTURE'] -eq 'AMD64'
            } else {
                $use64 = [Environment]::Is64BitOperatingSystem
            }
        }
    } else {
        $use64 = [Environment]::Is64BitOperatingSystem
    }

    $exeName = if ($use64) { 'LdapABManager_x64.exe' } else { 'LdapABManager_x86.exe' }
    $exePath = Join-Path $PSScriptRoot $exeName

    if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) {
        throw "Executable not found: $exePath"
    }

    # Read the display name from the rendered configuration.
    [xml]$xml = Get-Content -LiteralPath $ConfigPath -ErrorAction Stop
    if (-not $xml.ABConfiguration) {
        throw "Config is missing the ABConfiguration root element: $ConfigPath"
    }
    $displayName = $xml.ABConfiguration.DisplayName

    Write-Log "Using $exeName (64-bit: $use64)"
    Write-Log "Removing AB: $displayName"

    $managerArgs = @(
        '-m', 'remove',
        '-dn', $displayName,
        '-pm', 'default'
    )

    if ($DryRun) {
        Write-Log "Dry run: would execute $exeName in remove mode"
        Write-Log 'Dry run completed; no changes made'
        exit 0
    }

    Write-Log "Executing: $exePath $($managerArgs -join ' ')"
    $result = & $exePath $managerArgs 2>&1
    $exitCode = $LASTEXITCODE

    Write-Log "Exit code: $exitCode"
    Write-Log "Output: $result"

    # Remove the registry marker.
    if (Test-Path -LiteralPath $regPath) {
        Remove-Item -LiteralPath $regPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log 'Removed deployment registry marker'
    }

    if ($exitCode -ne 0) {
        Write-Log 'Removal failed' 'ERROR'
        exit $exitCode
    }

    Write-Log 'LDAP Address Book removed successfully'
    exit 0
} catch {
    Write-Log "Exception: $_" 'ERROR'
    exit 1
}
