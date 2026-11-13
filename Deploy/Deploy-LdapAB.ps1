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

    # Check if already deployed (idempotency).
    $alreadyDeployed = $false
    if (Test-Path -LiteralPath $regPath) {
        $version = Get-ItemProperty -LiteralPath $regPath -Name Version -ErrorAction SilentlyContinue
        if ($version -and $version.Version -eq '1.0') {
            $alreadyDeployed = $true
        }
    }

    if ($alreadyDeployed) {
        Write-Log 'Already deployed, skipping'
        exit 0
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

    Write-Log "Using $exeName (64-bit: $use64)"
    Write-Log "Config: $ConfigPath"

    $managerArgs = @(
        '-m', 'create',
        '-f', $ConfigPath,
        '-pm', 'default'
    )

    if ($DryRun) {
        Write-Log "Dry run: would execute $exeName in create mode"
        Write-Log 'Dry run completed; no changes made'
        exit 0
    }

    Write-Log "Executing: $exePath $($managerArgs -join ' ')"
    $result = & $exePath $managerArgs 2>&1
    $exitCode = $LASTEXITCODE

    Write-Log "Exit code: $exitCode"
    Write-Log "Output: $result"

    if ($exitCode -ne 0) {
        Write-Log 'Deployment failed' 'ERROR'
        exit $exitCode
    }

    # Mark as deployed. A failure here does not undo a successful AB creation.
    try {
        New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $regPath -Name Version -Value '1.0' -PropertyType String -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $regPath -Name DeployedOn -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -PropertyType String -Force -ErrorAction Stop | Out-Null
    } catch {
        Write-Log 'Could not write registry marker' 'WARN'
    }

    Write-Log 'LDAP Address Book deployed successfully'
    exit 0
} catch {
    Write-Log "Exception: $_" 'ERROR'
    exit 1
}
