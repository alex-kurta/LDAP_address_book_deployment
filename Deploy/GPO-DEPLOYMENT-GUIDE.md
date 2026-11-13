# GPO Deployment Guide: Outlook LDAP Address Book

## Overview

This package deploys an LDAP address book to Outlook profiles through a Group
Policy user logon script. Environment-specific values live in a local `.env`
file on the administrative workstation. The deployment package that is copied
to the domain controller contains only a rendered `ABConfiguration.xml`, the
client scripts, and the two LdapABManager executables — no `.env` and no
credentials.

Copying the client package to NETLOGON does **not** create or link the GPO. The
GPO must be created and linked separately, either from a workstation with RSAT
or remotely through WinRM.

## Current Status (reference deployment, verified 2026-11-13)

This section describes the validated reference deployment. Environment-specific
host names, domain names, GPO identifiers, and account names have been removed
from this public document.

| Item | Status |
|------|--------|
| `Outlook LDAP Address Book Deployment` GPO | **Created** and linked |
| GPO link | **Enabled** at the domain root (not enforced) |
| User logon script | `User\Scripts\psscripts.ini`, UTF-16LE with BOM |
| Scripts client-side extension | Registered in `gPCUserExtensionNames` |
| `Run logon scripts synchronously` | Enabled (`RunLogonScriptSync=1`) |
| GPO report | Includes `\\<DC_SERVER>\NETLOGON\LdapABDeploy\Deploy-LdapAB.ps1` |
| GPO applies on the reference workstation | **Yes** |
| Address book on the reference workstation | Deployed (`%TEMP%\LdapABDeploy.log`, registry marker present) |
| NETLOGON client package | Refreshed; matches the repository and `.env`, no credentials |
| `TARGET_OU` | Configured OU did not exist; domain-root linking was used instead |
| RSAT on the administrative workstation | Not installed |
| WinRM to the domain controller | Working with the deployment credentials |
| GroupPolicy module on the domain controller | Installed (Windows Server 2016; use `Get-GPInheritance`) |
| Secondary Samba DC | SMB session setup failed; SYSVOL mirror was empty |

### Known infrastructure limitation: Samba SYSVOL mirror

One reference-environment domain controller was a **Samba DC**, so Windows-only
tools such as `dcdiag`, `dfsrdiag`, WMI, and WinRM did not apply to it. Its
SYSVOL was populated by a **robocopy mirror** from the Windows DC, not by DFSR.

The Samba SMB service rejected session setup with **"The security ID structure
is invalid"** for both a standard user and a domain admin, so its SYSVOL and
NETLOGON shares were unreachable and the mirror was empty. The scheduled
robocopy task on the Windows DC had also been running since the previous
failure; because it was configured with `MultipleInstances=IgnoreNew`, later
retries were skipped.

The domain-based SYSVOL path resolved to the healthy Windows DC, so the GPO
applied on the reference workstation. Clients whose DFS referral landed on the
Samba DC could not read `gpt.ini` and would fail policy processing until the
Samba SMB service and SID mapping were repaired and the mirror was refreshed.
That infrastructure work is outside this deployment package.

## Prerequisites

- Windows PowerShell 5.1 or later
- Outlook installed on target clients
- One of these paths for GPO creation:
  - RSAT Group Policy tools on the workstation running the script, or
  - WinRM access to a domain controller that has the GroupPolicy module
- Write access to the domain controller's NETLOGON share
- A filled-in `.env` file (copy `.env.example`)

## 1. Configure .env

```powershell
Copy-Item .env.example .env
notepad .env
```

Set at least:

| Key | Purpose |
|-----|---------|
| `DOMAIN` | NetBIOS name or DNS name of the domain |
| `DC_SERVER` | Domain controller used for the NETLOGON copy |
| `LDAP_SERVER` | LDAP server host name or address |
| `LDAP_PORT` | LDAP port (389, or 636 with SSL) |
| `USE_SSL` | `true` or `false` |
| `LDAP_DISPLAY_NAME` | Address book name shown in Outlook |
| `TARGET_OU` | Existing OU distinguished name for pilot linking |

Optional keys:

- `NETLOGON_PATH` — defaults to `\\<DC_SERVER>\NETLOGON`
- `GPO_NAME` — defaults to `Outlook LDAP Address Book Deployment`
- `USERNAME` / `PASSWORD` — used for the NETLOGON SMB copy and the WinRM GPO
  connection. Leave blank to use the current Windows account.

For company-wide deployment, use `-LinkAtDomainRoot`. It derives the domain
distinguished name from `DOMAIN` or `TARGET_OU` and overrides the configured
`TARGET_OU` link scope.

`DOMAIN` may be a NetBIOS name. When it does not contain a dot, the scripts
derive the DNS domain from `TARGET_OU` (for example,
`OU=Users,DC=example,DC=com` becomes `example.com`). You can also set
`DOMAIN_FQDN` and `NETBIOS_DOMAIN` explicitly.

> Do not commit `.env`. It is gitignored.

## 2. Validate the configuration

Run the safe smoke test before deploying:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy\Test-Configuration.ps1
```

The test:

- loads `.env` and validates the required values without printing them
- derives and checks the domain and NETLOGON paths
- renders `ABConfiguration.xml` into a temporary directory
- checks that the rendered config uses Secure Password Authentication and has
  no embedded credentials
- runs `Deploy-LdapAB.ps1` and `Remove-LdapAB.ps1` in `-DryRun` mode
- runs `Deploy-All.ps1 -SkipConnectivityTest -SkipCopy` to validate package generation
- parses every PowerShell script for syntax errors
- verifies that `.env` is untracked and that tracked files contain no obvious
  token, IP-address, or password literals

The test does not modify Outlook or the registry.

## 3. Deploy the package

```powershell
.\Deploy\Deploy-All.ps1
```

Useful parameters:

```powershell
.\Deploy\Deploy-All.ps1 -EnvFile .\.env -SkipGPO
.\Deploy\Deploy-All.ps1 -DomainController dc01.example.com -TargetOU 'OU=Pilot,DC=example,DC=com'
.\Deploy\Deploy-All.ps1 -SkipConnectivityTest

# Validate .env, render the package, and skip all network/GPO changes:
.\Deploy\Deploy-All.ps1 -SkipConnectivityTest -SkipCopy
```

`Deploy-All.ps1` renders the client configuration into a temporary staging
folder and copies only these files to `\\<DC_SERVER>\NETLOGON\LdapABDeploy`:

```
ABConfiguration.xml
Deploy-LdapAB.ps1
Remove-LdapAB.ps1
LdapABManager_x86.exe
LdapABManager_x64.exe
```

The administrative scripts and `.env` are not copied.

If the GroupPolicy module is available locally, `Deploy-All.ps1` also creates,
links, and configures the GPO. If the module is not installed, use the remote
script in step 4 after the copy completes.

## ABConfiguration.xml Settings

`Deploy-LdapAB.ps1` reads this file on the client. The deployment script fills
in the display name, server, port, and SSL setting from `.env`:

```xml
<ABConfiguration>
  <DisplayName>Company Directory</DisplayName>
  <ServerName>ldap.example.com</ServerName>
  <ServerPort>389</ServerPort>
  <UseSSL>False</UseSSL>
  <Username></Username>
  <Password></Password>
  <RequireSecurePasswordAuth>True</RequireSecurePasswordAuth>
  <SearchTimeoutSeconds>60</SearchTimeoutSeconds>
  <MaxEntriesReturned>500</MaxEntriesReturned>
  <DefaultSearchBase>True</DefaultSearchBase>
  <CustomSearchBase></CustomSearchBase>
  <EnableBrowsing>True</EnableBrowsing>
</ABConfiguration>
```

Key settings:

- `RequireSecurePasswordAuth=True` uses the signed-in user's domain credentials.
- `Username` and `Password` stay empty; the deployment account is never written
  to the client configuration.
- A new address book is normally created at position #0 in Outlook's search
  order.

## 4. Create or link the GPO

Copying files to NETLOGON is not enough. The GPO must be created, linked, and
given the logon script. The GPO in this environment is already created and
linked; use this section to update it, recreate it, or change the link scope.

Choose the link scope first:

| Scope | Link target | Effect |
|-------|-------------|--------|
| Company-wide | Domain root (for example, `DC=example,DC=com`) | Applies to all users in the domain unless an OU blocks inheritance |
| Pilot | A real OU distinguished name | Applies only to users in that OU |

> In the reference environment, the configured `TARGET_OU` did not exist. A
> link command against a missing OU fails. Use `-LinkAtDomainRoot` for
> company-wide deployment, or pass an existing OU with `-TargetOU`.

### Option A: remote from a workstation without RSAT

`Configure-LdapAB-GPO-Remote.ps1` uses WinRM and the deployment credentials in
`.env` to run the GroupPolicy cmdlets on the domain controller. Preview first:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy\Configure-LdapAB-GPO-Remote.ps1 -DryRun -LinkAtDomainRoot
```

Then apply it company-wide:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy\Configure-LdapAB-GPO-Remote.ps1 -LinkAtDomainRoot
```

For a pilot OU instead:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy\Configure-LdapAB-GPO-Remote.ps1 -TargetOU 'OU=Pilot,DC=example,DC=com'
```

The script:

- creates the GPO if it does not exist
- links it to the domain root or the specified OU
- writes the user logon script to `User\Scripts\psscripts.ini` as UTF-16LE
  with a byte order mark (the Scripts CSE ignores other encodings and the old
  `Scripts\Logon` path)
- registers the Scripts client-side extension in the GPO
- enables **Run logon scripts synchronously**
- bumps the GPO user version and updates `gpt.ini` so clients process the change
- refuses to overwrite an existing GPO that already has other logon scripts

Rerunning the script is safe; it updates the existing GPO in place.

The deployment password is used only for the WinRM connection and is never
printed or written to the policy files.

### Option B: on a domain controller or RSAT workstation

```powershell
# Company-wide, linked at the domain root
.\Deploy\Configure-LdapAB-GPO.ps1 -LinkAtDomainRoot

# Pilot OU
.\Deploy\Configure-LdapAB-GPO.ps1 -TargetOU 'OU=Pilot,DC=example,DC=com'
```

`Deploy-All.ps1` also runs this configuration automatically when the
GroupPolicy module is available locally.

### Option C: manual GPMC steps

1. Open **Group Policy Management** (`gpmc.msc`).
2. Right-click the domain root (company-wide) or a pilot OU → **Create a GPO in
   this domain, and Link it here**.
3. Name it (or use the value of `GPO_NAME`).
4. Edit the GPO → **User Configuration → Policies → Windows Settings → Scripts
   (Logon/Logoff)**.
5. Open **Logon** → **PowerShell Scripts** → **Add**.
6. Script name: `\\<DC_SERVER>\NETLOGON\LdapABDeploy\Deploy-LdapAB.ps1`.
7. Leave parameters empty; the script uses its own folder for the config and
   executables.
8. Enable **User Configuration → Policies → Administrative Templates → System →
   Scripts → Run logon scripts synchronously**.

### Synchronous execution (recommended)

The automated scripts set the user policy
`HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\RunLogonScriptSync`
to `1`. This gives the Outlook profile time to exist before the script runs.
The same policy is available under **User Configuration** or **Computer
Configuration → Administrative Templates → System → Scripts**.

### Security filtering

The default **Authenticated Users** filter applies the GPO to everyone in the
linked scope. Restrict it to a pilot security group if needed.

## Verification Checklist

| Test | Method | Expected result |
|------|--------|-----------------|
| GPO exists | `Get-GPO -Name 'Outlook LDAP Address Book Deployment'` on the DC | GPO is returned |
| GPO linked | `Get-GPInheritance -Target 'DC=example,DC=com'` on the DC | GPO appears in `GpoLinks` |
| Logon script configured | Check `psscripts.ini` in the GPO's SYSVOL folder | Contains `Deploy-LdapAB.ps1` |
| Policy applied | `gpresult /r` on a client | GPO is listed under user policy |
| Script runs | `%TEMP%\LdapABDeploy.log` | Exit code 0, "deployed successfully" |
| AB created | Outlook → Address Book → Tools → Options | Configured display name is listed |
| Default AB | Open Address Book (`Ctrl+Shift+B`) | New address book appears first |
| Search works | Type a name in the **To:** field | Resolves from AD |
| Re-run idempotent | Log off and on again | Log shows "Already deployed, skipping" |

## Rollback

### Disable the GPO

Unlink the GPO from the domain root or OU, or disable the link:

```powershell
# Run on a domain controller or RSAT workstation
Remove-GPLink -Name 'Outlook LDAP Address Book Deployment' -Target 'DC=example,DC=com'
```

To remove the GPO entirely:

```powershell
Remove-GPO -Name 'Outlook LDAP Address Book Deployment'
```

### Run the removal script

```powershell
\\<DC_SERVER>\NETLOGON\LdapABDeploy\Remove-LdapAB.ps1
```

### Manual cleanup per user

```powershell
# Replace <DOMAIN> with the user's NetBIOS domain name.
Remove-Item "HKCU:\Software\<DOMAIN>\LdapABDeploy" -Recurse -Force
```

## Troubleshooting

### Script fails with "Outlook not found"

- Verify Outlook is installed.
- Check `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE`.

### Address book is created but not default

- The tool places a new address book at search index #0.
- If it is not first, check Outlook's "Choose automatically" setting:
  `HKCU\Software\Microsoft\Office\16.0\Outlook\Options\General\ABChooseDirAutomatically = 0`.

### Connection fails

- Verify the LDAP port is reachable from client VLANs to the LDAP server.
- Test with `ldp.exe` or `Test-NetConnection <LDAP_SERVER> -Port <LDAP_PORT>`.
- Check the Windows Firewall on the LDAP server.

### Duplicate address books on re-run

- The client script checks the registry marker
  `HKCU\Software\<DOMAIN>\LdapABDeploy`.
- LdapABManager also checks for an existing address book by display name.

### x86/x64 mismatch

- The client script auto-detects Outlook bitness and selects the matching
  executable.

## Requirements

- **LDAP server**: reachable on the configured port
- **Clients**: Outlook 2016/2019/2021/365 (x86 or x64)
- **Permissions**: domain users need read access to AD
- **PowerShell**: 3.0+ (5.1 recommended)

## Testing Before Full Rollout

1. Run `Test-Configuration.ps1`.
2. Test `Deploy-LdapAB.ps1` manually on one machine.
3. Verify the address book in Outlook and test a name lookup.
4. Link the GPO to a pilot OU first.
5. Monitor `%TEMP%\LdapABDeploy.log` on pilot machines.
6. Roll out to production OUs after the pilot succeeds.

## Log Locations

- **Deployment log**: `%TEMP%\LdapABDeploy.log`
- **Removal log**: `%TEMP%\LdapABRemove.log`
- **Registry marker**: `HKCU\Software\<DOMAIN>\LdapABDeploy`

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.3 | 2026-11-13 | GPO created and linked at the domain root; corrected `psscripts.ini` path, UTF-16LE encoding, Scripts CSE registration, and version handling; documented a Samba SYSVOL mirror limitation observed in the reference environment |
| 1.2 | 2026-11-12 | Corrected GPO status; added WinRM remote GPO script; fixed link checks; documented missing `TARGET_OU` |
| 1.1 | 2026-11-12 | `.env`-driven configuration, safe smoke test, sanitized package |
| 1.0 | 2026-11-11 | Initial deployment package |
