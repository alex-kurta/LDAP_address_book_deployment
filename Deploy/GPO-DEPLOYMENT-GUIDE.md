# GPO Deployment Guide: Outlook LDAP Address Book

## Overview

This package deploys an LDAP address book to Outlook profiles through a Group
Policy user logon script. Environment-specific values live in a local `.env`
file on the administrative workstation. The deployment package that is copied
to the domain controller contains only a rendered `ABConfiguration.xml`, the
client scripts, and the two LdapABManager executables — no `.env` and no
credentials.

## Prerequisites

- Windows PowerShell 5.1 or later
- Outlook installed on target clients
- RSAT Group Policy tools on the machine that creates the GPO
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
| `TARGET_OU` | Distinguished name of the OU to link the GPO to |

Optional keys:

- `NETLOGON_PATH` — defaults to `\\<DC_SERVER>\NETLOGON`
- `GPO_NAME` — defaults to `Outlook LDAP Address Book Deployment`
- `USERNAME` / `PASSWORD` — only needed when the current Windows account
  cannot write to NETLOGON. Leave blank to use the current account.

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

`Deploy-All.ps1` creates and links the GPO automatically when the GroupPolicy
module is available and `TARGET_OU` is set. Otherwise, run the dedicated script
on a domain controller or RSAT workstation:

```powershell
.\Deploy\Configure-LdapAB-GPO.ps1
```

Or create the GPO manually:

1. Open **Group Policy Management** (`gpmc.msc`).
2. Right-click the target OU → **Create a GPO in this domain, and Link it here**.
3. Name it (or use the value of `GPO_NAME`).
4. Edit the GPO → **User Configuration → Policies → Windows Settings → Scripts
   (Logon/Logoff)**.
5. Open **Logon** → **PowerShell Scripts** → **Add**.
6. Script name: `\\<DC_SERVER>\NETLOGON\LdapABDeploy\Deploy-LdapAB.ps1`.
7. Leave parameters empty; the script uses its own folder for the config and
   executables.

### Synchronous execution (recommended)

In the same GPO, enable:

**User Configuration → Administrative Templates → System → Scripts → Run logon
scripts synchronously**

This gives the Outlook profile time to exist before the script runs.

### Security filtering

The default **Authenticated Users** filter applies the GPO to everyone in the
linked OU. Restrict it to a pilot security group if needed.

## Verification Checklist

| Test | Method | Expected result |
|------|--------|-----------------|
| Script runs | `%TEMP%\LdapABDeploy.log` | Exit code 0, "deployed successfully" |
| AB created | Outlook → Address Book → Tools → Options | Configured display name is listed |
| Default AB | Open Address Book (`Ctrl+Shift+B`) | New address book appears first |
| Search works | Type a name in the **To:** field | Resolves from AD |
| Re-run idempotent | Log off and on again | Log shows "Already deployed, skipping" |

## Rollback

### Disable the GPO

Unlink the GPO from the OU or disable the link.

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
| 1.1 | 2026-11-12 | `.env`-driven configuration, safe smoke test, sanitized package |
| 1.0 | 2026-11-11 | Initial deployment package |
