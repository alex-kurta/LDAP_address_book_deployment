# Complete Process Documentation: Outlook LDAP Address Book GPO Deployment

## Overview

This document describes how an LDAP address book is deployed from Active
Directory to Outlook clients through Group Policy. Environment-specific values
are stored in a local `.env` file and are never committed to the repository.

## Reference Deployment Status (verified 2026-11-13)

The deployment was completed and validated in a reference Active Directory
environment. Environment-specific host names, domain names, GPO identifiers,
and account names have been removed from this public document.

| Item | Status |
|------|--------|
| `Outlook LDAP Address Book Deployment` GPO | Created and linked at the domain root (not enforced) |
| User logon script | `User\Scripts\psscripts.ini`, UTF-16LE with BOM |
| Scripts CSE | Registered; `RunLogonScriptSync=1` |
| NETLOGON `ABConfiguration.xml` | Matches `.env`, no credentials |
| NETLOGON `Deploy-LdapAB.ps1` / `Remove-LdapAB.ps1` | Match the repository |
| GPO applies on the reference workstation | Yes (2026-11-13) |
| Address book on the reference workstation | Deployed (`%TEMP%\LdapABDeploy.log`, registry marker present) |
| `TARGET_OU` | Configured OU did not exist; domain root used instead |
| RSAT on the administrative workstation | Not installed |
| WinRM to the domain controller | Working with the deployment credentials |
| GroupPolicy module on the domain controller | Installed (Server 2016; has `Get-GPInheritance`, no `Get-GPLink`) |
| Secondary Samba DC | SMB session setup failed; SYSVOL mirror was empty (superseded — see "Addendum: Production hardening (2026-11-16)") |

The GPO applied on the Windows DC and on the reference workstation. However,
clients whose DFS referral for the domain SYSVOL path landed on the secondary
Samba DC could not read `gpt.ini` and would fail policy processing until that
controller's SMB service and SYSVOL mirror were repaired. See "Known
Limitations".

## Problem Statement

- Users run Outlook 2016/2019/2021 or Microsoft 365.
- The LDAP address book is not loaded automatically for domain users.
- Manual setup requires entering the LDAP server and authentication details.
- The address book must be added programmatically and appear first in the
  Outlook address book list.

## Solution

Use LdapABManager (a MAPI tool) deployed through a GPO user logon script to add
the LDAP address book to each user's Outlook profile. The client configuration
uses Secure Password Authentication (SPA), so the signed-in user's domain
credentials are used and no password is stored on the client.

## Process Steps

### 1. Discovery and Analysis

**Project location:** the repository root.

**Key files:**

- `LdapABManager_Code/` — C++ source code (based on the Microsoft MAPI sample)
- `LdapABManager_Exe/` — prebuilt 32-bit and 64-bit executables
- `LdapABManager_ABConfiguration.xml` — sanitized configuration template
- `Deploy/` — deployment package and administrative scripts

**Analysis:**

- The tool uses MAPI to manage LDAP address books in Outlook profiles.
- It reads an XML configuration and creates, updates, or removes the LDAP
  service in the profile.
- Passwords, when supplied, are encrypted with the Windows CryptoAPI.
- Authentication is performed through the current MAPI profile.

### 2. Configuration Approach

The client-side configuration uses SPA:

```xml
<DisplayName>Company Directory</DisplayName>
<ServerName>ldap.example.com</ServerName>
<ServerPort>389</ServerPort>
<UseSSL>False</UseSSL>
<Username></Username>
<Password></Password>
<RequireSecurePasswordAuth>True</RequireSecurePasswordAuth>
```

**Rationale:** the signed-in domain user's credentials are used instead of a
shared administrator account. No password is stored in a file, script, or
registry value.

The deployment values themselves are supplied through `.env`:

- `DOMAIN`, `DC_SERVER`, `NETLOGON_PATH`
- `LDAP_SERVER`, `LDAP_PORT`, `USE_SSL`, `LDAP_DISPLAY_NAME`
- `TARGET_OU`, `GPO_NAME`
- `USERNAME` / `PASSWORD` (only for the NETLOGON SMB copy, if needed)

Copy `.env.example` to `.env` and fill it in. Never commit `.env`.

### 3. Source Code Enhancements

**Modified files:**

- `LdapABManager_Code/LdapABManager.h` — added the `nSearchIndex` field
- `LdapABManager_Code/LdapABManager.cpp` — added `-si` parameter parsing and
  search-order logic
- `LdapABManager_Code/ProfileFunctions.h` — added
  `HrGetABSearchOrderRowCount` and `HrSetABSearchOrder`
- `LdapABManager_Code/ProfileFunctions.cpp` — implemented search-order
  manipulation through `IAddrBook`

**Purpose:** set the LDAP address book as the default (search index 1 = first
position). Testing showed new address books normally appear at position #0
already, so this is a completeness feature rather than a requirement.

### 4. Deployment Package

**Created:** `Deploy/` with:

- `LdapABManager_x86.exe` / `LdapABManager_x64.exe`
- `ABConfiguration.xml` — generic template, rendered at deploy time
- `Deploy-LdapAB.ps1` — client logon script
- `Remove-LdapAB.ps1` — rollback script
- `Configure-LdapAB-GPO.ps1` — local GPO automation (DC or RSAT)
- `Configure-LdapAB-GPO-Remote.ps1` — GPO automation through WinRM
- `Deploy-All.ps1` — NETLOGON package deployment automation
- `DeploymentHelpers.ps1` — shared `.env` and configuration helpers
- `Test-Configuration.ps1` — safe smoke test
- `GPO-DEPLOYMENT-GUIDE.md` — manual instructions

**Deploy-LdapAB.ps1 features:**

- Idempotency check through
  `HKCU\Software\<DOMAIN>\LdapABDeploy` (the domain comes from the current
  user's environment, not a hardcoded value)
- Auto-detects Outlook bitness (executable architecture, running process, or OS)
- Uses the matching x86/x64 LdapABManager executable
- Calls `LdapABManager.exe -m create -f ABConfiguration.xml -pm default`
- Creates a registry marker on success
- Logs to `%TEMP%\LdapABDeploy.log`
- Supports `-DryRun` for validation without changing Outlook or the registry

### 5. Local Testing

Run the safe smoke test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy\Test-Configuration.ps1
```

The test loads `.env`, validates required values, renders the client config to
a temporary folder, dry-runs the deployment and removal scripts, parses all
PowerShell files, and checks repository hygiene. It does not modify Outlook or
the registry.

For a manual pilot:

1. Run `Deploy-LdapAB.ps1` on a test workstation.
2. Confirm the address book appears in Outlook.
3. Run the script again and confirm it logs "Already deployed, skipping".
4. Run `Remove-LdapAB.ps1` and confirm the address book is removed.

### 6. Git Repository Setup

**Repository:** local Git working tree with `origin` pointing to the project's
GitHub repository.

**Ignored content:**

- `.env` and other local environment files
- `.kilo/` agent state
- Build outputs and local staging folders

**Important:** the deployment account password is never committed. If a secret
was ever committed, rotate it and rewrite history before pushing.

### 7. Deployment to the Domain Controller

**Target:** NETLOGON share on the domain controller, specifically
`\\<DC_SERVER>\NETLOGON\LdapABDeploy`.

**Process:**

1. Run `.\Deploy\Deploy-All.ps1`.
2. The script renders `ABConfiguration.xml` in a temporary staging folder.
3. It copies only the client files to NETLOGON.
4. If the current account cannot write to NETLOGON, supply `-AdminUser` and
   `-AdminPassword` or the corresponding `.env` values so the script can create
   a temporary SMB mapping.
5. The mapping is removed when the copy finishes.

**Files on the domain controller:**

```
ABConfiguration.xml
Deploy-LdapAB.ps1
Remove-LdapAB.ps1
LdapABManager_x64.exe
LdapABManager_x86.exe
```

### 8. GPO Configuration

The GPO is a separate step from the NETLOGON copy. In this environment it is
already created and linked; the commands below update it in place or recreate
it if necessary.

**Choose the link scope:**

- **Company-wide:** link at the domain root (for example,
  `DC=example,DC=com`). This applies to all domain users unless an OU blocks
  inheritance.
- **Pilot:** link to a real OU distinguished name.
- In the reference environment, the configured `TARGET_OU` did not exist, so
  domain-root linking was used instead.

**Remote automation (no RSAT on the administrative workstation):**

```powershell
# Preview
.\Deploy\Configure-LdapAB-GPO-Remote.ps1 -DryRun -LinkAtDomainRoot

# Apply company-wide
.\Deploy\Configure-LdapAB-GPO-Remote.ps1 -LinkAtDomainRoot

# Or link to a real pilot OU
.\Deploy\Configure-LdapAB-GPO-Remote.ps1 -TargetOU 'OU=Pilot,DC=example,DC=com'
```

The script connects to the domain controller through WinRM with the deployment
credentials from `.env`. It creates the GPO if needed, links it, writes the
user logon script to `User\Scripts\psscripts.ini` (UTF-16LE with BOM),
registers the Scripts client-side extension, enables synchronous logon
scripts, and bumps the GPO version so clients reprocess the policy.

**Local automation (DC or RSAT workstation):**

```powershell
.\Deploy\Configure-LdapAB-GPO.ps1 -LinkAtDomainRoot
```

`Deploy-All.ps1` runs the local automation automatically only when the
GroupPolicy module is installed on the workstation. It does not create the GPO
from a machine without RSAT; use the remote script for that.

**Manual steps via GPMC:**

1. Create a GPO and link it to the domain root or a pilot OU.
2. Edit the GPO → **User Configuration → Policies → Windows Settings → Scripts
   (Logon/Logoff)**.
3. Add a PowerShell logon script:
   `\\<DC_SERVER>\NETLOGON\LdapABDeploy\Deploy-LdapAB.ps1`.
4. Enable **User Configuration → Administrative Templates → System → Scripts →
   Run logon scripts synchronously**.
5. Optionally restrict security filtering to a pilot group.

## Key Technical Details

### Authentication

- **Method:** Secure Password Authentication (SPA/NTLM/Kerberos)
- **Config:** `RequireSecurePasswordAuth=True`, empty username and password
- **Benefit:** uses the signed-in user's domain credentials; no passwords are
  stored in files or scripts

### Default Address Book

- LdapABManager creates a new address book at position #0 in the MAPI search
  order by default.
- It appears first when the user opens the Outlook address book
  (`Ctrl+Shift+B`).
- No additional configuration is required.

### Compatibility

- **Outlook:** 2016, 2019, 2021, Microsoft 365
- **Architecture:** x86 and x64 auto-detected
- **OS:** Windows with MAPI installed
- **Protocol:** LDAP on port 389, or LDAPS on port 636 with `USE_SSL=true`

### Idempotency

- Registry marker: `HKCU\Software\<DOMAIN>\LdapABDeploy` with `Version=2.0`
- LdapABManager checks for an existing address book by display name
- Safe to run multiple times

## File Structure

```
<repository root>
├── .env.example                  # Template; copy to .env
├── .gitignore                    # Ignores .env and build output
├── DEPLOYMENT_PROCESS.md         # This document
├── LdapABManager_ABConfiguration.xml
├── LdapABManager_Code/           # C++ source
├── LdapABManager_Exe/            # Prebuilt executables
└── Deploy/
    ├── ABConfiguration.xml       # Generic client template
    ├── Configure-LdapAB-GPO.ps1
    ├── Configure-LdapAB-GPO-Remote.ps1
    ├── Deploy-All.ps1
    ├── Deploy-LdapAB.ps1
    ├── DeploymentHelpers.ps1
    ├── GPO-DEPLOYMENT-GUIDE.md
    ├── LdapABManager_x64.exe
    ├── LdapABManager_x86.exe
    ├── Remove-LdapAB.ps1
    └── Test-Configuration.ps1
```

## Known Limitations

- In the reference environment, a secondary Samba DC populated its SYSVOL with
  a robocopy mirror from the Windows DC rather than DFSR. On 2026-11-13 its
  SMB service rejected session setup with "The security ID structure is
  invalid" (for both standard users and domain admins), so its SYSVOL/NETLOGON
  shares were unreachable. That failure is **resolved as of 2026-11-16** —
  see "Addendum: Production hardening (2026-11-16)" for the current state.
- GPO creation needs either the GroupPolicy module locally or WinRM access to a
  domain controller that has it.
- `Get-GPLink` is not present in the Windows Server 2016 GroupPolicy module.
  The scripts use `Get-GPInheritance` instead.
- The Scripts CSE requires `User\Scripts\psscripts.ini` encoded as UTF-16LE
  with a BOM. ASCII files and the `Scripts\Logon` path are ignored.
- NETLOGON write access may require explicit credentials.
- The deployment package must be re-rendered when `.env` changes.
- Rewriting Git history is required if a secret was committed before it was
  removed.

## Next Steps

1. Fill in `.env` and run `Deploy\Test-Configuration.ps1`.
2. Run `Deploy\Deploy-All.ps1` to refresh the NETLOGON package.
3. Run `Deploy\Configure-LdapAB-GPO-Remote.ps1 -DryRun`, then apply it without
   `-DryRun` to create or update the GPO.
4. Verify the GPO link, `User\Scripts\psscripts.ini` encoding, Scripts CSE
   registration, and that the AD `versionNumber` matches `gpt.ini` (kept at
   `393221` since 2026-11-16; Samba and Windows DC agree).
5. Optionally pilot any *new* configuration change on a test machine first;
   the current deployment is live in production and monitoring
   `%TEMP%\LdapABDeploy.log` is the primary health signal.
6. Optionally enable LDAPS (port 636) with a valid certificate.

## Addendum: Production hardening (2026-11-16)

The deployment was running but employees reported a console window flashing
every 15 minutes. The scheduled-task preference that drives the deployment was
hardened and re-targeted. These changes live in the deployed GPO preference
item, not in this repository; environment-specific identifiers are not
published here.

### What changed

1. **Hidden task launcher.** The task action changed from
   `cmd.exe /c "\\<DOMAIN>\NETLOGON\LdapABDeploy\Run-LdapAB.cmd"` to
   `C:\Windows\System32\conhost.exe --headless cmd.exe /c "...\Run-LdapAB.cmd"`.
   Verified on a Windows 11 build 26200 client: both `conhost.exe` and
   `cmd.exe` reported `MainWindowHandle=0` and an empty window title, and the
   deployment still completed with `VERIFIED default and search order`.
2. **OS targeting.** The preference item now carries item-level targeting: a
   WMI filter `SELECT * FROM Win32_OperatingSystem WHERE Version >= '10.0.17763'`
   (Windows 10 1809+ and Windows 11). The item is also configured to remove
   itself where the filter stops matching. Windows 7/XP clients and the
   Windows Server 2016 DC receive no task and show no window.
3. **GPO version.** `GPT.INI` bumped from `393220` to `393221`, the Scripts /
   Preferences CSE registrations were unchanged, and the AD `versionNumber`
   was updated to match via `ldbmodify`/`Get-ADObject` on the Samba and
   Windows DCs respectively. `samba-tool ntacl sysvolcheck` passes.
4. **Samba SYSVOL mirror.** The Windows-to-Samba robocopy mirror task was run
   and its outcome (`lastResult=1`, i.e. files copied) confirmed; SYSVOL hashes
   for the GPO now match on both DCs.
5. **Force-refresh of a stale client.** One reported PC still ran
   the old `cmd.exe` action until a `gpupdate /target:user /force` (issued as
   the affected user) replaced the local task with the headless one. Deleting
   the local task on a reference workstation and forcing the refresh also
   recreated it from the updated preference with no GPO processing errors.
   The Windows Server 2016 DC had a stuck legacy task for the disconnected
   Administrator session; it was stopped and removed, and it is not recreated
   because the build fails the 1809+ filter.
6. **Repository hygiene.** No credentials were committed; `.env` remains
   gitignored; tracked files contain no copy of the `.env` password (verified
   via tracked-file scan).

### Why the preference approach

A Group Policy Preferences scheduled task runs as the interactive user at
logon and on a 15-minute repetition, so deployments recur for already logged-in
and unlocked sessions without user action. Item-level targeting lets a single
domain-linked GPO serve both modern and legacy clients without a separate OU
or WMI-filter infrastructure.

