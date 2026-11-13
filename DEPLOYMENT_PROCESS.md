# Complete Process Documentation: Outlook LDAP Address Book GPO Deployment

## Overview

This document describes how an LDAP address book is deployed from Active
Directory to Outlook clients through Group Policy. Environment-specific values
are stored in a local `.env` file and are never committed to the repository.

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
- `Configure-LdapAB-GPO.ps1` — GPO automation
- `Deploy-All.ps1` — end-to-end deployment automation
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

**Automated options:**

- `Deploy-All.ps1` creates and links the GPO when the GroupPolicy module is
  available and `TARGET_OU` is set.
- `Configure-LdapAB-GPO.ps1` performs the same steps standalone on a domain
  controller or RSAT workstation.

**Manual steps via GPMC:**

1. Create a GPO and link it to the target OU.
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

- Registry marker: `HKCU\Software\<DOMAIN>\LdapABDeploy` with `Version=1.0`
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

- GPO creation requires the GroupPolicy module (RSAT).
- NETLOGON write access may require explicit credentials.
- The deployment package must be re-rendered when `.env` changes.
- Rewriting Git history is required if a secret was committed before it was
  removed.

## Next Steps

1. Fill in `.env` and run `Deploy\Test-Configuration.ps1`.
2. Run `Deploy\Deploy-All.ps1` or copy the package manually.
3. Create and link the GPO if it was not automated.
4. Pilot the deployment on a test OU.
5. Monitor `%TEMP%\LdapABDeploy.log` on pilot machines.
6. Roll out to production OUs.
7. Optionally enable LDAPS (port 636) with a valid certificate.
