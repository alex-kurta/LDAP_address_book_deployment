# Outlook LDAP Address Book Deployment

Deploys an LDAP address book to Outlook profiles for domain users through Group
Policy, sets it as the default directory and first entry in the address-book
search path, and keeps it current without user interaction. Clients bind with
Secure Password Authentication (SPA) as the signed-in user; no password is
stored on client machines or in the NETLOGON package.

## How it works

1. `Deploy-All.ps1` renders `ABConfiguration.xml` from your `.env` values and
   copies the client package to `\\<DC_SERVER>\NETLOGON\LdapABDeploy`.
2. A user-scoped GPO linked at the domain root (or a pilot OU) runs the
   package as each logged-on user:
   - **Recommended:** a Group Policy Preferences scheduled task that launches
     hidden through `conhost.exe --headless`, repeats every 15 minutes while
     the user is logged on, and is WMI-targeted to Windows 10 1809+ (Option A
     in `Deploy\GPO-DEPLOYMENT-GUIDE.md`).
   - **Alternative:** a user logon script via `psscripts.ini` (Option B).
3. `Run-LdapAB.cmd` copies the package to `%LOCALAPPDATA%\LdapABDeploy` and
   runs `Deploy-LdapAB.ps1`, which creates or updates the address book in each
   Outlook profile, then runs `SetLdapABDefault` to make it the default
   directory and first search-path entry. Each run logs to
   `%TEMP%\LdapABDeploy.log` and ends with
   `LDAP Address Book deployed and set as default`.

`conhost.exe --headless` runs the wrapper with no visible console window. It
exists only on Windows 10 1809+, so the preference item carries a WMI
item-level filter (`Version >= '10.0.17763'`) and removes itself where the
filter no longer matches; older clients register no task and see no window.

## Requirements

- Admin workstation: Windows PowerShell 5.1; RSAT Group Policy tools **or**
  WinRM access to a domain controller with the GroupPolicy module.
- `.env` deployment account with NETLOGON write access and DC logon rights
  (leave `USERNAME`/`PASSWORD` blank to use the current account).
- Clients: Outlook 2016/2019/2021/Microsoft 365, x86 or x64 (auto-detected);
  Windows 10 1809+ for the hidden task.
- LDAP server reachable from client VLANs on `LDAP_PORT` (389, or 636 with
  `USE_SSL=true`).

## Setup

Full walkthrough: `Deploy\GPO-DEPLOYMENT-GUIDE.md`.

```powershell
Copy-Item .env.example .env          # fill in your environment values
.\Deploy\Test-Configuration.ps1      # safe smoke test, changes nothing
.\Deploy\Deploy-All.ps1              # render + copy package to NETLOGON
.\Deploy\Configure-LdapAB-GPO-Remote.ps1 -DryRun -LinkAtDomainRoot
.\Deploy\Configure-LdapAB-GPO-Remote.ps1 -LinkAtDomainRoot
```

The scripts create the logon-script GPO (Option B). For the recommended
hidden scheduled task (Option A), add the preference item in GPMC as described
in the guide; both mechanisms run the same client package.

## Repository layout

| Path | Purpose |
|------|---------|
| `Deploy/` | NETLOGON client package: scripts, prebuilt x86/x64 helpers, GPO automation, smoke test |
| `Deploy/GPO-DEPLOYMENT-GUIDE.md` | Step-by-step setup guide (Options A/B) |
| `Deploy/DeploymentHelpers.ps1` | Shared helpers used by the deploy scripts |
| `Deploy/Run-LdapAB.cmd` | Client wrapper: copies package locally, runs deploy script |
| `Deploy/SetLdapABDefault_x86/x64.exe` | MAPI helper: sets default address book + search order |
| `SetLdapABDefault_Code/` | C++ source for the helper (MinGW build commands included) |
| `LdapABManager_Exe/` | Prebuilt x86/x64 executables of Microsoft's LDAP MAPI sample (object code; see notices) |
| `LdapABManager_ABConfiguration.xml` | Sanitized configuration template |
| `THIRD-PARTY-NOTICES.md` | Copyright/terms for the Microsoft LdapABManager binaries |
| `.env.example` | Template for local configuration |

## Verification

- `gpresult /r` on a client lists the GPO; `schtasks /query` shows the task
  with the `conhost.exe --headless` action.
- `%TEMP%\LdapABDeploy.log` ends with
  `LDAP Address Book deployed and set as default`.
- Outlook address book (`Ctrl+Shift+B`) lists the configured display name
  first.

## Rollback

- Unlink or disable the GPO (or delete the preference item).
- Run `Deploy\Remove-LdapAB.ps1` on clients; it logs
  `LDAP Address Book removed successfully`.
- Remove the registry marker `HKCU\Software\<DOMAIN>\LdapABDeploy` if present.

## Security

- `.env` is gitignored. Never commit it.
- The deployment account in `.env` is used only for the SMB copy to NETLOGON
  and the WinRM connection used to create the GPO. It is never rendered into
  `ABConfiguration.xml`.

## Third-party components

The `LdapABManager_*.exe` executables are Microsoft's prebuilt "LdapABManager"
sample binaries, redistributed as object code only under the sample-code terms
of the Agreement for Microsoft Services; the sample source is not redistributed.
See `THIRD-PARTY-NOTICES.md` for copyright, checksums, and rebuild
prerequisites. `SetLdapABDefault` (source included) is this project's own work.
- The client configuration uses `RequireSecurePasswordAuth=True` with empty
  username and password fields.
- If a credential was ever committed or pushed, rotate it before continuing.
