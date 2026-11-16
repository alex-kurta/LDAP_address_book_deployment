# Outlook LDAP Address Book Deployment

Deploys an LDAP address book to Outlook profiles for all domain users through
Group Policy, sets it as the default directory/search path entry, and keeps it
healthy without user interaction. The deployment authenticates as the signed-in
user (Secure Password Authentication); no password is stored on client machines.

## How it works

The NETLOGON package (`Deploy/`) is copied to
`\\<DC_SERVER>\NETLOGON\LdapABDeploy`. A user-scoped Group Policy Preferences
scheduled task on the domain runs the deployment package as the logged-on
user:

| Setting | Value |
|---------|-------|
| Preference item | Scheduled Tasks (user context, `TaskV2`), action `U` |
| Trigger | At logon, plus a 15-minute repetition while logged on |
| Run as | `%LogonDomain%\%LogonUser%`, interactive token, least privilege |
| Task action | `C:\Windows\System32\conhost.exe --headless cmd.exe /c "\\<DOMAIN>\NETLOGON\LdapABDeploy\Run-LdapAB.cmd"` |
| Item-level targeting | WMI filter: `SELECT * FROM Win32_OperatingSystem WHERE Version >= '10.0.17763'` |
| GPO link | Domain root (targeting OU did not exist), not enforced |

`Run-LdapAB.cmd` copies the package to `%LOCALAPPDATA%\LdapABDeploy` and runs
`Deploy-LdapAB.ps1`, which creates/updates the LDAP address book in each
Outlook profile, then sets it as the default directory and first entry in the
search path through the MAPI helper. Each run writes a per-user log to
`%TEMP%\LdapABDeploy.log`.

### Hidden background execution (no console popups)

Earlier revisions of the task action invoked `cmd.exe` directly, so employees
saw a console window flash every 15 minutes. The task action now launches
through `conhost.exe --headless`, which runs the wrapper with no visible
window. Verified on a Windows 11 build 26200 client: both `conhost.exe` and
`cmd.exe` reported `MainWindowHandle=0` and an empty window title, and the
deployment still completed with `VERIFIED default and search order`.

`conhost.exe --headless` is only available on Windows 10 version 1809 and
later. The preference item therefore carries a WMI item-level targeting filter
(`Version >= '10.0.17763'`) and is configured to remove itself where the
filter no longer matches. Windows 7, Windows XP, and Server 2016 clients
register no task and see no window.

### Authentication

SPA (`RequireSecurePasswordAuth=True`) with empty username/password fields.
The signed-in domain user is the directory account; nothing is rendered into
`ABConfiguration.xml`.

## Repository layout

| Path | Purpose |
|------|---------|
| `Deploy/` | NETLOGON client package: scripts, prebuilt x86/x64 exes, GPO creation automation, tests |
| `Deploy/GPO-DEPLOYMENT-GUIDE.md` | GPO creation walkthrough (RSAT and remote/WinRM paths) |
| `Deploy/DeploymentHelpers.ps1` | Shared helper functions used by the deploy scripts |
| `LdapABManager_Code/` | C++ source for the MAPI tool (based on the Microsoft MAPI sample) |
| `LdapABManager_Exe/` | Prebuilt x86/x64 MAPI executables (same binaries as `Deploy/`) |
| `LdapABManager_ABConfiguration.xml` | Sanitized configuration template |
| `.env.example` | Template for local configuration |
| `DEPLOYMENT_PROCESS.md` | Full process and decision log |

## Quick start

1. Create the local environment file:

   ```powershell
   Copy-Item .env.example .env
   notepad .env
   ```

2. Validate the configuration without changing Outlook or the registry:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy\Test-Configuration.ps1
   ```

3. Deploy the client package to NETLOGON:

   ```powershell
   .\Deploy\Deploy-All.ps1
   ```

4. Create and link the GPO. If RSAT is not installed on the administrative
   workstation, use WinRM. Preview first, then apply company-wide:

   ```powershell
   .\Deploy\Configure-LdapAB-GPO-Remote.ps1 -DryRun -LinkAtDomainRoot
   .\Deploy\Configure-LdapAB-GPO-Remote.ps1 -LinkAtDomainRoot
   ```

   The GPO is **not** created by copying files to NETLOGON. See
   `Deploy\GPO-DEPLOYMENT-GUIDE.md` for pilot-OU and manual GPMC instructions,
   and `DEPLOYMENT_PROCESS.md` for the full process notes.

> The repository automation provisions the GPO and the client package. The
> production scheduled-task preference (headless launcher, OS targeting,
> logon + 15-minute repetition) is documented above and in
> `DEPLOYMENT_PROCESS.md`; it is not scripted here and contains no
> environment-specific identifiers.

## Current status (verified 2026-11-16)

Validated in the reference Active Directory environment (identifiers removed):

| Item | Status |
|------|--------|
| `Outlook LDAP Address Book Deployment` GPO | Created, linked at the domain root (not enforced) |
| Scheduled task on modern clients | Runs hidden; logon trigger and 15-minute repetition |
| Task action | `conhost.exe --headless` wrapper; verified `MainWindowHandle=0`, no visible window |
| OS targeting | WMI `Version >= '10.0.17763'` (Windows 10 1809+ / Windows 11); older OSes get no task |
| Default address book | Deployed and set as first/default entry (`VERIFIED default and search order` in `%TEMP%\LdapABDeploy.log`) |
| Clients notified | Reference clients forced through `gpupdate /target:user /force`; task action confirmed as `conhost.exe --headless` and running successfully |
| Legacy client | Windows Server 2016 (build 14393) fails the OS filter; its stale task was stopped/removed |
| GPO version | Bumped to `393221`; identical `GPT.INI`/`ScheduledTasks.xml` hashes on both the Windows DC and the Samba DC SYSVOL mirror |
| Secondary Samba DC | Robocopy SYSVOL mirror healthy (`samba-tool ntacl sysvolcheck` passes); AD `versionNumber` replicated to `393221` |
| Untracked secrets | `.env` remains gitignored; no credentials in the repository |
| Production marker | `2.0` |

## Deployment notes

- Keep the GPO linked at the domain root when the desired OU does not exist.
- Do not switch the link to a pilot OU without first creating the target OU.
- Any new client (including unlocked sessions) receives the task at its next
  logon/background refresh; the task repeats every 15 minutes.
- If a client still shows a console window, confirm the client has refreshed
  the preference (`gpupdate /target:user /force`) and that it meets the
  1809+ OS floor.

## Security

- `.env` is gitignored. Never commit it.
- The deployment account in `.env` is used only for the SMB copy to NETLOGON
  and the WinRM connection used to create the GPO. It is never rendered into
  `ABConfiguration.xml`.
- The client configuration uses `RequireSecurePasswordAuth=True` with empty
  username and password fields.
- If a credential was ever committed or pushed, rotate it before continuing.
- Repository scans (tracked files) show no presence of the `.env` password.

## Rollback

- Unlink the GPO, or disable its scheduled-task preference item.
- Run `Deploy\Remove-LdapAB.ps1` on clients to remove the LDAP address book.
- Reverting the hidden launcher: restore the task action to the previous
  `cmd.exe` invocation, but that reintroduces console flashes — not
  recommended.
