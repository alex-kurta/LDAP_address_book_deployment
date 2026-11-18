# GPO Deployment Guide: Outlook LDAP Address Book

This guide walks through setting up the Outlook LDAP address book deployment
from scratch on a new domain: configure, validate, copy the client package to
NETLOGON, and create/link the GPO. Environment-specific values live in a local
`.env` file; the package copied to the domain controller contains only a
rendered `ABConfiguration.xml`, the client scripts, and the prebuilt
executables — no `.env` and no credentials.

Copying the client package to NETLOGON does **not** create or link the GPO.
The GPO is created separately, either through the scripts in `Deploy/` or
manually in `gpmc.msc`.

## Prerequisites

- Windows PowerShell 5.1 or later on the administrative workstation
- Outlook installed on target clients
- One of these for GPO creation:
  - RSAT Group Policy tools on the workstation, or
  - WinRM access to a domain controller with the GroupPolicy module
- Write access to the domain controller's NETLOGON share
- An account with domain logon rights for the NETLOGON copy and (if used) the
  WinRM connection. Blank `USERNAME`/`PASSWORD` in `.env` uses your current
  account.

## 1. Configure `.env`

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
- `DOMAIN_FQDN` / `NETBIOS_DOMAIN` — set explicitly when `DOMAIN` is a
  NetBIOS name and cannot be derived from `TARGET_OU`

For company-wide deployment, pass `-LinkAtDomainRoot`; it derives the domain
distinguished name from `DOMAIN` or `TARGET_OU` and overrides the configured
`TARGET_OU` link scope. Use a pilot OU (`-TargetOU`) first if you want to
validate on a subset of users.

> Do not commit `.env`. It is gitignored.

## 2. Validate the configuration

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy\Test-Configuration.ps1
```

The smoke test (changes nothing in Outlook or the registry):

- loads `.env` and validates required values without printing them
- derives and checks the domain and NETLOGON paths
- renders `ABConfiguration.xml` into a temporary directory
- checks the rendered config uses Secure Password Authentication and has
  embedded credentials empty
- runs `Deploy-LdapAB.ps1` and `Remove-LdapAB.ps1` in `-DryRun` mode
- runs `Deploy-All.ps1 -SkipConnectivityTest -SkipCopy` to validate package
  generation
- parses every PowerShell script for syntax errors
- verifies `.env` is untracked and tracked files contain no obvious token,
  IP-address, or password literals

## 3. Deploy the client package

```powershell
.\Deploy\Deploy-All.ps1
```

Useful parameters:

```powershell
.\Deploy\Deploy-All.ps1 -EnvFile .\.env -SkipGPO
.\Deploy\Deploy-All.ps1 -DomainController dc01.example.com -TargetOU 'OU=Pilot,DC=example,DC=com'
.\Deploy\Deploy-All.ps1 -SkipConnectivityTest

# Validate and render only; copy nothing:
.\Deploy\Deploy-All.ps1 -SkipConnectivityTest -SkipCopy
```

`Deploy-All.ps1` renders the client configuration into a temporary staging
folder and copies only these files to `\\<DC_SERVER>\NETLOGON\LdapABDeploy`:

```
ABConfiguration.xml
Deploy-LdapAB.ps1
Remove-LdapAB.ps1
Run-LdapAB.cmd
LdapABManager_x86.exe
LdapABManager_x64.exe
SetLdapABDefault_x86.exe
SetLdapABDefault_x64.exe
```

The administrative scripts and `.env` are not copied.

If the GroupPolicy module is available locally, `Deploy-All.ps1` also creates,
links, and configures the logon-script GPO (Option B). If it is not installed,
configure the GPO in step 4 after the copy completes.

## 4. Create and link the GPO

Choose one mechanism. Both run the same client package.

| Mechanism | Behavior | Use when |
|-----------|----------|----------|
| A. Scheduled task (recommended) | Hidden launch at logon + 15-minute repeat while logged on; Windows 10 1809+ targeting | You want already-logged-on sessions to pick up the deployment |
| B. User logon script | Runs once at each logon; set up automatically by the scripts | You prefer a simple, scriptable path |

### Option A: scheduled task preference (recommended)

1. In **Group Policy Management** (`gpmc.msc`), right-click the domain root
   (company-wide) or a pilot OU → **Create a GPO in this domain, and Link it
   here**. Name it (or use the value of `GPO_NAME`).
2. Edit the GPO → **User Configuration → Preferences → Control Panel Settings
   → Scheduled Tasks** → **New → Scheduled Task (At least Windows Vista)**.
   Set the action on the **General** tab to:
   - Name: `Outlook LDAP Address Book Deployment`
   - Security options: **Run only when user is logged on**; **Do not store
     password**; do **not** check *Run with highest privileges*
   - The task runs as the user to whom the preference item applies, with an
     interactive token and least privilege.
3. On the **Triggers** tab, add:
   - **At log on** — any user, delay **1 minute**
   - A **daily** trigger with **Repeat task every: 15 minutes** for
     **Indefinitely** (keeps already-logged-on sessions current)
4. On the **Actions** tab, **Start a program**:
   - Program: `C:\Windows\System32\conhost.exe`
   - Arguments: `--headless cmd.exe /c "\\<DOMAIN>\NETLOGON\LdapABDeploy\Run-LdapAB.cmd"`
   - `conhost.exe --headless` runs the wrapper with no visible console window.
     It exists only on Windows 10 version 1809 and later.
5. On the **Settings** tab:
   - If the task is already running: **Do not start a new instance**
   - **Start the task as soon as possible after a scheduled start is missed**
   - **Stop the task if it runs longer than 1 hour**
6. On the **Common** tab:
   - Check **Remove this item when it is no longer applied**
   - Open **Item-level targeting** → **New → WMI query** and set:
     - Namespace: `Root\cimv2`
     - Query: `SELECT * FROM Win32_OperatingSystem WHERE Version >= '10.0.17763'`
   - This restricts the item to Windows 10 1809+/Windows 11 and removes the
     task from clients that no longer match.
7. **Apply**, then close the editor. Clients apply the item at the next
   background refresh or logon; `gpupdate /target:user /force` on a test
   machine forces immediate application.

The resulting `User\Preferences\ScheduledTasks\ScheduledTasks.xml` on SYSVOL
resembles the following (GPMC generates `uid`/`changed` automatically; `<DOMAIN>`
is your DNS domain name):

```xml
<?xml version="1.0" encoding="utf-8"?>
<ScheduledTasks clsid="{CC63F200-7309-4ba0-B154-A71CD118DBCC}">
  <TaskV2 clsid="{D8896631-B747-47a7-84A6-C155337F3BC8}"
          name="Outlook LDAP Address Book Deployment"
          image="0" uid="{00000000-0000-0000-0000-000000000000}"
          userContext="0" removePolicy="1"
          desc="Deploys the Outlook LDAP address book at logon and every 15 minutes while the user is logged on.">
    <Properties action="U" name="Outlook LDAP Address Book Deployment"
                runAs="%LogonDomain%\%LogonUser%" logonType="InteractiveToken">
      <Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
        <Triggers>
          <LogonTrigger>
            <Enabled>true</Enabled>
            <Delay>PT1M</Delay>
          </LogonTrigger>
          <CalendarTrigger>
            <Enabled>true</Enabled>
            <Repetition>
              <Interval>PT15M</Interval>
              <StopAtDurationEnd>false</StopAtDurationEnd>
            </Repetition>
            <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
          </CalendarTrigger>
        </Triggers>
        <Principals>
          <Principal id="Author">
            <UserId>%LogonDomain%\%LogonUser%</UserId>
            <LogonType>InteractiveToken</LogonType>
            <RunLevel>LeastPrivilege</RunLevel>
          </Principal>
        </Principals>
        <Actions Context="Author">
          <Exec>
            <Command>C:\Windows\System32\conhost.exe</Command>
            <Arguments>--headless cmd.exe /c "\\<DOMAIN>\NETLOGON\LdapABDeploy\Run-LdapAB.cmd"</Arguments>
          </Exec>
        </Actions>
      </Task>
    </Properties>
    <Filters>
      <FilterWmi bool="AND" not="0" nameSpace="Root\cimv2"
                query="SELECT * FROM Win32_OperatingSystem WHERE Version &gt;= '10.0.17763'"/>
    </Filters>
  </TaskV2>
</ScheduledTasks>
```

If you verify the file instead of creating the item through the GUI, confirm
`removePolicy="1"`, the `conhost.exe --headless` action, and the
`FilterWmi` query are present. If you hand-edit the XML on SYSVOL, clients
pick up the change at the next background refresh (or after `gpupdate`).

### Option B: user logon script (automated)

For a workstation without RSAT, `Configure-LdapAB-GPO-Remote.ps1` uses WinRM
and the deployment credentials in `.env` to run the GroupPolicy cmdlets on
the domain controller. Preview first:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy\Configure-LdapAB-GPO-Remote.ps1 -DryRun -LinkAtDomainRoot
```

Then apply it:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy\Configure-LdapAB-GPO-Remote.ps1 -LinkAtDomainRoot
```

For a pilot OU instead:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy\Configure-LdapAB-GPO-Remote.ps1 -TargetOU 'OU=Pilot,DC=example,DC=com'
```

On a domain controller or RSAT workstation:

```powershell
.\Deploy\Configure-LdapAB-GPO.ps1 -LinkAtDomainRoot
# or
.\Deploy\Configure-LdapAB-GPO.ps1 -TargetOU 'OU=Pilot,DC=example,DC=com'
```

The scripts:

- create the GPO if it does not exist
- link it to the domain root or the specified OU
- write the user logon script to `User\Scripts\psscripts.ini` as UTF-16LE
  with a byte-order mark (the Scripts CSE ignores other encodings and the
  `Scripts\Logon` path)
- register the Scripts client-side extension in the GPO
- enable **Run logon scripts synchronously**
- bump the GPO user version and update `gpt.ini` so clients process the change
- refuse to overwrite an existing GPO that already has other logon scripts

Rerunning is safe; the GPO is updated in place. The deployment password is
used only for the WinRM connection and is never printed or written to the
policy files.

Logon scripts run at logon only; they do not repeat on already-logged-on
sessions and do not run hidden for this mechanism — use Option A if you need
either property.

### Option C: logon script (manual GPMC)

1. Open **Group Policy Management** (`gpmc.msc`).
2. Right-click the domain root (company-wide) or a pilot OU → **Create a GPO
   in this domain, and Link it here**. Name it (or use the value of `GPO_NAME`).
3. Edit the GPO → **User Configuration → Policies → Windows Settings →
   Scripts (Logon/Logoff)** → **Logon** → **PowerShell Scripts** → **Add**.
4. Script name: `\\<DC_SERVER>\NETLOGON\LdapABDeploy\Deploy-LdapAB.ps1`.
   Leave parameters empty.
5. Enable **User Configuration → Policies → Administrative Templates → System
   → Scripts → Run logon scripts synchronously**.

### Security filtering

The default **Authenticated Users** filter applies the GPO to everyone in the
linked scope. Restrict it to a pilot security group if needed.

## 5. How the client side works

`Deploy-LdapAB.ps1` reads the pre-rendered `ABConfiguration.xml` on the
client; it does not read `.env` and contains no credentials. The script:

- detects Outlook bitness (x86/x64) and selects the matching prebuilt
  executables (`LdapABManager_*.exe`, `SetLdapABDefault_*.exe`)
- creates or updates the LDAP address book as **Secure Password Authentication**
  with the signed-in user's credentials (`RequireSecurePasswordAuth=True`,
  empty username/password)
- runs `SetLdapABDefault` to make the new address book the default directory
  and the first entry in the address-book search path
- writes a registry marker `HKCU\Software\<DOMAIN>\LdapABDeploy` and logs to
  `%TEMP%\LdapABDeploy.log`
- exits without changes in `-DryRun` mode

`LdapABManager_ABConfiguration.xml` documents every supported setting. The
`LdapABManager_*.exe` files are prebuilt Microsoft "LdapABManager" sample
binaries, redistributed as object code only (see `THIRD-PARTY-NOTICES.md`);
their source is not included or redistributed. To rebuild them, obtain the
sample from Microsoft and install the Microsoft Outlook 2010-or-later MAPI
headers, then build with Visual Studio C++ targeting x86 and x64.
`SetLdapABDefault_Code/SetLdapABDefault.cpp` is this project's own source; to
build it instead of using the prebuilt binaries:

- `SetLdapABDefault_Code/SetLdapABDefault.cpp` — MinGW:
  `x86_64-w64-mingw32-g++ -O2 -municode -static -static-libgcc -static-libstdc++ -o SetLdapABDefault_x64.exe SetLdapABDefault.cpp -lmapi32 -lole32 -loleaut32`
  (use the 32-bit toolchain for the `x86` binary)

## Verification checklist

| Test | Method | Expected result |
|------|--------|-----------------|
| GPO exists | `Get-GPO -Name 'Outlook LDAP Address Book Deployment'` on the DC | GPO is returned |
| GPO linked | `Get-GPInheritance -Target 'DC=example,DC=com'` on the DC | GPO appears in `GpoLinks` |
| Task present (Option A) | `schtasks /query /tn 'Outlook LDAP Address Book Deployment'` on a client | Action is `conhost.exe --headless ...` |
| Policy applied | `gpresult /r` on a client | GPO listed under user policy |
| Script ran | `%TEMP%\LdapABDeploy.log` | Last line: `LDAP Address Book deployed and set as default` |
| AB created | Outlook → Address Book → Tools → Options | Configured display name is listed |
| AB is default | Open Address Book (`Ctrl+Shift+B`) | New address book appears first |
| Search works | Type a name in the **To:** field | Resolves from AD |
| Re-run is idempotent | Log off and on again | Log shows normal redeploy with no duplicates |
| Legacy client (Option A) | `gpresult /r` | Off-filter client registers no task |

## Troubleshooting

### Script fails with "Outlook not found"

- Verify Outlook is installed.
- Check `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE`.
- The client script auto-detects Outlook bitness and selects the matching
  executables.

### Address book is created but not default

- The deployment runs `SetLdapABDefault`, which places the address book at
  search index #0 and sets it as the default directory.
- If it is not first, check Outlook's "Choose automatically" setting:
  `HKCU\Software\Microsoft\Office\16.0\Outlook\Options\General\ABChooseDirAutomatically = 0`.

### Connection fails

- Verify the LDAP port is reachable from client VLANs to the LDAP server.
- Test with `ldp.exe` or `Test-NetConnection <LDAP_SERVER> -Port <LDAP_PORT>`.
- Check the firewall on the LDAP server.

### Duplicate address books on re-run

- The script checks for an existing address book by display name and skips
  re-creation; it re-verifies the default position on every run.
- The registry marker `HKCU\Software\<DOMAIN>\LdapABDeploy` records the
  deployed version.

### Users see a flashing console window (Option A)

- The task action must be `conhost.exe --headless ...`, not `cmd.exe`.
- Confirm the item-level WMI filter is present and the client is at or above
  the target build.
- Force a refresh in the affected user's session: `gpupdate /target:user /force`,
  then log off and on.

### GPO does not apply

- `gpresult /r` and `gpresult /h report.html` on the client; check security
  filtering and WMI filters.
- From client VLANs, verify SYSVOL reachability (`\\<DOMAIN>\SYSVOL`); on
  multi-DC environments confirm every DC serves the same `gpt.ini` and the
  GPO's preference files.
- Logon-script GPOs need the Scripts CSE registered (see Option B) and
  UTF-16LE `psscripts.ini`; other encodings are ignored.

## Rollback

### Disable the GPO

Unlink the GPO from the domain root or OU, or disable the link:

```powershell
Remove-GPLink -Name 'Outlook LDAP Address Book Deployment' -Target 'DC=example,DC=com'
```

To remove the GPO entirely:

```powershell
Remove-GPO -Name 'Outlook LDAP Address Book Deployment'
```

To remove only the scheduled-task item (Option A): edit the GPO and delete the
preference item. Clients with `removePolicy="1"` drop their local task
automatically off-filter or after the item is deleted.

### Remove the address book on clients

```powershell
\\<DC_SERVER>\NETLOGON\LdapABDeploy\Remove-LdapAB.ps1
```

The script logs `LDAP Address Book removed successfully` and removes the
registry marker. To clean up manually per user, remove
`HKCU\Software\<DOMAIN>\LdapABDeploy`.

## Requirements

- **LDAP server**: reachable on the configured port
- **Clients**: Outlook 2016/2019/2021/Microsoft 365 (x86 or x64); Windows
  10 1809+ or Windows 11 for Option A
- **Permissions**: domain users need read access to the directory
- **PowerShell**: 3.0+ on clients (5.1 recommended)

## Log locations

- **Deployment log**: `%TEMP%\LdapABDeploy.log`
- **Removal log**: `%TEMP%\LdapABRemove.log`
- **Registry marker**: `HKCU\Software\<DOMAIN>\LdapABDeploy`
- **Package copy staging**: `%LOCALAPPDATA%\LdapABDeploy`

## Version history

| Version | Date | Changes |
|---------|------|---------|
| 2.0 | 2026-11-16 | Setup-only documentation; hidden `conhost.exe --headless` launcher; Windows 10 1809+ WMI item-level targeting with remove-off-filter; `SetLdapABDefault` helper; `%LOGONSERVER%`-relative wrapper |
