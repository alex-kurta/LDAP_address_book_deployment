# Outlook LDAP Address Book Deployment

Deploys an LDAP address book to Outlook profiles through a Group Policy user
logon script. The deployment uses Secure Password Authentication, so the
address book authenticates as the signed-in user and no password is stored on
client machines.

## Deployment Status

The deployment package was validated against a reference Active Directory
environment. In that environment the `Outlook LDAP Address Book Deployment` GPO
was created and linked at the domain root, and the NETLOGON client package was
refreshed to match this repository.

> Reference-environment limitation: a secondary Samba domain controller's SMB
> service rejected session setup with "The security ID structure is invalid",
> so its robocopy-mirrored SYSVOL was empty. Clients whose DFS referral landed
> on that controller could not process domain GPOs until the Samba SMB service
> and SID mapping were repaired and the mirror was refreshed. Windows-only
> diagnostics such as `dcdiag`, `dfsrdiag`, WMI, and WinRM do not apply to
> Samba domain controllers.

Environment-specific names, addresses, and identifiers are intentionally not
stored in this repository. Supply them through a local `.env` file.

## Quick Start

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

## Security

- `.env` is gitignored. Never commit it.
- The deployment account in `.env` is used only for the SMB copy to NETLOGON
  and the WinRM connection used to create the GPO. It is never rendered into
  `ABConfiguration.xml`.
- The client configuration uses `RequireSecurePasswordAuth=True` with empty
  username and password fields.
- If a credential was ever committed or pushed, rotate it before continuing.

## Repository Layout

| Path | Purpose |
|------|---------|
| `Deploy/` | Client scripts, executables, GPO automation, tests |
| `LdapABManager_Code/` | C++ source for the MAPI tool |
| `LdapABManager_Exe/` | Prebuilt x86/x64 executables |
| `.env.example` | Template for local configuration |
| `DEPLOYMENT_PROCESS.md` | Detailed process documentation |
