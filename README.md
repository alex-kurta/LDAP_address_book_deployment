# Outlook LDAP Address Book Deployment

Deploys an LDAP address book to Outlook profiles through a Group Policy user
logon script. The deployment uses Secure Password Authentication, so the
address book authenticates as the signed-in user and no password is stored on
client machines.

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

3. Deploy the client package to NETLOGON and optionally create the GPO:

   ```powershell
   .\Deploy\Deploy-All.ps1
   ```

See `Deploy\GPO-DEPLOYMENT-GUIDE.md` for manual GPO steps and
`DEPLOYMENT_PROCESS.md` for the full process notes.

## Security

- `.env` is gitignored. Never commit it.
- The deployment account in `.env` is used only for the SMB copy to NETLOGON.
  It is never rendered into `ABConfiguration.xml`.
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
