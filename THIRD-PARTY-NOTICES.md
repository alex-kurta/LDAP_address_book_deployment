# Third-Party Notices

This repository contains third-party components. Their copyrights remain with
their respective owners; their inclusion does not grant any additional rights.

## LdapABManager (Microsoft LDAP Address Book Manager)

**Files covered:**

- `LdapABManager_Exe/x64/LdapABManager.exe`
- `LdapABManager_Exe/x86/LdapABManager.exe`
- `Deploy/LdapABManager_x64.exe`
- `Deploy/LdapABManager_x86.exe`

**Provenance.** These are the prebuilt 32-bit and 64-bit Windows executables of
the Microsoft sample **"LdapABManager"**, written by Andrei Ghita
(Microsoft). The version resources embedded in the binaries identify them as
`ProductName: Microsoft Outlook`, `CompanyName: Microsoft Corporation`,
`FileDescription: LdapABManager.exe`.

**Copyright.** Copyright (C) 2016 Microsoft Corporation. All Rights Reserved.

**Microsoft sample disclaimer** (printed by the tool at startup, and carried in
the sample's source header as distributed by Microsoft):

> Microsoft provides programming examples for illustration only, without
> warranty either expressed or implied. This includes, but is not limited to,
> the implied warranties of merchantability or fitness for a particular
> purpose. This sample assumes that you are familiar with the programming
> language that is being demonstrated with the tools that are used to create
> and to debug procedures. Microsoft support engineers can help explain the
> functionality of a particular procedure, but they will not modify these
> examples to provide added functionality or construct procedures to meet your
> specific requirements.

**Terms.** These binaries are redistributed here as **object code only**,
under the sample-code provisions of the Agreement for Microsoft Services
(<https://www.microsoft.com/en-us/services>, section "Sample Code"). That
section grants a perpetual, non-exclusive right to use, reproduce, and
distribute object code of sample code provided by Microsoft, conditioned on
(i) not using Microsoft's name, logo, or trademarks to market your software,
(ii) including a valid Microsoft copyright notice, and (iii) indemnifying
Microsoft. Consistent with those terms, the corresponding **source code is
not included or redistributed** in this repository.

**No endorsement.** Microsoft makes no warranties, express or implied,
regarding the sample or these binaries, and does not support, endorse, or
guarantee them. No Microsoft trademark is used to market this project.

**Rebuilding the binaries.** The source must be obtained from Microsoft
(contact Microsoft Customer Support for the "LdapABManager" sample). Rebuilding
requires the Microsoft Outlook 2010-or-later MAPI headers (freely downloadable
from Microsoft) and Visual Studio C++ targeting both x86 and x64; the sample
requires at least the Outlook 2010 MAPI headers (`MAPIFORM_CPU_X64` must be
defined). Microsoft support engineers can explain the sample's functionality
but will not modify it for you.

**Checksums (SHA-256) of the distributed binaries:**

- `LdapABManager_x64.exe` — `48a45b60add842e7b50c158bc950d73310eea6d7c9f6a939dc5b93ce8bcc92a2`
- `LdapABManager_x86.exe` — `0fc82ec72985ba69d84679a4d489816ceb6d86fe8a22c7ad34cd1f47ca6c1af2`

## SetLdapABDefault (this project)

`SetLdapABDefault_Code/SetLdapABDefault.cpp` and the prebuilt
`Deploy/SetLdapABDefault_x86.exe` / `Deploy/SetLdapABDefault_x64.exe` are
original work for this project and are not Microsoft sample code. Build
commands for both bitnesses are documented in the source file header (MinGW,
`-lmapi32 -lole32 -loleaut32`).
