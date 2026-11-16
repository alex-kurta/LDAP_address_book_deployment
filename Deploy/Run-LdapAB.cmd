@echo off
setlocal
set "SRC=%LOGONSERVER%\NETLOGON\LdapABDeploy"
set "DST=%LOCALAPPDATA%\LdapABDeploy"

if not exist "%DST%" mkdir "%DST%" >nul 2>&1

robocopy "%SRC%" "%DST%" /E /R:1 /W:1 /NP /NFL /NDL /NJH /NJS >nul
set "RC=%ERRORLEVEL%"
if %RC% GEQ 8 (
    echo %DATE% %TIME% [ERROR] Package copy failed with robocopy exit %RC% >> "%TEMP%\LdapABDeploy.log"
    exit /b %RC%
)

"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%DST%\Deploy-LdapAB.ps1"
exit /b %ERRORLEVEL%
