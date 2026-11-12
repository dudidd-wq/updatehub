@echo off
rem UpdateHub - first start / update from the extracted folder.
rem It installs UpdateHub into C:\Program Files\UpdateHub, creates the "UpdateHub" shortcuts (desktop + Start menu)
rem and opens the dashboard. After that, use the shortcut - this file is not needed any more.
rem Every program is called by its full System32 path: cmd.exe searches the CURRENT folder first, so a file
rem named net.com or powershell.bat planted next to this script would otherwise run with administrator rights.
rem Messages are outside ( ) blocks on purpose: a folder name like "UpdateHub (1)" would break a block.
setlocal
set "NoDefaultCurrentDirectoryInExePath=1"
set "SYS=%SystemRoot%\System32"
set "PS=%SYS%\WindowsPowerShell\v1.0\powershell.exe"
set "UH_DIR=%~dp0"

rem A Group Policy execution policy (AllSigned / Restricted) overrides -ExecutionPolicy Bypass: say so clearly.
"%PS%" -NoProfile -NonInteractive -Command "$p = @((Get-ExecutionPolicy -Scope MachinePolicy), (Get-ExecutionPolicy -Scope UserPolicy)); if ($p -contains 'AllSigned' -or $p -contains 'Restricted') { exit 5 }"
if %errorlevel% equ 5 goto :gpo

rem Before ANYTHING from this folder runs as administrator: if other accounts on this PC can change files here
rem (e.g. C:\Tools inherits "Authenticated Users: Modify" from C:\), refuse. The check is inline - it runs no script
rem from this folder - and it runs twice: before asking for elevation and again in the elevated pass.
"%PS%" -NoProfile -NonInteractive -Command "try { $d = $env:UH_DIR.TrimEnd('\'); $b = 'S-1-1-0','S-1-5-11','S-1-5-32-545','S-1-5-4','S-1-5-32-546'; $m = 0x2 -bor 0x4 -bor 0x40 -bor 0x10000 -bor 0x40000 -bor 0x80000 -bor 0x10000000 -bor 0x40000000; $t = @($d) + @(Get-ChildItem -LiteralPath $d -File -Force -ErrorAction Stop | ForEach-Object { $_.FullName }); foreach ($x in $t) { $a = Get-Acl -LiteralPath $x -ErrorAction Stop; foreach ($e in $a.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier])) { if ($e.AccessControlType -ne 'Allow') { continue }; if ($x -eq $d -and ($e.PropagationFlags -band 2)) { continue }; if (($b -contains $e.IdentityReference.Value) -and ([int64]$e.FileSystemRights -band $m)) { exit 6 } } }; exit 0 } catch { exit 7 }"
if %errorlevel% neq 0 goto :unsafe

"%SYS%\net.exe" session >nul 2>&1
if %errorlevel% equ 0 goto :elevated
echo Requesting administrator rights...
"%PS%" -NoProfile -NonInteractive -Command "Start-Process -FilePath ($env:UH_DIR + 'run.bat') -Verb RunAs"
exit /b

:elevated
title UpdateHub setup
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"
if %errorlevel% neq 0 pause
exit /b

:gpo
echo.
echo  UpdateHub cannot start: a Group Policy allows only signed PowerShell scripts on this PC.
echo  Ask the administrator of this PC, or run UpdateHub on a machine without that policy.
echo.
pause
exit /b 5

:unsafe
echo.
echo  This folder can be changed by other accounts on this PC, or its permissions could not be read:
rem The path is printed by PowerShell: a folder name with "&" must never be parsed by cmd as a second command.
"%PS%" -NoProfile -NonInteractive -Command "Write-Host ('    ' + $env:UH_DIR)"
echo  For your safety UpdateHub will not run anything from here as administrator.
echo  Move the UpdateHub folder into your own profile (for example Documents) and run run.bat from there.
echo.
pause
exit /b 6
