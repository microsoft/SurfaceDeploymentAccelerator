@echo off
setlocal EnableExtensions EnableDelayedExpansion
title "Main WinPE Prompt - startnet"
echo Disabling Quickedit
reg delete HKCU\Console /v Quickedit /f

:NET
wpeinit
wpeutil WaitForRemovableStorage
wpeutil UpdateBootInfo
wpeutil InitializeNetwork

REM If WinPE RAM disk X: is being used, discover drive letter for its physical media
set BOOTVOLUME=
if exist "X:\windows\system32\*" (
    REM Discover location of 'DeviceName / Boot' volume
    echo Checking for X:\
    if exist x:\install.cmd (
        set BOOTVOLUME=x:
        echo BOOTMOLUME !BOOTVOLUME!
    )

) else (
    echo X:\Windows\System32 not found
    set BOOTVOLUME=%~d0
    echo BOOTVOLUME !BOOTVOLUME!
)

if "!BOOTVOLUME!"=="" goto :errCannotFindSourceMedia
echo Pre-Trim  BOOTVOLUME !BOOTVOLUME!
if "!BOOTVOLUME:~-1,1!"=="\" set BOOTVOLUME=!BOOTVOLUME:~0,-1!
echo Post-Trim BOOTVOLUME !BOOTVOLUME!

REM Where is install.cmd?
if exist "!BOOTVOLUME!\install.cmd" set INSTALLCMD=!BOOTVOLUME!\install.cmd

if not exist "!INSTALLCMD!" (
    echo install.cmd does not exist on at !INSTALLCMD!
    echo Exiting to command-line ...
    goto :EOF
)

echo Calling !INSTALLCMD!
start /max cmd.exe /k "Mode con lines=9900 && !INSTALLCMD!"
exit /b 0
goto :EOF

:errCannotFindSourceMedia
    cls
    echo.
    echo ERROR: Cannot find the drive letter for the USB media where WinPE was booted from.
    exit /b 1

:EOF
