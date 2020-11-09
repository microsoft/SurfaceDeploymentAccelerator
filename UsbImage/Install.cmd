setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

powercfg /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c

set psExe=powershell
if "%PROCESSOR_ARCHITECTURE%" equ "ARM64" set psExe="%ProgramFiles%\PowerShell\pwsh.exe"
pushd %~dp0
echo Preparing to Install
mode con:lines=9000
%psExe% -NoProfile -ExecutionPolicy bypass -command "& { %~dp0\Imaging.ps1 ; exit $LASTEXITCODE }"
set BANGERROR=!ERRORLEVEL!
popd
exit /b !BANGERROR!

