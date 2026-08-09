@echo off
REM Runs the installer's *text* as a scriptblock rather than loading the .ps1
REM from disk. Under a policy-enforced AllSigned execution policy -- common on
REM managed corporate desktops -- an unsigned script file cannot be loaded at
REM all, and the -ExecutionPolicy Bypass passed on the command line does not
REM override a policy set by Group Policy. Execution policy governs script
REM files, not commands, so this path works in both environments.
REM
REM Unblock-File first clears Mark of the Web on the extracted files, which is
REM the other reason a freshly downloaded bundle refuses to run.
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "Get-ChildItem -LiteralPath '%~dp0' -File | Unblock-File -ErrorAction SilentlyContinue; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '%~dp0Install-LudicrousSpeed.ps1'))) -BundleDir '%~dp0'"
echo.
pause
