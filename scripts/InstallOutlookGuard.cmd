@echo off
REM Same scriptblock approach as Install.cmd -- see the comment there for why
REM the .ps1 is not loaded from disk.
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "Get-ChildItem -LiteralPath '%~dp0' -File | Unblock-File -ErrorAction SilentlyContinue; & ([scriptblock]::Create((Get-Content -Raw -LiteralPath '%~dp0Install-OutlookGuard.ps1'))) -BundleDir '%~dp0'"
echo.
pause
