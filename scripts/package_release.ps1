<#
Stages a clean, minimal beta bundle from an existing build:
  LudicrousSpeed.xll (the ExcelDna packed add-in -- one file, no loose deps)
  ludicrous_engine.dll (the native engine the add-in P/Invokes into)
  Install-LudicrousSpeed.ps1 / Install.cmd / Uninstall.cmd
  README.txt

...then zips it. Run scripts\build_windows.ps1 first.
#>
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [string]$Version = "dev"
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$addinOutput = Join-Path $repoRoot "excel-addin\LudicrousSpeed.ExcelAddIn\bin\$Configuration\net48"
$engineDll = Join-Path $repoRoot "target\$($Configuration.ToLower())\ludicrous_engine.dll"

if (-not (Test-Path $addinOutput)) {
    throw "Add-in output not found at $addinOutput. Run scripts\build_windows.ps1 -Configuration $Configuration first."
}
if (-not (Test-Path $engineDll)) {
    throw "Native engine DLL not found at $engineDll. Run scripts\build_windows.ps1 -Configuration $Configuration first."
}

# ExcelDna.AddIn (Pack="true" in the .dna) emits both an unpacked loader xll
# and a self-contained "*packed*.xll" that embeds the managed assemblies.
# The packed one is what we want to ship -- everything managed in one file.
$packedXlls = @(Get-ChildItem $addinOutput -Filter "*packed*.xll")
if ($packedXlls.Count -eq 0) {
    throw "No packed .xll found in $addinOutput. Expected ExcelDna.AddIn to produce one (Pack=`"true`" is set in LudicrousSpeed.ExcelAddIn.dna) -- check the dotnet build output."
}
$packedXll = $packedXlls | Where-Object { $_.Name -match "64" } | Select-Object -First 1
if (-not $packedXll) { $packedXll = $packedXlls[0] }

Write-Host "Using packed add-in: $($packedXll.Name)"

$dist = Join-Path $repoRoot "dist"
$stagingName = "ludicrous-windows-$Version"
$staging = Join-Path $dist $stagingName
if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
New-Item -ItemType Directory -Force -Path $staging | Out-Null

Copy-Item $packedXll.FullName (Join-Path $staging "LudicrousSpeed.xll") -Force
Copy-Item $engineDll (Join-Path $staging "ludicrous_engine.dll") -Force
Copy-Item (Join-Path $PSScriptRoot "Install-LudicrousSpeed.ps1") $staging -Force
Copy-Item (Join-Path $PSScriptRoot "Install.cmd") $staging -Force
Copy-Item (Join-Path $PSScriptRoot "Uninstall.cmd") $staging -Force

@"
LudicrousSpeed beta ($Version)

Tip: before unzipping, right-click the downloaded .zip > Properties > check
"Unblock" > OK. That skips the Windows security prompt below entirely.

Install:
  1. Double-click Install.cmd
     If Windows says "Windows protected your PC", click "More info", then
     "Run anyway" -- this build isn't code-signed yet.
  2. Open Excel. LudicrousSpeed should already be on the ribbon.
     If it isn't: File > Options > Add-ins > Manage: Excel Add-ins > Go... >
     Browse... and select LudicrousSpeed.xll from the path Install.cmd printed.

Uninstall:
  Double-click Uninstall.cmd
"@ | Set-Content (Join-Path $staging "README.txt")

$zipPath = Join-Path $dist "$stagingName.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $staging "*") -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "Packaged beta bundle:"
Write-Host "  $staging\"
Write-Host "  $zipPath"
