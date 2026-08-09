<#
Stages a clean, minimal beta bundle from an existing build:
  LudicrousSpeed.xll (the ExcelDna packed add-in)
  ludicrous_engine.dll (the native engine the add-in P/Invokes into)
  any managed dependencies the packed .xll does not embed (see below)
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
#
# Search recursively: for SDK-style projects ExcelDna.AddIn puts the packed
# output in a "publish" subdirectory of the build output, not beside the
# unpacked loader. That default is configurable (ExcelDnaPackXllOutputDirectory,
# where %none% means "same folder"), so recursing covers both layouts rather
# than hard-coding publish\ and breaking if it is ever changed.
$packedXlls = @(Get-ChildItem $addinOutput -Filter "*packed*.xll" -Recurse)
if ($packedXlls.Count -eq 0) {
    $found = @(Get-ChildItem $addinOutput -Filter "*.xll" -Recurse |
        ForEach-Object { $_.FullName.Substring($addinOutput.Length).TrimStart('\') })
    $foundText = if ($found.Count -gt 0) { $found -join ", " } else { "(no .xll files at all)" }
    throw "No packed .xll found under $addinOutput. Expected ExcelDna.AddIn to produce one (Pack=`"true`" is set in LudicrousSpeed.ExcelAddIn.dna). Found instead: $foundText"
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

# Managed dependencies that the packed .xll does not necessarily embed.
# Under .NET Framework, ExcelDnaPack packs the ExternalLibrary and anything
# listed as a <Reference Pack="true"> in the .dna -- and our .dna lists only
# the ExternalLibrary. So Newtonsoft.Json in particular is a real runtime
# dependency that would otherwise be missing, and the add-in would install
# cleanly and then fail on load. Excel-DNA resolves assemblies from the .xll's
# own folder, so shipping them loose beside it is enough.
#
# Deliberately an allowlist rather than "every .dll in the output": a stray
# loose copy of ExcelDna.Integration.dll can conflict with the one built into
# the packed .xll. Missing files are fatal, redundant ones are not, so
# anything genuinely embedded already just sits here unused.
$sidecars = @("Newtonsoft.Json.dll", "Microsoft.Office.Interop.Excel.dll", "office.dll")
foreach ($name in $sidecars) {
    $candidate = Join-Path $addinOutput $name
    if (Test-Path $candidate) {
        Copy-Item $candidate $staging -Force
        Write-Host "Included dependency: $name"
    }
}

# Not optional. The add-in deserializes every engine response with Json.NET, so
# a bundle without it installs cleanly and then fails on load with an
# assembly-load dialog that says nothing about packaging. Fail the build here
# instead, where the cause is obvious.
if (-not (Test-Path (Join-Path $staging "Newtonsoft.Json.dll"))) {
    throw "Newtonsoft.Json.dll was not found in $addinOutput, so the bundle would ship without it and the add-in would fail to load in Excel. Check that the build actually copied it to the output directory."
}
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
