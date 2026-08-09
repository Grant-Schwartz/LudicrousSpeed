<#
Installs the LudicrousSpeed Excel add-in for the current user only:
  1. copies LudicrousSpeed.xll and ludicrous_engine.dll to %LOCALAPPDATA%\LudicrousSpeed
  2. unblocks them, so Windows stops treating them as downloaded-from-the-internet
  3. registers that folder as a trusted location in every detected Excel version,
     so the Trust Center doesn't block the add-in
  4. registers the add-in to auto-load in every detected Excel version, so it's
     already on the ribbon next time Excel opens -- no manual Browse step

Everything here is per-user (HKCU, %LOCALAPPDATA%) -- no admin rights needed.
Prefer double-clicking Install.cmd / Uninstall.cmd over running this directly;
they set the execution policy for you so you don't hit a PowerShell prompt.

Run this from inside the unzipped LudicrousSpeed folder (it looks for
LudicrousSpeed.xll and ludicrous_engine.dll next to itself). Safe to re-run.

Usage:
  .\Install-LudicrousSpeed.ps1
  .\Install-LudicrousSpeed.ps1 -Uninstall
#>
param(
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$installDir = Join-Path $env:LOCALAPPDATA "LudicrousSpeed"
$trustDescription = "LudicrousSpeed Beta"

function Get-ExcelVersionKeys {
    Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue |
        Where-Object { Test-Path (Join-Path $_.PSPath "Excel") } |
        ForEach-Object { $_.PSChildName }
}

function Remove-TrustedLocation {
    param([string]$OfficeVersion)

    $locationsKey = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Excel\Security\Trusted Locations"
    if (-not (Test-Path $locationsKey)) { return }

    Get-ChildItem $locationsKey | ForEach-Object {
        $desc = (Get-ItemProperty $_.PSPath -Name Description -ErrorAction SilentlyContinue).Description
        if ($desc -eq $trustDescription) {
            Remove-Item $_.PSPath -Force -Recurse
            Write-Host "Removed trusted location entry: $($_.PSChildName) (Excel $OfficeVersion)"
        }
    }
}

function Add-TrustedLocation {
    param([string]$OfficeVersion, [string]$Path)

    $locationsKey = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Excel\Security\Trusted Locations"
    New-Item -Path $locationsKey -Force | Out-Null

    $already = Get-ChildItem $locationsKey -ErrorAction SilentlyContinue | Where-Object {
        (Get-ItemProperty $_.PSPath -Name Path -ErrorAction SilentlyContinue).Path -eq "$Path\"
    }
    if ($already) {
        Write-Host "Excel $OfficeVersion already trusts $Path"
        return
    }

    $index = 0
    while (Test-Path (Join-Path $locationsKey "Location$index")) { $index++ }
    $newKey = Join-Path $locationsKey "Location$index"

    New-Item -Path $newKey -Force | Out-Null
    Set-ItemProperty -Path $newKey -Name Path -Value "$Path\"
    Set-ItemProperty -Path $newKey -Name AllowSubFolders -Value 1 -Type DWord
    Set-ItemProperty -Path $newKey -Name Date -Value (Get-Date -Format "MM/dd/yyyy HH:mm:ss")
    Set-ItemProperty -Path $newKey -Name Description -Value $trustDescription
    Write-Host "Trusted $Path for Excel $OfficeVersion"
}

# Excel reads HKCU:\...\Excel\Options\OPEN, OPEN1, OPEN2, ... at startup and
# loads whatever each one points to, the same as typing it on a command line.
# "/R <path>" is the documented switch for auto-loading an add-in this way.
function Find-ExistingAutoLoadSlot {
    param([string]$OptionsKey, [string]$XllPath)

    if (-not (Test-Path $OptionsKey)) { return $null }
    $props = Get-ItemProperty -Path $OptionsKey -ErrorAction SilentlyContinue
    if (-not $props) { return $null }

    foreach ($name in @("OPEN") + (1..50 | ForEach-Object { "OPEN$_" })) {
        $val = $props.$name
        if ($val -and $val.Contains($XllPath)) { return $name }
    }
    return $null
}

function Find-FreeAutoLoadSlot {
    param([string]$OptionsKey)

    $props = if (Test-Path $OptionsKey) { Get-ItemProperty -Path $OptionsKey -ErrorAction SilentlyContinue } else { $null }
    foreach ($name in @("OPEN") + (1..50 | ForEach-Object { "OPEN$_" })) {
        if (-not $props -or -not ($props.PSObject.Properties.Name -contains $name)) {
            return $name
        }
    }
    throw "No free Excel auto-load slot found (checked OPEN through OPEN50)."
}

function Add-ExcelAutoLoad {
    param([string]$OfficeVersion, [string]$XllPath)

    $optionsKey = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Excel\Options"
    New-Item -Path $optionsKey -Force | Out-Null

    $existing = Find-ExistingAutoLoadSlot -OptionsKey $optionsKey -XllPath $XllPath
    if ($existing) {
        Write-Host "Excel $OfficeVersion already auto-loads LudicrousSpeed ($existing)"
        return
    }

    $slot = Find-FreeAutoLoadSlot -OptionsKey $optionsKey
    Set-ItemProperty -Path $optionsKey -Name $slot -Value "/R `"$XllPath`""
    Write-Host "Registered LudicrousSpeed to auto-load for Excel $OfficeVersion ($slot)"
}

function Remove-ExcelAutoLoad {
    param([string]$OfficeVersion, [string]$XllPath)

    $optionsKey = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Excel\Options"
    $slot = Find-ExistingAutoLoadSlot -OptionsKey $optionsKey -XllPath $XllPath
    if ($slot) {
        Remove-ItemProperty -Path $optionsKey -Name $slot -Force
        Write-Host "Removed auto-load entry $slot (Excel $OfficeVersion)"
    }
}

if ($Uninstall) {
    $xllInstalled = Join-Path $installDir "LudicrousSpeed.xll"
    foreach ($ver in Get-ExcelVersionKeys) {
        try {
            Remove-TrustedLocation -OfficeVersion $ver
        } catch {
            Write-Warning "Could not clean up Excel $ver trusted locations: $_"
        }
        try {
            Remove-ExcelAutoLoad -OfficeVersion $ver -XllPath $xllInstalled
        } catch {
            Write-Warning "Could not clean up Excel $ver auto-load entry: $_"
        }
    }

    if (Test-Path $installDir) {
        Remove-Item $installDir -Recurse -Force
        Write-Host "Removed $installDir"
    }

    Write-Host ""
    Write-Host "Uninstalled. If LudicrousSpeed is still listed under Excel > File > Options > Add-ins, remove it from there too."
    return
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$xllSource = Join-Path $scriptDir "LudicrousSpeed.xll"
$dllSource = Join-Path $scriptDir "ludicrous_engine.dll"

foreach ($f in @($xllSource, $dllSource)) {
    if (-not (Test-Path $f)) {
        throw "Expected file not found next to this script: $f`nRun this script from inside the unzipped LudicrousSpeed folder."
    }
}

New-Item -ItemType Directory -Force -Path $installDir | Out-Null

# Copy every .dll in the bundle, not just the engine. The packed .xll does not
# embed the add-in's managed dependencies under .NET Framework, so
# Newtonsoft.Json ships beside it -- and installing only the two named files
# left it behind, producing "Could not load file or assembly 'Newtonsoft.Json'"
# the moment Excel ran anything. Copying whatever the bundle contains means a
# new dependency needs no change here.
Copy-Item $xllSource $installDir -Force
foreach ($dll in Get-ChildItem $scriptDir -Filter "*.dll" -File) {
    Copy-Item $dll.FullName $installDir -Force
}
Get-ChildItem $installDir -File | Unblock-File

# The add-in cannot start without this one, and the failure surfaces as an
# assembly-load dialog in Excel rather than anything pointing back here.
$jsonInstalled = Join-Path $installDir "Newtonsoft.Json.dll"
if (-not (Test-Path $jsonInstalled)) {
    Write-Warning "Newtonsoft.Json.dll was not in this bundle. Excel will fail to load the add-in with an assembly-load error. Re-download the release zip; if it is genuinely missing from the zip, the packaging step is at fault."
}

$xllInstalled = Join-Path $installDir "LudicrousSpeed.xll"
Write-Host "Copied LudicrousSpeed to $installDir"

$versions = @(Get-ExcelVersionKeys)
if ($versions.Count -eq 0) {
    Write-Warning "Could not detect an Excel installation under HKCU:\Software\Microsoft\Office. Skipping Trust Center and add-in setup -- add the folder as a trusted location and load the add-in manually."
}

foreach ($ver in $versions) {
    try {
        Add-TrustedLocation -OfficeVersion $ver -Path $installDir
    } catch {
        Write-Warning "Could not register a trusted location for Excel $ver -- $_"
        Write-Warning "Add it manually: File > Options > Trust Center > Trust Center Settings > Trusted Locations > Add new location > $installDir"
    }

    try {
        Add-ExcelAutoLoad -OfficeVersion $ver -XllPath $xllInstalled
    } catch {
        Write-Warning "Could not register LudicrousSpeed to auto-load for Excel $ver -- $_"
    }
}

Write-Host ""
Write-Host "Install finished. Open Excel -- LudicrousSpeed should already be on the ribbon."
Write-Host ""
Write-Host "If you don't see it, add it manually:"
Write-Host "  File > Options > Add-ins > Manage: Excel Add-ins > Go... > Browse..."
Write-Host "  Select: $xllInstalled"
Write-Host ""
Write-Host "If Windows shows a SmartScreen warning first, click 'More info' then 'Run anyway' -- this build isn't code-signed yet."
