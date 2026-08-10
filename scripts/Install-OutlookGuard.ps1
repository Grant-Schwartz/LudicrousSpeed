<#
Installs the LudicrousSpeed Outlook attachment guard for the current user only:
  1. copies LudicrousSpeed.OutlookAddIn.dll to %LOCALAPPDATA%\LudicrousSpeedOutlook
  2. unblocks it, so Windows stops treating it as downloaded-from-the-internet
  3. registers it as a per-user COM class, in both registry views, so it loads
     into 32-bit and 64-bit Outlook alike
  4. registers it as an Outlook add-in set to load at startup

Everything here is per-user (HKCU, %LOCALAPPDATA%) -- no admin rights, and no
regasm, which would want to write machine-wide.

Separate from the Excel add-in on purpose, in its own folder and its own
installer: either one can be removed without disturbing the other, and the
guard is useful to anyone who receives converted models even if they never
convert one themselves.

Prefer double-clicking InstallOutlookGuard.cmd / UninstallOutlookGuard.cmd over
running this directly; they set the execution policy for you.

Close Outlook first. Windows will not let the file be replaced or deleted
while Outlook has it loaded.

Usage:
  .\Install-OutlookGuard.ps1
  .\Install-OutlookGuard.ps1 -Uninstall
#>
param(
    [switch]$Uninstall,

    # Where the bundle's files are. Normally left empty and discovered below.
    # InstallOutlookGuard.cmd passes it explicitly because it runs this
    # script's *text* rather than the file, which leaves the usual
    # self-location properties empty -- see the comment on $scriptDir.
    [string]$BundleDir
)

$ErrorActionPreference = "Stop"

$installDir = Join-Path $env:LOCALAPPDATA "LudicrousSpeedOutlook"
$dllName = "LudicrousSpeed.OutlookAddIn.dll"

# These three have to match the source exactly:
#   [Guid] and [ProgId] on LudicrousSpeedOutlookAddIn, and its full type name.
# Everything else about the registration is derived from the built file.
$clsid = "{996BCB68-6439-44E2-84F0-9DF0FC563205}"
$progId = "LudicrousSpeed.OutlookAddIn"
$className = "LudicrousSpeed.OutlookAddIn.LudicrousSpeedOutlookAddIn"
$friendlyName = "LudicrousSpeed Attachment Guard"
$description = "Warns before a workbook with LudicrousSpeed live cells is mailed out."

# The category regasm stamps on managed classes. Nothing reads it to decide
# whether to load us, but tools that audit COM registrations do, and its
# absence makes a hand-written registration look like malware doing the same
# thing.
$managedCategory = "{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"

$addinKeyPath = "Software\Microsoft\Office\Outlook\Addins\$progId"
$classesRoot = "Software\Classes"

# Both views, so the registration is visible whichever bitness Outlook is,
# regardless of the bitness of the PowerShell running this. HKCU\Software\
# Classes is one of the redirected hives, so writing it once is not enough --
# and the .NET view API is used rather than an HKCU: path with Wow6432Node
# spelled into it, because that spelling behaves differently depending on the
# caller's own bitness.
function Get-RegistryViews {
    if ([Environment]::Is64BitOperatingSystem) {
        return @(
            [Microsoft.Win32.RegistryView]::Registry64,
            [Microsoft.Win32.RegistryView]::Registry32)
    }

    return @([Microsoft.Win32.RegistryView]::Default)
}

function Set-KeyValues {
    param(
        [Microsoft.Win32.RegistryKey]$BaseKey,
        [string]$Path,
        [hashtable]$Values
    )

    $key = $BaseKey.CreateSubKey($Path)
    try {
        foreach ($name in $Values.Keys) {
            $key.SetValue($name, $Values[$name])
        }
    } finally {
        $key.Dispose()
    }
}

function Add-ComRegistration {
    param([string]$DllPath)

    $assembly = [System.Reflection.AssemblyName]::GetAssemblyName($DllPath)
    $codeBase = "file:///" + $DllPath.Replace('\', '/')

    # The shim, not our DLL: Outlook is a native host, so the InprocServer32 it
    # loads is the CLR's, and the values beside it tell the CLR which managed
    # class to hand back.
    $shim = Join-Path $env:SystemRoot "System32\mscoree.dll"

    $serverValues = @{
        "Class"          = $className
        "Assembly"       = $assembly.FullName
        "RuntimeVersion" = "v4.0.30319"
        "CodeBase"       = $codeBase
    }

    foreach ($view in Get-RegistryViews) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', $view)
        try {
            Set-KeyValues $base "$classesRoot\CLSID\$clsid" @{ "" = $friendlyName }
            Set-KeyValues $base "$classesRoot\CLSID\$clsid\ProgID" @{ "" = $progId }
            Set-KeyValues $base `
                "$classesRoot\CLSID\$clsid\Implemented Categories\$managedCategory" @{}

            $inproc = $serverValues.Clone()
            $inproc[""] = $shim
            $inproc["ThreadingModel"] = "Both"
            Set-KeyValues $base "$classesRoot\CLSID\$clsid\InprocServer32" $inproc

            # regasm also writes a subkey named for the assembly version, and
            # the shim looks there first. Without it, a later build with a
            # bumped version would be ignored in favour of whatever the
            # unversioned values still say.
            Set-KeyValues $base `
                "$classesRoot\CLSID\$clsid\InprocServer32\$($assembly.Version)" $serverValues

            Set-KeyValues $base "$classesRoot\$progId" @{ "" = $friendlyName }
            Set-KeyValues $base "$classesRoot\$progId\CLSID" @{ "" = $clsid }

            Write-Host "Registered COM class in the $view view"
        } finally {
            $base.Dispose()
        }
    }
}

function Remove-ComRegistration {
    foreach ($view in Get-RegistryViews) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', $view)
        try {
            foreach ($path in @("$classesRoot\CLSID\$clsid", "$classesRoot\$progId")) {
                try {
                    $base.DeleteSubKeyTree($path, $false)
                } catch {
                    Write-Warning "Could not remove $path from the $view view -- $_"
                }
            }
        } finally {
            $base.Dispose()
        }
    }
}

# LoadBehavior 3 is "connect at startup, and stay connected". Outlook rewrites
# it to 2 or 0 when an add-in throws during load or is switched off by hand, so
# re-running the installer is the documented way to turn it back on.
function Add-OutlookAddIn {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        'CurrentUser', [Microsoft.Win32.RegistryView]::Default)
    try {
        Set-KeyValues $base $addinKeyPath @{
            "FriendlyName"    = $friendlyName
            "Description"     = $description
            "LoadBehavior"    = 3
            "CommandLineSafe" = 0
        }
        Write-Host "Registered $friendlyName to load with Outlook"
    } finally {
        $base.Dispose()
    }
}

function Remove-OutlookAddIn {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        'CurrentUser', [Microsoft.Win32.RegistryView]::Default)
    try {
        $base.DeleteSubKeyTree($addinKeyPath, $false)
    } catch {
        Write-Warning "Could not remove the Outlook add-in entry -- $_"
    } finally {
        $base.Dispose()
    }
}

function Test-OutlookRunning {
    return @(Get-Process -Name "outlook" -ErrorAction SilentlyContinue).Count -gt 0
}

if (Test-OutlookRunning) {
    Write-Warning "Outlook is running. It holds the add-in file open, so this may fail to copy or delete it. Close Outlook and run this again."
    Write-Host ""
}

if ($Uninstall) {
    Remove-OutlookAddIn
    Remove-ComRegistration

    if (Test-Path $installDir) {
        try {
            Remove-Item $installDir -Recurse -Force
            Write-Host "Removed $installDir"
        } catch {
            Write-Warning "Could not remove $installDir -- $_"
            Write-Warning "The registry entries are gone, so Outlook will not load it again. Delete the folder once Outlook is closed."
        }
    }

    Write-Host ""
    Write-Host "Uninstalled. Restart Outlook to unload it from the running session."
    return
}

# Install.cmd-style invocation runs this script's text, not the file, so
# $PSScriptRoot and $MyInvocation.MyCommand.Path are empty, hence -BundleDir.
$scriptDir = if ($BundleDir) {
    $BundleDir
} elseif ($PSScriptRoot) {
    $PSScriptRoot
} else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

if (-not $scriptDir) {
    throw "Could not determine which folder this script is running from. Pass -BundleDir with the path to the unzipped LudicrousSpeed folder."
}

$dllSource = Join-Path $scriptDir $dllName
if (-not (Test-Path $dllSource)) {
    throw "Expected file not found next to this script: $dllSource`nRun this script from inside the unzipped LudicrousSpeed folder."
}

New-Item -ItemType Directory -Force -Path $installDir | Out-Null
Copy-Item $dllSource $installDir -Force
Get-ChildItem $installDir -File | Unblock-File

$dllInstalled = Join-Path $installDir $dllName
Write-Host "Copied the attachment guard to $installDir"

Add-ComRegistration -DllPath $dllInstalled
Add-OutlookAddIn

Write-Host ""
Write-Host "Install finished. Start Outlook -- attach a workbook that still has"
Write-Host "LudicrousSpeed live cells in it and you should be warned about it."
Write-Host ""
Write-Host "If nothing happens:"
Write-Host "  File > Options > Add-ins. If '$friendlyName' is under 'Inactive'"
Write-Host "  or 'Disabled Application Add-ins', select 'COM Add-ins' next to Manage,"
Write-Host "  click Go... and tick it. Outlook disables add-ins that fail once."
Write-Host ""
Write-Host "  It also keeps a log at %LOCALAPPDATA%\LudicrousSpeed\outlook-guard.log"
Write-Host "  which records anything it skipped and why."
