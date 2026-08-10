param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"

function Invoke-Native {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(ValueFromRemainingArguments = $true)]
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$FilePath failed with exit code $LASTEXITCODE."
    }
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$addinProject = Join-Path $repoRoot "excel-addin\LudicrousSpeed.ExcelAddIn\LudicrousSpeed.ExcelAddIn.csproj"
$addinOutput = Join-Path $repoRoot "excel-addin\LudicrousSpeed.ExcelAddIn\bin\$Configuration\net48"
$outlookProject = Join-Path $repoRoot "outlook-addin\LudicrousSpeed.OutlookAddIn\LudicrousSpeed.OutlookAddIn.csproj"
$outlookOutput = Join-Path $repoRoot "outlook-addin\LudicrousSpeed.OutlookAddIn\bin\$Configuration\net48"

if ($Configuration -eq "Release") {
    $cargoBuildArgs = @("build", "-p", "ludicrous-engine", "--release")
    $targetFolder = "release"
} else {
    $cargoBuildArgs = @("build", "-p", "ludicrous-engine")
    $targetFolder = "debug"
}

$engineDll = Join-Path $repoRoot "target\$targetFolder\ludicrous_engine.dll"

Write-Host "LudicrousSpeed Windows build"
Write-Host "Repo: $repoRoot"
Write-Host "Configuration: $Configuration"
Write-Host ""

if (-not $SkipTests) {
    Write-Host "Running Rust tests..."
    Invoke-Native "cargo" "test" "-p" "ludicrous-engine"
    Write-Host ""
}

Write-Host "Building Rust native engine..."
Invoke-Native "cargo" @cargoBuildArgs

if (-not (Test-Path $engineDll)) {
    throw "Expected native engine DLL was not produced: $engineDll"
}

Write-Host ""
Write-Host "Building Excel add-in..."
Invoke-Native "dotnet" "build" $addinProject "-c" $Configuration

if (-not (Test-Path $addinOutput)) {
    throw "Expected add-in output directory was not produced: $addinOutput"
}

Write-Host ""
Write-Host "Copying native engine DLL beside add-in output..."
Copy-Item $engineDll $addinOutput -Force

# The Outlook attachment guard shares nothing with the engine -- it reads saved
# workbook files, not live ones -- so it builds independently and its failure
# says nothing about the Excel add-in.
Write-Host ""
Write-Host "Building Outlook attachment guard..."
Invoke-Native "dotnet" "build" $outlookProject "-c" $Configuration

$outlookDll = Join-Path $outlookOutput "LudicrousSpeed.OutlookAddIn.dll"
if (-not (Test-Path $outlookDll)) {
    throw "Expected Outlook add-in DLL was not produced: $outlookDll"
}

$xlls = Get-ChildItem $addinOutput -Filter "*.xll" -ErrorAction SilentlyContinue
$dna = Get-ChildItem $addinOutput -Filter "*.dna" -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Build complete."
Write-Host "Native DLL:"
Write-Host "  $engineDll"
Write-Host ""
Write-Host "Add-in output:"
Write-Host "  $addinOutput"
Write-Host ""
Write-Host "Outlook attachment guard:"
Write-Host "  $outlookDll"

if ($xlls.Count -gt 0) {
    Write-Host ""
    Write-Host "Load one of these in Excel:"
    foreach ($xll in $xlls) {
        Write-Host "  $($xll.FullName)"
    }
} elseif ($dna.Count -gt 0) {
    Write-Host ""
    Write-Host "No .xll was found. Use ExcelDnaPack or load/check this .dna output:"
    foreach ($item in $dna) {
        Write-Host "  $($item.FullName)"
    }
} else {
    Write-Host ""
    Write-Host "No .xll or .dna was found in the add-in output. Inspect the dotnet build output above."
}
