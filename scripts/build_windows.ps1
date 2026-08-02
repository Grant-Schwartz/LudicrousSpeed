param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$addinProject = Join-Path $repoRoot "excel-addin\WarpSpeed.ExcelAddIn\WarpSpeed.ExcelAddIn.csproj"
$addinOutput = Join-Path $repoRoot "excel-addin\WarpSpeed.ExcelAddIn\bin\$Configuration\net48"

if ($Configuration -eq "Release") {
    $cargoBuildArgs = @("build", "-p", "warpspeed-engine", "--release")
    $targetFolder = "release"
} else {
    $cargoBuildArgs = @("build", "-p", "warpspeed-engine")
    $targetFolder = "debug"
}

$engineDll = Join-Path $repoRoot "target\$targetFolder\warpspeed_engine.dll"

Write-Host "WarpSpeed Windows build"
Write-Host "Repo: $repoRoot"
Write-Host "Configuration: $Configuration"
Write-Host ""

if (-not $SkipTests) {
    Write-Host "Running Rust tests..."
    cargo test -p warpspeed-engine
    Write-Host ""
}

Write-Host "Building Rust native engine..."
cargo @cargoBuildArgs

if (-not (Test-Path $engineDll)) {
    throw "Expected native engine DLL was not produced: $engineDll"
}

Write-Host ""
Write-Host "Building Excel add-in..."
dotnet build $addinProject -c $Configuration

if (-not (Test-Path $addinOutput)) {
    throw "Expected add-in output directory was not produced: $addinOutput"
}

Write-Host ""
Write-Host "Copying native engine DLL beside add-in output..."
Copy-Item $engineDll $addinOutput -Force

$xlls = Get-ChildItem $addinOutput -Filter "*.xll" -ErrorAction SilentlyContinue
$dna = Get-ChildItem $addinOutput -Filter "*.dna" -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Build complete."
Write-Host "Native DLL:"
Write-Host "  $engineDll"
Write-Host ""
Write-Host "Add-in output:"
Write-Host "  $addinOutput"

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
