# Windows Build and Test

## Prerequisites

- 64-bit Windows with desktop Excel installed.
- Rust MSVC toolchain, ideally `stable-x86_64-pc-windows-msvc`.
- Visual Studio Build Tools with the C++ MSVC toolchain.
- .NET SDK plus the .NET Framework 4.8 Developer Pack, because the add-in targets `net48`.

## Build

From the repository root in PowerShell:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\build_windows.ps1 -Configuration Release
```

The script runs Rust tests, builds `target\release\warpspeed_engine.dll`,
builds the Excel-DNA add-in, copies the native DLL beside the add-in output, and
prints the `.xll` path to load in Excel.

For a faster compile while iterating:

```powershell
.\scripts\build_windows.ps1 -Configuration Debug -SkipTests
```

## CLI Smoke Test

Use the edited-assumption benchmark to avoid measuring only a no-op cache hit:

```powershell
cargo run -p warpspeed-engine --bin warpspeed-cli -- outputs\lbo-fixtures\ALMS_v11.xlsx benchmark --eval-data-tables --edit "LBO (Share Price)!G9=0.25"
```

Expected shape:

- Run 1 is `cold_full` and pays workbook load cost.
- Run 2 is `warm_full_with_dirty_plan`.
- Run 2 should show `model=true`, `graph=true`, and `result=false`.
- Run 2 should show dirty formulas and dirty data tables for the purchase
  premium change.

## Excel Acceptance Test

1. Add the add-in output folder as a trusted location in Excel.
2. Load the generated `.xll` through `File > Options > Add-ins > Excel Add-ins > Browse`.
3. Open `outputs\lbo-fixtures\ALMS_v11.xlsx`.
4. Click `WarpSpeed > Benchmark` once to warm the engine.
5. Change `LBO (Share Price)!G9` from `20.0%` to `25.0%`.
6. Click `WarpSpeed > Benchmark` again.
7. Check `_WarpSpeed_Report`.

The second report should show a warm dirty run, not a no-change result cache hit:

- `Strategy` = `warm_full_with_dirty_plan`
- `Model cache hit` = `TRUE`
- `Graph cache hit` = `TRUE`
- `Result cache hit` = `FALSE`
- `Dirty data tables` > `0`
- `Snapshot skipped` should be `TRUE` when change tracking is certain.

## Live Formula-Cache Writeback Acceptance Test

Use a simple workbook first so fallback regions do not intentionally block the
MVP:

1. Create a workbook with `A1 = 100`, `A2 = 25`, and `B1 = SUM(A1:A2)`.
2. Click `WarpSpeed > Recalculate with WarpSpeed`.
3. Check `_WarpSpeed_Report`.

Expected shape for the gated MVP:

- `Writeback mode` = `live_formula_cache`
- `Candidate cells` > `0`
- `Host writeback status` = `blocked` unless a supported formula-cache setter
  passes the scratch-workbook probe.
- `Writeback skipped reasons` includes `probe_blocked` when blocked.
- The model formula cell still contains its original formula.

If a future supported setter passes the probe, the same smoke test should show
`Written cells` > `0`, Excel should remain in manual calculation mode, and
`Restore Last Results` should run a full Excel rebuild and restore the previous
calculation mode.

## Notes

- `warpspeed_engine.dll` must sit beside the Excel add-in output because
  `NativeEngineClient` imports that exact DLL name.
- Use 64-bit Excel with the 64-bit Rust build. If testing on 32-bit Excel, build
  an `i686-pc-windows-msvc` native DLL and use the matching 32-bit Excel-DNA
  add-in output.
- If Windows blocks the downloaded add-in, run `Unblock-File` on the `.xll` and
  `warpspeed_engine.dll`, or use a trusted location.
