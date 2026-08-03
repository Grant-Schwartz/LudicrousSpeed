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

## In-Memory Snapshot Acceptance Test (Unverified Feature)

`InMemoryWorkbookReader` and its wiring in `WorkbookSnapshotService` were
written without access to Windows or Excel and have not been built or run
against live Excel. Before trusting `WARPSPEED_INLINE_SNAPSHOT=1` for
anything beyond experimentation, verify:

1. `dotnet build` / MSBuild succeeds — first confirm the COM member accesses
   (`Range.Formula`, `Workbook.Names`, `Name.RefersTo`) actually compile
   against this project's pinned `ExcelDna.Interop` version.
2. Set the environment variable, open a workbook covering: plain numbers,
   text, booleans, dates, at least one formula error (e.g. `=1/0`), a
   workbook-scoped defined name, and a sheet-scoped defined name.
3. Run `Analyze Workbook` once with the variable set and once unset (file-based
   path). Diff the two `_WarpSpeed_Report` sheets' coverage, fallback, and
   fallback-detail sections — they should match exactly except for timing.
4. Repeat step 3 on `outputs\lbo-fixtures\ALMS_v11.xlsx` (or another large real
   model) and compare `Rust load ms` between the two runs — this is the number
   the feature exists to shrink.
5. Try a workbook containing a native two-input data table and confirm it
   still falls back correctly (this path doesn't detect data tables yet; see
   `docs/architecture.md`).
6. Only after 1–5 pass, consider flipping the default in
   `WorkbookSnapshotService.InlineSnapshotEnabled`.

## Async Ribbon Acceptance Test (Unverified Feature)

`WarpSpeedRibbon.RunAsync` (native engine call off Excel's UI thread via
`ExcelAsyncUtil.QueueAsMacro`) was also written without access to Windows or
Excel. The `ExcelAsyncUtil.QueueAsMacro` API itself was confirmed to exist in
the pinned `ExcelDna.Integration` 1.9.0 package by inspecting the DLL, but the
threading behavior has not been exercised. Before trusting
`WARPSPEED_ASYNC_RUN=1` for anything beyond experimentation, verify:

1. `dotnet build` / MSBuild succeeds.
2. On a large workbook (e.g. `ALMS_v11.xlsx`), click `Recalculate with
   WarpSpeed` with the variable set. Confirm Excel's UI stays responsive
   (you can click cells, switch sheets, see the status bar message) while the
   native call is in flight, rather than showing "Not Responding."
3. Confirm the completion dialog, `_WarpSpeed_Report` sheet, and writeback
   results are identical to a run with the variable unset (same workbook,
   same inputs) — this flag should only change *when* work happens, not
   *what* happens.
4. Force a failure (e.g. temporarily rename `warpspeed_engine.dll`) and
   confirm: an error dialog still appears, the status bar clears, and —
   importantly — Excel's calculation mode is correctly restored rather than
   left stuck on Manual (check `ExcelCalculationGuard`'s effect via
   Formulas > Calculation Options after the failed run).
5. Click the ribbon button twice in quick succession and confirm nothing
   deadlocks or corrupts the report sheet (there's currently no
   re-entrancy guard against overlapping runs in either the sync or async
   path — this is a pre-existing gap, not new, but async runs make it easier
   to trigger by clicking again before the first run's dialog appears).
6. Only after 1–5 pass, consider flipping the default in
   `WarpSpeedRibbon.AsyncRunEnabled`.

## Notes

- `warpspeed_engine.dll` must sit beside the Excel add-in output because
  `NativeEngineClient` imports that exact DLL name.
- Use 64-bit Excel with the 64-bit Rust build. If testing on 32-bit Excel, build
  an `i686-pc-windows-msvc` native DLL and use the matching 32-bit Excel-DNA
  add-in output.
- If Windows blocks the downloaded add-in, run `Unblock-File` on the `.xll` and
  `warpspeed_engine.dll`, or use a trusted location.
