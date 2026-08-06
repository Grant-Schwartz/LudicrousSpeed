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

The script runs Rust tests, builds `target\release\ludicrous_engine.dll`,
builds the Excel-DNA add-in, copies the native DLL beside the add-in output, and
prints the `.xll` path to load in Excel.

For a faster compile while iterating:

```powershell
.\scripts\build_windows.ps1 -Configuration Debug -SkipTests
```

## CLI Smoke Test

Use the edited-assumption benchmark to avoid measuring only a no-op cache hit:

```powershell
cargo run -p ludicrous-engine --bin ludicrous-cli -- outputs\lbo-fixtures\ALMS_v11.xlsx benchmark --eval-data-tables --edit "LBO (Share Price)!G9=0.25"
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
4. Click `LudicrousSpeed > Benchmark` once to warm the engine.
5. Change `LBO (Share Price)!G9` from `20.0%` to `25.0%`.
6. Click `LudicrousSpeed > Benchmark` again.
7. Check `_LudicrousSpeed_Report`.

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
2. Click `LudicrousSpeed > Recalculate with LudicrousSpeed`.
3. Check `_LudicrousSpeed_Report`.

Expected shape for the gated MVP:

- `Writeback mode` = `live_formula_cache`
- `Candidate cells` > `0`
- `Host writeback status` = `blocked` unless a supported formula-cache setter
  passes the scratch-workbook probe.
- `Writeback skipped reasons` includes `probe_blocked` when blocked.
- The model formula cell still contains its original formula.

### xlSet mechanism (unverified)

`LiveFormulaCacheProbe` now tries the XLL C API's `xlSet` before falling back
to checking (and, as always, rejecting) COM `Value2`. This was written and
reasoned through against the documented XLL SDK semantics but has not been
built or run against live Excel -- there was no Windows/Excel available to
verify it. Before trusting it:

1. Confirm `dotnet build` succeeds -- first check that `XlCall.xlSet`,
   `XlCall.xlSheetId`, and the five-argument `ExcelReference` constructor
   used in `LiveFormulaCacheProbe.BuildReference` actually compile and match
   this project's pinned `ExcelDna.Integration` version's API shape.
2. Run the smoke test above. Watch specifically for:
   - `Live formula-cache probe` note in the report should say `xl_set`
     succeeded, not `com_value2` or the final "no supported" message.
   - `Host writeback status` = `applied` (not `blocked`), `Written cells` > 0.
   - The model formula cell (`B1`) must still show `=SUM(A1:A2)` in the
     formula bar after writeback, not a literal number.
3. Force a `#DIV/0!`-style formula elsewhere and confirm `xlSet` doesn't
   throw or corrupt unrelated cells; also confirm calling it from the ribbon
   callback context doesn't throw an `XlCallException` about invalid calling
   context (a real risk noted in `LiveFormulaCacheProbe`'s comments -- some
   XLL C API functions are only valid from specific Excel-DNA execution
   contexts).
4. Click `Restore Last Results` and confirm Excel runs a full rebuild and the
   calculation mode is restored to what it was before the run.
5. Repeat on a large real workbook (e.g. `ALMS_v11.xlsx`) with
   `--eval-data-tables` on, and confirm Excel doesn't hang or show stale
   values in cells `xlSet` touched after a manual F9 recalculation --
   per the XLL SDK, `xlSet`'s value on a worksheet cell is expected to be a
   temporary display value that a real recalculation overwrites, which is the
   intended behavior here, but is worth seeing happen for real.

If `xlSet` doesn't pass the probe or doesn't behave as documented, the system
degrades safely to today's behavior (`probe_blocked`, no cells written) --
this is why it's implemented as a probed, try-first mechanism rather than an
assumed one.

## In-Memory Snapshot Acceptance Test (Unverified Feature)

`InMemoryWorkbookReader` and its wiring in `WorkbookSnapshotService` were
written without access to Windows or Excel and have not been built or run
against live Excel. Before trusting `LUDICROUS_INLINE_SNAPSHOT=1` for
anything beyond experimentation, verify:

1. `dotnet build` / MSBuild succeeds — first confirm the COM member accesses
   (`Range.Formula`, `Workbook.Names`, `Name.RefersTo`) actually compile
   against this project's pinned `ExcelDna.Interop` version.
2. Set the environment variable, open a workbook covering: plain numbers,
   text, booleans, dates, at least one formula error (e.g. `=1/0`), a
   workbook-scoped defined name, and a sheet-scoped defined name.
3. Run `Analyze Workbook` once with the variable set and once unset (file-based
   path). Diff the two `_LudicrousSpeed_Report` sheets' coverage, fallback, and
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

`LudicrousSpeedRibbon.RunAsync` (native engine call off Excel's UI thread via
`ExcelAsyncUtil.QueueAsMacro`) was also written without access to Windows or
Excel. The `ExcelAsyncUtil.QueueAsMacro` API itself was confirmed to exist in
the pinned `ExcelDna.Integration` 1.9.0 package by inspecting the DLL, but the
threading behavior has not been exercised. Before trusting
`LUDICROUS_ASYNC_RUN=1` for anything beyond experimentation, verify:

1. `dotnet build` / MSBuild succeeds.
2. On a large workbook (e.g. `ALMS_v11.xlsx`), click `Recalculate with
   LudicrousSpeed` with the variable set. Confirm Excel's UI stays responsive
   (you can click cells, switch sheets, see the status bar message) while the
   native call is in flight, rather than showing "Not Responding."
3. Confirm the completion dialog, `_LudicrousSpeed_Report` sheet, and writeback
   results are identical to a run with the variable unset (same workbook,
   same inputs) — this flag should only change *when* work happens, not
   *what* happens.
4. Force a failure (e.g. temporarily rename `ludicrous_engine.dll`) and
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
   `LudicrousSpeedRibbon.AsyncRunEnabled`.

## Notes

- `ludicrous_engine.dll` must sit beside the Excel add-in output because
  `NativeEngineClient` imports that exact DLL name.
- Use 64-bit Excel with the 64-bit Rust build. If testing on 32-bit Excel, build
  an `i686-pc-windows-msvc` native DLL and use the matching 32-bit Excel-DNA
  add-in output.
- If Windows blocks the downloaded add-in, run `Unblock-File` on the `.xll` and
  `ludicrous_engine.dll`, or use a trusted location.
