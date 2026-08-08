# LudicrousSpeed

LudicrousSpeed is a Windows desktop Excel accelerator prototype for finance workbooks.
It provides an Excel-DNA ribbon add-in that snapshots the active workbook, calls a
Rust calculation engine backed by IronCalc, and reports coverage, fallback, and
benchmark diagnostics.

## Installing a beta build

Prebuilt Windows builds are published on the
[Releases page](https://github.com/Grant-Schwartz/LudicrousSpeed/releases/latest).
See [INSTALL.md](INSTALL.md) -- no build tools required.

## Repository layout

- `crates/ludicrous-engine` - Rust engine and C ABI used by the Excel add-in.
- `excel-addin/LudicrousSpeed.ExcelAddIn` - Excel-DNA host, ribbon commands, workbook
  snapshot/export, native engine calls, and user reports.
- `scripts` - Windows build/package scripts and the beta installer.
- `fixtures` - Synthetic workbook scenarios and fixture notes.
- `docs` - Architecture and validation notes.
- `site` - Source for the GitHub Pages landing page.

## Flow

1. User clicks `Analyze Workbook`, `Recalculate with LudicrousSpeed`, or
   `Benchmark` in Excel.
2. The add-in saves a temporary `.xlsx` copy of the active workbook (skipped on
   warm runs, which send only the changed cells).
3. Rust loads the workbook with IronCalc and evaluates the model.
4. Rust returns a computed value for every cell it can vouch for.
5. The add-in publishes those values to `LS.LIVE` cells over RTD, then lets
   Excel propagate downstream from each one.
6. Excel remains the fallback and correctness authority for unsupported
   behavior.

## How values reach the sheet

Excel has no way to set a formula cell's cached value in place. `Range.Value2`
replaces the formula outright; the XLL C API's `xlSet` is macro-sheet-only and
silently does nothing to a worksheet cell; and setting `Value2` then restoring
the formula re-evaluates it. All three were tested against live Excel and
failed, each for a different reason. This is a property of Excel treating a
formula cell's formula and value as one unit owned by its calc engine, which is
why RTD is the mechanism every market-data vendor uses.

So values arrive through `=LS.LIVE("Sheet!Cell")` cells, backed by an RTD
server. An RTD value landing marks its dependents dirty, so Excel recalculates
downstream normally -- which means a handful of live cells trigger a whole
workbook refresh rather than only updating themselves.

The cells worth wiring up are the ones where Excel does *repeated* work.
Injecting into an ordinary formula in a linear chain saves nothing, because
Excel still has to evaluate that cell's precedents. Two structures do have
multiplicative cost:

- **Data tables** - Excel re-evaluates the source formula's dependency cone
  once per scenario cell, so a 5x5 two-variable table costs 25 passes.
  `Convert to Live` replaces the native table with `LS.LIVE` cells and the
  kernel computes the grid in one parallel pass. `Restore Native` puts the
  original table back from definitions recorded at conversion time.
- **Circular components** - Excel iterates the region to convergence.

## F9

The `F9 Uses LudicrousSpeed` toggle routes F9 to the engine instead of Excel's
calculation, so the accelerator sits on the keystroke people already press. The
setting persists between sessions.

It is off by default and has not been verified against live Excel. F9 has a
second job -- pressed in the formula bar with a sub-expression selected, it
evaluates that fragment in place -- and confirming that still works is the first
item in the acceptance test in `docs/windows-testing.md`. Shift+F9, Ctrl+Alt+F9
and Ctrl+Alt+Shift+F9 stay native on purpose, so Excel's own answer is always one
keystroke away.

## Upgrading from WS.LIVE

`LS.LIVE` is the only live-cell function; the earlier `WS.LIVE` name is gone.
A workbook converted before the rename holds `=WS.LIVE(...)` formulas, which
now resolve to `#NAME?`. Either re-run `Convert to Live`, or fix them in place
with Find & Replace across the whole workbook, searching in Formulas:

```
Find:    WS.LIVE(
Replace: LS.LIVE(
```

## Build prerequisites

- Windows desktop Excel.
- .NET SDK / MSBuild compatible with Excel-DNA.
- Rust toolchain.

## Rust engine

```sh
cargo test -p ludicrous-engine
cargo build -p ludicrous-engine --release
cargo run -p ludicrous-engine --bin ludicrous-cli -- path/to/workbook.xlsx analyze
cargo run -p ludicrous-engine --bin ludicrous-cli -- path/to/workbook.xlsx benchmark --eval-data-tables
cargo run -p ludicrous-engine --bin ludicrous-cli -- path/to/workbook.xlsx benchmark --runs 2
cargo run -p ludicrous-engine --bin ludicrous-cli -- path/to/workbook.xlsx benchmark --eval-data-tables --edit 'Assumptions!C12=125.0'
```

The regression tests generate temporary `.xlsx` workbooks with IronCalc, then
load them back through the same engine path used by the Excel add-in.

By default, the CLI prints a compact summary. Add `--json` to print the full
engine response, and `--runs N` to benchmark warm in-process cache reuse after
the first cold load. Add `--edit 'Sheet!A1=value'` to run a cold baseline first,
then apply that changed cell through the warm cache on the second run without
reloading the workbook from disk. Data table recomputation is opt-in with
`--eval-data-tables`: LudicrousSpeed parses OpenXML data tables, preserves Excel
cached values, and uses a batched Rust formula-cone kernel to validate
conventional two-variable sensitivity tables before treating Rust-computed table
outputs as trustworthy.

## Excel add-in

```sh
dotnet build excel-addin/LudicrousSpeed.ExcelAddIn/LudicrousSpeed.ExcelAddIn.csproj
```

Copy the built `ludicrous_engine` native library beside the packed Excel-DNA
add-in before loading it into Excel.

For a full Windows build-and-test flow, see `docs/windows-testing.md`.
