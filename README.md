# WarpSpeed

WarpSpeed is a Windows desktop Excel accelerator prototype for finance workbooks.
It provides an Excel-DNA ribbon add-in that snapshots the active workbook, calls a
Rust calculation engine backed by IronCalc, and reports coverage, fallback, and
benchmark diagnostics.

## Repository layout

- `crates/warpspeed-engine` - Rust engine and C ABI used by the Excel add-in.
- `excel-addin/WarpSpeed.ExcelAddIn` - Excel-DNA host, ribbon commands, workbook
  snapshot/export, native engine calls, and user reports.
- `fixtures` - Synthetic workbook scenarios and fixture notes.
- `docs` - Architecture and validation notes.

## Prototype flow

1. User clicks `Analyze Workbook`, `Recalculate with WarpSpeed`, or
   `Benchmark` in Excel.
2. The add-in saves a temporary `.xlsx` copy of the active workbook.
3. Rust loads the workbook with IronCalc and evaluates the model.
4. For `Recalculate with WarpSpeed`, Rust returns formula-cache writeback
   candidates only when the workbook has no fallback regions.
5. The add-in probes whether Excel exposes a supported formula-preserving live
   cache setter. If the probe fails, formulas and model cells are left untouched.
6. The add-in displays coverage, timing, and writeback diagnostics.
7. Excel remains the fallback and correctness authority for unsupported behavior.

The MVP intentionally refuses to fake formula-cache writeback with `Range.Value2`,
because that replaces formula cells. It preserves formulas, reports candidate and
blocked-writeback details, and leaves full live writeback behind a host probe.

## Build prerequisites

- Windows desktop Excel.
- .NET SDK / MSBuild compatible with Excel-DNA.
- Rust toolchain.

## Rust engine

```sh
cargo test -p warpspeed-engine
cargo build -p warpspeed-engine --release
cargo run -p warpspeed-engine --bin warpspeed-cli -- path/to/workbook.xlsx analyze
cargo run -p warpspeed-engine --bin warpspeed-cli -- path/to/workbook.xlsx benchmark --eval-data-tables
cargo run -p warpspeed-engine --bin warpspeed-cli -- path/to/workbook.xlsx benchmark --runs 2
cargo run -p warpspeed-engine --bin warpspeed-cli -- path/to/workbook.xlsx benchmark --eval-data-tables --edit 'Assumptions!C12=125.0'
```

The regression tests generate temporary `.xlsx` workbooks with IronCalc, then
load them back through the same engine path used by the Excel add-in.

By default, the CLI prints a compact summary. Add `--json` to print the full
engine response, and `--runs N` to benchmark warm in-process cache reuse after
the first cold load. Add `--edit 'Sheet!A1=value'` to run a cold baseline first,
then apply that changed cell through the warm cache on the second run without
reloading the workbook from disk. Data table recomputation is opt-in with
`--eval-data-tables`: WarpSpeed parses OpenXML data tables, preserves Excel
cached values, and uses a batched Rust formula-cone kernel to validate
conventional two-variable sensitivity tables before treating Rust-computed table
outputs as trustworthy.

## Excel add-in

```sh
dotnet build excel-addin/WarpSpeed.ExcelAddIn/WarpSpeed.ExcelAddIn.csproj
```

Copy the built `warpspeed_engine` native library beside the packed Excel-DNA
add-in before loading it into Excel.

For a full Windows build-and-test flow, see `docs/windows-testing.md`.
