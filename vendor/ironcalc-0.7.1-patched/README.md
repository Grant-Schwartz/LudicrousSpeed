# ironcalc 0.7.1 (patched)

This is a trimmed, patched vendor copy of the `ironcalc` crate (library
target only — bins, examples, tests, and doc assets from the upstream
package were removed since we only depend on it as a library).

## Why this fork exists

Upstream `ironcalc`'s xlsx importer (`src/import/worksheets.rs`,
`load_sheet`) calls `from_a1_to_rc` once per formula cell to normalize
A1 references to RC form. That function constructs a brand new
`Parser` per call via `new_parser_english(worksheets, defined_names,
tables)`, which **deep-clones the full `defined_names` list on every
formula cell**.

On a real ~35k-formula LBO model with ~27,700 defined names (common in
real-world models after repeated Bloomberg/CapIQ/FactSet paste-special
operations, which inject large numbers of junk named ranges), this
made workbook import take about 49 seconds — dominated entirely by
that per-cell clone, not by any actual parsing or I/O work. Still
present in the latest release (0.8.3) as of 2026-08-03; not yet
reported upstream.

## What changed

`load_sheet` now builds one `Parser` per **sheet** instead of one per
**formula cell**, and calls `parser.parse(&formula, &cell_reference)`
directly instead of going through `from_a1_to_rc` per cell.
`Parser::parse` is designed to be called repeatedly on one instance —
it resets its own lexer state and cell context on every call — so this
changes nothing about parsing behavior, only how often the parser
(and the things it owns) gets constructed. Verified against the same
real model: import time drops from ~49s to ~2s with identical output.

The formula-string dedup lookup (`get_formula_index`, previously a
linear scan over `shared_formulas`) was also replaced with a
`HashMap<String, i32>` for O(1) lookup. Measured to make no real
difference on its own (`shared_formulas` per sheet stays small), but
it's a correctness-neutral improvement worth keeping.

Diff is scoped to `src/import/worksheets.rs`
(`from_a1_to_rc`/`get_formula_index` are now dead code, left in place
to keep the diff minimal and easy to compare against upstream).

## Removing this fork

Once this is fixed upstream and released, drop the
`[patch.crates-io]` entry in the workspace `Cargo.toml` and delete
this directory.
