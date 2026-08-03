# Architecture

WarpSpeed is split into an Excel host and a Rust engine.

## Excel host

The Excel-DNA add-in owns all live Excel interaction:

- ribbon commands and user feedback;
- temporary workbook snapshot creation;
- Excel baseline calculation timing;
- restore/writeback orchestration;
- fallback to Excel for unsupported behavior.

## Rust engine

The Rust engine owns workbook evaluation planning:

- load an `.xlsx` workbook with IronCalc;
- evaluate the model where IronCalc supports the workbook;
- return a structured calculation result;
- report unsupported features, formula coverage, timings, and cache metrics.

## Current bridge

The prototype bridge is file-based: Excel saves a temporary `.xlsx`, Rust loads
it, evaluates it, and returns JSON over a C ABI. This favors fast validation over
maximum performance. Once coverage and correctness are proven, direct in-memory
translation can replace the file bridge behind the same `CalcEngine` boundary.

For recalculation runs where the workbook has no fallback regions, Rust includes
typed formula-result candidates in the JSON response. The Excel host only applies
those candidates if a live formula-cache probe proves the available automation
path can update displayed values without replacing formulas; otherwise the
candidate cells remain report-only and the model workbook is left unchanged.

The V1 benchmark path keeps a process-local Rust engine alive behind the C ABI.
The Excel host tracks edited cells with `Application.SheetChange`; after a
successful cold snapshot run, warm runs can send only changed cell inputs and
skip `SaveCopyAs` when the sheet topology is unchanged. Rust reuses the loaded
IronCalc model and dependency graph, computes the dirty formula closure for
metrics, and still uses full IronCalc evaluation after edits. If no cells changed
and the graph has no volatile or unsupported barriers, Rust can return the cached
last result without evaluating.

## Data tables

OpenXML data table formulas are parsed before IronCalc import. WarpSpeed records
the table range, input cells, inferred source formula cell, and cached Excel
outputs, then strips the unsupported `dataTable` formula marker in a temporary
copy so the rest of the workbook can load. Data tables remain fallback regions
unless the caller explicitly requests Rust validation.

The V1.5 validator supports conventional two-variable sensitivity tables where
the source formula is diagonally above-left of the result body, column scenarios
sit immediately above the body, and row scenarios sit immediately to the left.
Validation uses a batched Rust formula-cone kernel for supported formulas such
as arithmetic, comparisons, `IF`, `SUM`, `MIN`, `MAX`, `INDEX`, `MATCH`, and
static `OFFSET` references. Tables are scheduled across worker threads, and each
supported table evaluates its scenario vector without mutating the whole
workbook or running full workbook calculation per scenario. Mismatches or
unsupported layouts stay visible in the benchmark report and are never silently
treated as correct.

## Correctness policy

Excel remains authoritative. IronCalc results are trusted only for supported
regions that pass validation. Unsupported formulas, external links, circular
references, VBA/UDFs, volatile semantics, and other hard cases are fallback
regions.
