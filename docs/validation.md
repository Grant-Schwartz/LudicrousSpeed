# Validation Plan

LudicrousSpeed should earn trust in layers:

1. Run IronCalc against synthetic `.xlsx` fixtures through `ludicrous-cli`.
2. Compare IronCalc results to Excel full rebuild results in `BenchmarkWorkbook`.
3. Enable cached value writeback only for regions with strict Excel parity.
4. Expand formula coverage based on fallback reports from real sanitized models.

## First fixture set

- `operating-model-basic.xlsx`
- `debt-schedule-basic.xlsx`
- `returns-analysis-basic.xlsx`
- `lookup-assumptions-basic.xlsx`
- `sensitivity-table-basic.xlsx`
- `unsupported-vba-udf.xlsx`
- `unsupported-external-link.xlsx`
- `unsupported-circular-reference.xlsx`

## Acceptance gate for writeback

Before writing values into the live workbook, the same cells must pass:

- formula remains unchanged;
- value matches Excel within numeric tolerance;
- no fallback reason touches the dependency region;
- workbook can be restored from the add-in snapshot.
