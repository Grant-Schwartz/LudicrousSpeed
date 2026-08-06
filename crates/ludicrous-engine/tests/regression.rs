use std::fs::File;
use std::io::{self, Read, Write};
use std::path::{Path, PathBuf};

use ironcalc::{
    base::{expressions::utils::number_to_column, Model},
    export::save_to_xlsx,
};
use ludicrous_engine::{
    CalcMode, ChangedCell, DataTableEvaluationStatus, DataTableOverride, FormulaValueKind,
    InlineCell, InlineDefinedName, InlineSheet, InlineWorkbook, LudicrousSpeedEngine,
    WorkbookSnapshot, WritebackMode,
};
use zip::{write::FileOptions, ZipArchive, ZipWriter};

#[test]
fn evaluates_basic_ma_workbook_and_preserves_formula_writeback_policy() {
    let fixture = create_basic_ma_fixture();
    let result = run_engine(&fixture, CalcMode::Recalculate, None);

    assert_eq!(result.analysis.coverage.formula_cells, 5);
    assert_eq!(result.analysis.coverage.supported_formula_cells, 5);
    assert_eq!(result.analysis.coverage.fallback_formula_cells, 0);
    assert!(result.analysis.ironcalc_can_evaluate);
    assert!(result.analysis.fallback_reasons.is_empty());
    assert!(result.writeback.preserve_formulas);
    assert_eq!(result.writeback.mode, WritebackMode::LiveFormulaCache);
    assert_eq!(result.writeback.value_cells_to_update, 5);
    assert_eq!(result.writeback.cells.len(), 5);
    assert!(result
        .writeback
        .cells
        .iter()
        .any(|cell| cell.sheet_name == "Sheet1"
            && cell.address == "B1"
            && cell.value_kind == FormulaValueKind::Number
            && cell.value == serde_json::json!(125.0)));
}

#[test]
fn reports_speedup_when_excel_baseline_is_supplied() {
    let fixture = create_basic_ma_fixture();
    let result = run_engine(&fixture, CalcMode::Benchmark, Some(1_000));

    assert_eq!(result.benchmark.excel_baseline_ms, Some(1_000));
    assert!(result.benchmark.speedup_vs_excel.is_some());
    assert_eq!(result.writeback.mode, WritebackMode::None);
    assert_eq!(result.writeback.value_cells_to_update, 0);
}

#[test]
fn plan_hash_is_stable_for_unchanged_workbook() {
    let fixture = create_basic_ma_fixture();
    let first = run_engine(&fixture, CalcMode::Analyze, None);
    let second = run_engine(&fixture, CalcMode::Analyze, None);

    assert_eq!(first.plan.workbook_hash, second.plan.workbook_hash);
    assert_eq!(first.plan.workbook_hash.len(), 64);
}

#[test]
fn rejects_missing_workbook_path() {
    let snapshot = WorkbookSnapshot {
        workbook_path: "/definitely/not/a/workbook.xlsx".to_string(),
        workbook_name: Some("missing.xlsx".to_string()),
        workbook_id: None,
        mode: CalcMode::Analyze,
        excel_baseline_ms: None,
        force_reload: false,
        changed_cells: Vec::new(),
        evaluate_data_tables: false,
        locale: "en".to_string(),
        timezone: "UTC".to_string(),
        language: "en".to_string(),
        inline_workbook: None,
        data_table_overrides: Vec::new(),
        include_analytics: true,
    };

    let err = LudicrousSpeedEngine::new().run(&snapshot).unwrap_err();
    assert!(err.to_string().contains("failed to load workbook"));
}

#[test]
fn handles_large_sensitivity_style_formula_block() {
    let fixture = create_sensitivity_fixture();
    let result = run_engine(&fixture, CalcMode::Recalculate, None);

    assert_eq!(result.analysis.coverage.formula_cells, 50);
    assert_eq!(result.analysis.coverage.supported_formula_cells, 50);
    assert_eq!(result.analysis.coverage.fallback_formula_cells, 0);
    assert!(result.benchmark.total_ludicrous_ms >= result.benchmark.ironcalc_ms);
}

#[test]
fn validates_openxml_data_table_in_parallel_mode() {
    let fixture = create_data_table_fixture();
    let mut snapshot = snapshot_for(&fixture, CalcMode::Benchmark, None);
    snapshot.evaluate_data_tables = true;

    let result = LudicrousSpeedEngine::new().run(&snapshot).unwrap();

    assert_eq!(result.benchmark.data_tables.data_table_count, 1);
    assert_eq!(result.benchmark.data_tables.data_table_cells, 4);
    assert_eq!(
        result.benchmark.data_tables.status,
        DataTableEvaluationStatus::Validated
    );
    assert_eq!(result.benchmark.data_tables.validated_data_table_cells, 4);
    assert_eq!(result.benchmark.data_tables.mismatched_data_table_cells, 0);
    assert!(!result
        .analysis
        .fallback_reasons
        .iter()
        .any(|reason| reason.code == "data_table_formula"));
    assert_eq!(result.analysis.coverage.fallback_formula_cells, 0);
    assert_eq!(result.writeback.mode, WritebackMode::None);
    assert_eq!(result.writeback.value_cells_to_update, 0);
}

#[test]
fn allows_formula_writeback_candidates_when_data_tables_validate() {
    let fixture = create_data_table_fixture();
    let mut snapshot = snapshot_for(&fixture, CalcMode::Recalculate, None);
    snapshot.evaluate_data_tables = true;

    let result = LudicrousSpeedEngine::new().run(&snapshot).unwrap();

    assert_eq!(
        result.benchmark.data_tables.status,
        DataTableEvaluationStatus::Validated
    );
    assert!(!result
        .analysis
        .fallback_reasons
        .iter()
        .any(|reason| reason.code == "data_table_formula"));
    assert_eq!(result.writeback.mode, WritebackMode::LiveFormulaCache);
    assert!(result.writeback.value_cells_to_update > 0);
    assert!(!result
        .writeback
        .skipped_reasons
        .iter()
        .any(|reason| reason.code == "fallback_regions_present"));
}

#[test]
fn resolves_data_table_fallback_per_table_not_workbook_wide() {
    let fixture = create_two_data_table_fixture();
    let mut snapshot = snapshot_for(&fixture, CalcMode::Recalculate, None);
    snapshot.evaluate_data_tables = true;

    let result = LudicrousSpeedEngine::new().run(&snapshot).unwrap();

    assert_eq!(result.benchmark.data_tables.data_table_count, 2);
    // Table 2 (F8:F9) was seeded with deliberately wrong cached values, so
    // the workbook overall does NOT fully validate...
    assert_eq!(
        result.benchmark.data_tables.status,
        DataTableEvaluationStatus::Mismatch
    );
    assert!(result.benchmark.data_tables.mismatched_data_table_cells > 0);
    // ...but table 1 (C3:D4) validated fine on its own, so its
    // data_table_formula fallback must be resolved even though table 2's
    // is not -- this is exactly the per-table behavior the old
    // workbook-wide gate got wrong.
    let data_table_fallback_locations = result
        .analysis
        .fallback_reasons
        .iter()
        .filter(|reason| reason.code == "data_table_formula")
        .filter_map(|reason| reason.location.as_deref())
        .collect::<Vec<_>>();
    assert!(
        !data_table_fallback_locations
            .iter()
            .any(|location| location.ends_with("!C3:D4")),
        "validated table 1 should not still be reported as a fallback: {data_table_fallback_locations:?}"
    );
    assert!(
        data_table_fallback_locations
            .iter()
            .any(|location| location.ends_with("!F8:F9")),
        "mismatched table 2 should still be reported as a fallback: {data_table_fallback_locations:?}"
    );
}

#[test]
fn writeback_includes_unaffected_cells_when_other_regions_have_fallbacks() {
    let inline = InlineWorkbook {
        sheets: vec![InlineSheet {
            name: "Sheet1".to_string(),
            cells: vec![
                InlineCell {
                    row: 1,
                    column: 1,
                    input: "10".to_string(),
                }, // A1
                InlineCell {
                    row: 2,
                    column: 1,
                    input: "=A1*2".to_string(),
                }, // A2, clean, unrelated
                InlineCell {
                    row: 1,
                    column: 2,
                    input: "=INDIRECT(\"A1\")".to_string(),
                }, // B1, fallback
                InlineCell {
                    row: 1,
                    column: 3,
                    input: "=B1+1".to_string(),
                }, // C1, downstream of B1
            ],
        }],
        defined_names: Vec::new(),
    };

    let snapshot = inline_snapshot(
        "writeback-partial",
        inline,
        CalcMode::Recalculate,
        Vec::new(),
    );
    let result = LudicrousSpeedEngine::new().run(&snapshot).unwrap();

    assert!(result
        .analysis
        .fallback_reasons
        .iter()
        .any(|reason| reason.code == "dynamic_reference"));

    // A2 doesn't depend on B1 at all, so one fallback region elsewhere in the
    // workbook must not block it from being writeback-safe.
    assert_eq!(result.writeback.mode, WritebackMode::LiveFormulaCache);
    assert!(result
        .writeback
        .cells
        .iter()
        .any(|cell| cell.address == "A2" && cell.value == serde_json::json!(20.0)));

    // C1's own formula ("=B1+1") uses only supported constructs, but its
    // value is built on B1's untrustworthy result, so it must still be
    // excluded -- and reported as excluded, not silently dropped.
    assert!(!result
        .writeback
        .cells
        .iter()
        .any(|cell| cell.address == "C1"));
    assert!(result
        .writeback
        .skipped_reasons
        .iter()
        .any(|reason| reason.code == "downstream_of_fallback" && reason.count >= 1));
}

#[test]
fn skips_evaluation_on_warm_noop_cache_hit() {
    let fixture = create_basic_ma_fixture();
    let engine = LudicrousSpeedEngine::new();
    let first = engine
        .run(&snapshot_for(&fixture, CalcMode::Benchmark, Some(1_000)))
        .unwrap();
    let second = engine
        .run(&snapshot_for(&fixture, CalcMode::Benchmark, Some(1_000)))
        .unwrap();

    assert_eq!(
        first.benchmark.strategy,
        ludicrous_engine::CalculationStrategy::ColdFull
    );
    assert_eq!(
        second.benchmark.strategy,
        ludicrous_engine::CalculationStrategy::WarmNoopCacheHit
    );
    assert!(second.benchmark.model_cache_hit);
    assert!(second.benchmark.graph_cache_hit);
    assert!(second.benchmark.result_cache_hit);
    assert_eq!(second.benchmark.ironcalc_ms, 0);
    assert_eq!(second.benchmark.cache_hit_rate, 1.0);
}

#[test]
fn reuses_model_and_graph_for_warm_value_edits() {
    let fixture = create_basic_ma_fixture();
    let engine = LudicrousSpeedEngine::new();
    let mut warm_snapshot = snapshot_for(&fixture, CalcMode::Recalculate, None);
    engine.run(&warm_snapshot).unwrap();

    warm_snapshot.workbook_path = String::new();
    warm_snapshot.changed_cells = vec![ChangedCell {
        sheet_name: "Sheet1".to_string(),
        row: 1,
        column: 1,
        address: "A1".to_string(),
        input: "150".to_string(),
        is_formula: false,
    }];

    let result = engine.run(&warm_snapshot).unwrap();
    assert_eq!(
        result.benchmark.strategy,
        ludicrous_engine::CalculationStrategy::WarmFullWithDirtyPlan
    );
    assert!(result.benchmark.model_cache_hit);
    assert!(result.benchmark.graph_cache_hit);
    assert!(!result.benchmark.result_cache_hit);
    assert!(result.benchmark.dirty_formula_cells > 0);
    assert_eq!(result.writeback.mode, WritebackMode::LiveFormulaCache);
    assert!(result
        .writeback
        .cells
        .iter()
        .any(|cell| cell.sheet_name == "Sheet1"
            && cell.address == "B1"
            && cell.value == serde_json::json!(175.0)));
}

fn run_engine(
    workbook_path: &Path,
    mode: CalcMode,
    excel_baseline_ms: Option<u128>,
) -> ludicrous_engine::CalcResult {
    let snapshot = snapshot_for(workbook_path, mode, excel_baseline_ms);

    LudicrousSpeedEngine::new().run(&snapshot).unwrap()
}

fn snapshot_for(
    workbook_path: &Path,
    mode: CalcMode,
    excel_baseline_ms: Option<u128>,
) -> WorkbookSnapshot {
    WorkbookSnapshot {
        workbook_path: workbook_path.to_string_lossy().to_string(),
        workbook_name: workbook_path
            .file_name()
            .map(|file_name| file_name.to_string_lossy().to_string()),
        workbook_id: Some(workbook_path.to_string_lossy().to_string()),
        mode,
        excel_baseline_ms,
        force_reload: false,
        changed_cells: Vec::new(),
        evaluate_data_tables: false,
        locale: "en".to_string(),
        timezone: "UTC".to_string(),
        language: "en".to_string(),
        inline_workbook: None,
        data_table_overrides: Vec::new(),
        include_analytics: true,
    }
}

fn create_basic_ma_fixture() -> PathBuf {
    let tempdir = tempfile::tempdir().unwrap().keep();
    let workbook_path = tempdir.join("basic-ma-model.xlsx");

    let mut model = Model::new_empty("basic-ma-model", "en", "UTC", "en").unwrap();
    model.set_user_input(0, 1, 1, "100".to_string()).unwrap();
    model.set_user_input(0, 2, 1, "25".to_string()).unwrap();
    model
        .set_user_input(0, 1, 2, "=SUM(A1:A2)".to_string())
        .unwrap();
    model
        .set_user_input(0, 2, 2, "=IF(B1>100,B1*2,B1)".to_string())
        .unwrap();
    model
        .set_user_input(0, 1, 3, "=INDEX(A1:A2,1)".to_string())
        .unwrap();
    model
        .set_user_input(0, 2, 3, "=B1+B2+C1".to_string())
        .unwrap();
    model.add_sheet("Debt").unwrap();
    model
        .set_user_input(1, 1, 1, "=Sheet1!C2*0.4".to_string())
        .unwrap();
    model.evaluate();

    save_to_xlsx(&model, workbook_path.to_string_lossy().as_ref()).unwrap();
    workbook_path
}

fn create_sensitivity_fixture() -> PathBuf {
    let tempdir = tempfile::tempdir().unwrap().keep();
    let workbook_path = tempdir.join("sensitivity-model.xlsx");

    let mut model = Model::new_empty("sensitivity-model", "en", "UTC", "en").unwrap();
    model
        .set_user_input(0, 1, 1, "Revenue".to_string())
        .unwrap();
    model.set_user_input(0, 1, 2, "1000".to_string()).unwrap();
    model.set_user_input(0, 2, 1, "Margin".to_string()).unwrap();
    model.set_user_input(0, 2, 2, "0.25".to_string()).unwrap();

    for row in 1..=10 {
        for column in 1..=5 {
            let excel_column = number_to_column(column + 3).unwrap();
            let formula = format!("=$B$1*$B$2*{row}*{column}");
            model
                .set_user_input(0, row + 4, column + 3, formula)
                .unwrap_or_else(|err| panic!("failed to set {excel_column}{}: {err}", row + 4));
        }
    }
    model.evaluate();

    save_to_xlsx(&model, workbook_path.to_string_lossy().as_ref()).unwrap();
    workbook_path
}

fn create_data_table_fixture() -> PathBuf {
    let tempdir = tempfile::tempdir().unwrap().keep();
    let workbook_path = tempdir.join("data-table-model.xlsx");
    let patched_path = tempdir.join("data-table-model-patched.xlsx");

    let mut model = Model::new_empty("data-table-model", "en", "UTC", "en").unwrap();
    model.set_user_input(0, 1, 1, "2".to_string()).unwrap();
    model.set_user_input(0, 2, 1, "10".to_string()).unwrap();
    model.set_user_input(0, 2, 2, "=A1*A2".to_string()).unwrap();
    model.set_user_input(0, 2, 3, "2".to_string()).unwrap();
    model.set_user_input(0, 2, 4, "3".to_string()).unwrap();
    model.set_user_input(0, 3, 2, "10".to_string()).unwrap();
    model.set_user_input(0, 4, 2, "20".to_string()).unwrap();
    model.set_user_input(0, 3, 3, "20".to_string()).unwrap();
    model.set_user_input(0, 3, 4, "30".to_string()).unwrap();
    model.set_user_input(0, 4, 3, "40".to_string()).unwrap();
    model.set_user_input(0, 4, 4, "60".to_string()).unwrap();
    model.evaluate();

    save_to_xlsx(&model, workbook_path.to_string_lossy().as_ref()).unwrap();
    inject_data_table_formula(&workbook_path, &patched_path);
    patched_path
}

fn inject_data_table_formula(source_path: &Path, output_path: &Path) {
    let source = File::open(source_path).unwrap();
    let mut archive = ZipArchive::new(source).unwrap();
    let output = File::create(output_path).unwrap();
    let mut writer = ZipWriter::new(output);

    for index in 0..archive.len() {
        let mut file = archive.by_index(index).unwrap();
        let name = file.name().to_string();
        let options = FileOptions::default()
            .compression_method(file.compression())
            .last_modified_time(file.last_modified());

        if name.ends_with('/') {
            writer.add_directory(name, options).unwrap();
            continue;
        }

        writer.start_file(name.clone(), options).unwrap();
        if name == "xl/worksheets/sheet1.xml" {
            let mut xml = String::new();
            file.read_to_string(&mut xml).unwrap();
            let patched = insert_data_table_formula(
                &xml,
                "C3",
                r#"<f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="A1" r2="A2" ca="1"/>"#,
            );
            writer.write_all(patched.as_bytes()).unwrap();
        } else {
            io::copy(&mut file, &mut writer).unwrap();
        }
    }

    writer.finish().unwrap();
}

/// A workbook with two independent native Excel data tables: one whose
/// cached body matches what the kernel computes (validates cleanly) and one
/// whose cached body is deliberately wrong (guaranteed mismatch), to test
/// that a mismatch in one table doesn't mask the other table's fallback
/// having been resolved.
fn create_two_data_table_fixture() -> PathBuf {
    let tempdir = tempfile::tempdir().unwrap().keep();
    let workbook_path = tempdir.join("two-data-table-model.xlsx");
    let patched_path = tempdir.join("two-data-table-model-patched.xlsx");

    let mut model = Model::new_empty("two-data-table-model", "en", "UTC", "en").unwrap();

    // Table 1 (A1:D4), same shape as create_data_table_fixture: validates.
    model.set_user_input(0, 1, 1, "2".to_string()).unwrap();
    model.set_user_input(0, 2, 1, "10".to_string()).unwrap();
    model.set_user_input(0, 2, 2, "=A1*A2".to_string()).unwrap();
    model.set_user_input(0, 2, 3, "2".to_string()).unwrap();
    model.set_user_input(0, 2, 4, "3".to_string()).unwrap();
    model.set_user_input(0, 3, 2, "10".to_string()).unwrap();
    model.set_user_input(0, 4, 2, "20".to_string()).unwrap();
    model.set_user_input(0, 3, 3, "20".to_string()).unwrap();
    model.set_user_input(0, 3, 4, "30".to_string()).unwrap();
    model.set_user_input(0, 4, 3, "40".to_string()).unwrap();
    model.set_user_input(0, 4, 4, "60".to_string()).unwrap();

    // Table 2: one-variable, column-oriented (F1 varied, F8:F9 body),
    // deliberately given wrong cached body values so it always mismatches.
    model.set_user_input(0, 1, 6, "5".to_string()).unwrap(); // F1 input
    model.set_user_input(0, 7, 6, "=F1*3".to_string()).unwrap(); // F7 source
    model.set_user_input(0, 8, 5, "1".to_string()).unwrap(); // E8 axis
    model.set_user_input(0, 9, 5, "2".to_string()).unwrap(); // E9 axis
    model.set_user_input(0, 8, 6, "999".to_string()).unwrap(); // F8 body (wrong)
    model.set_user_input(0, 9, 6, "999".to_string()).unwrap(); // F9 body (wrong)
    model.evaluate();

    save_to_xlsx(&model, workbook_path.to_string_lossy().as_ref()).unwrap();

    let source = File::open(&workbook_path).unwrap();
    let mut archive = ZipArchive::new(source).unwrap();
    let output = File::create(&patched_path).unwrap();
    let mut writer = ZipWriter::new(output);

    for index in 0..archive.len() {
        let mut file = archive.by_index(index).unwrap();
        let name = file.name().to_string();
        let options = FileOptions::default()
            .compression_method(file.compression())
            .last_modified_time(file.last_modified());

        if name.ends_with('/') {
            writer.add_directory(name, options).unwrap();
            continue;
        }

        writer.start_file(name.clone(), options).unwrap();
        if name == "xl/worksheets/sheet1.xml" {
            let mut xml = String::new();
            file.read_to_string(&mut xml).unwrap();
            let patched = insert_data_table_formula(
                &xml,
                "C3",
                r#"<f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="A1" r2="A2" ca="1"/>"#,
            );
            let patched = insert_data_table_formula(
                &patched,
                "F8",
                r#"<f t="dataTable" ref="F8:F9" dtr="0" r1="F1" ca="1"/>"#,
            );
            writer.write_all(patched.as_bytes()).unwrap();
        } else {
            io::copy(&mut file, &mut writer).unwrap();
        }
    }

    writer.finish().unwrap();
    patched_path
}

fn insert_data_table_formula(xml: &str, cell_address: &str, formula_tag: &str) -> String {
    let cell_start = xml
        .find(&format!(r#"<c r="{cell_address}""#))
        .expect("cell exists");
    let open_end = cell_start + xml[cell_start..].find('>').expect("cell open tag ends") + 1;
    let close_start = open_end + xml[open_end..].find("</c>").expect("cell close tag exists");
    format!(
        "{}{}{}",
        &xml[..open_end],
        formula_tag,
        &xml[open_end..close_start]
    ) + &xml[close_start..]
}

fn inline_snapshot(
    workbook_id: &str,
    inline: InlineWorkbook,
    mode: CalcMode,
    changed_cells: Vec<ChangedCell>,
) -> WorkbookSnapshot {
    WorkbookSnapshot {
        workbook_path: String::new(),
        workbook_name: Some(workbook_id.to_string()),
        workbook_id: Some(workbook_id.to_string()),
        mode,
        excel_baseline_ms: None,
        force_reload: false,
        changed_cells,
        evaluate_data_tables: false,
        locale: "en".to_string(),
        timezone: "UTC".to_string(),
        language: "en".to_string(),
        inline_workbook: Some(inline),
        data_table_overrides: Vec::new(),
        include_analytics: true,
    }
}

#[test]
fn builds_and_evaluates_a_workbook_from_inline_cells_without_any_file() {
    let inline = InlineWorkbook {
        sheets: vec![
            InlineSheet {
                name: "Assumptions".to_string(),
                cells: vec![
                    InlineCell {
                        row: 1,
                        column: 1,
                        input: "10".to_string(),
                    },
                    InlineCell {
                        row: 2,
                        column: 1,
                        input: "5".to_string(),
                    },
                ],
            },
            InlineSheet {
                name: "Model".to_string(),
                cells: vec![InlineCell {
                    row: 1,
                    column: 1,
                    input: "=Assumptions!A1*Assumptions!A2".to_string(),
                }],
            },
        ],
        defined_names: vec![InlineDefinedName {
            name: "Total".to_string(),
            scope_sheet_name: None,
            formula: "Model!$A$1".to_string(),
        }],
    };

    let snapshot = inline_snapshot("inline-basic", inline, CalcMode::Recalculate, Vec::new());
    let result = LudicrousSpeedEngine::new().run(&snapshot).unwrap();

    assert_eq!(result.analysis.coverage.formula_cells, 1);
    assert_eq!(result.analysis.coverage.fallback_formula_cells, 0);
    assert!(result.analysis.ironcalc_can_evaluate);
    assert_eq!(result.writeback.cells.len(), 1);
    assert_eq!(result.writeback.cells[0].value, serde_json::json!(50.0));
}

#[test]
fn inline_build_skips_bad_cells_as_fallbacks_instead_of_failing_entirely() {
    let inline = InlineWorkbook {
        sheets: vec![InlineSheet {
            name: "Sheet1".to_string(),
            cells: vec![
                InlineCell {
                    row: 1,
                    column: 1,
                    input: "10".to_string(),
                },
                // Row 0 is invalid (Excel/IronCalc are 1-indexed) -- exactly
                // the shape of bug a COM bulk-read on the host side could
                // introduce (an off-by-one row/column index). This must be
                // recorded as a fallback and skipped, not abort the whole
                // workbook build.
                InlineCell {
                    row: 0,
                    column: 1,
                    input: "5".to_string(),
                },
                InlineCell {
                    row: 3,
                    column: 1,
                    input: "=A1+1".to_string(),
                },
            ],
        }],
        defined_names: Vec::new(),
    };

    let snapshot = inline_snapshot("inline-bad-cell", inline, CalcMode::Analyze, Vec::new());
    let result = LudicrousSpeedEngine::new().run(&snapshot).unwrap();

    // The bad cell is recorded as a fallback rather than aborting the build...
    assert_eq!(result.analysis.coverage.fallback_formula_cells, 1);
    assert!(result
        .analysis
        .fallback_reasons
        .iter()
        .any(|reason| reason.code == "unsupported_formula"));
    // ...and the good formula cell (A3) still evaluates normally alongside it.
    assert_eq!(result.analysis.coverage.formula_cells, 2);
    assert_eq!(result.analysis.coverage.supported_formula_cells, 1);
}

#[test]
fn inline_workbook_id_can_be_warmed_with_changed_cells_on_a_later_run() {
    let inline = InlineWorkbook {
        sheets: vec![InlineSheet {
            name: "Sheet1".to_string(),
            cells: vec![
                InlineCell {
                    row: 1,
                    column: 1,
                    input: "10".to_string(),
                },
                InlineCell {
                    row: 1,
                    column: 2,
                    input: "=A1*2".to_string(),
                },
            ],
        }],
        defined_names: Vec::new(),
    };

    let cold_snapshot = inline_snapshot("inline-warm", inline, CalcMode::Recalculate, Vec::new());
    let engine = LudicrousSpeedEngine::new();
    let cold_result = engine.run(&cold_snapshot).unwrap();
    assert_eq!(
        cold_result.writeback.cells[0].value,
        serde_json::json!(20.0)
    );

    // Warm run: no workbook_path and no inline_workbook, just the changed
    // cell and the same workbook_id, exactly like a live-edit follow-up call.
    let warm_snapshot = WorkbookSnapshot {
        workbook_path: String::new(),
        workbook_name: None,
        workbook_id: Some("inline-warm".to_string()),
        mode: CalcMode::Recalculate,
        excel_baseline_ms: None,
        force_reload: false,
        changed_cells: vec![ChangedCell {
            sheet_name: "Sheet1".to_string(),
            row: 1,
            column: 1,
            address: "A1".to_string(),
            input: "7".to_string(),
            is_formula: false,
        }],
        evaluate_data_tables: false,
        locale: "en".to_string(),
        timezone: "UTC".to_string(),
        language: "en".to_string(),
        inline_workbook: None,
        data_table_overrides: Vec::new(),
        include_analytics: true,
    };
    let warm_result = engine.run(&warm_snapshot).unwrap();
    assert_eq!(
        warm_result.writeback.cells[0].value,
        serde_json::json!(14.0)
    );
}

#[test]
fn data_table_values_persist_across_warm_runs() {
    let fixture = create_data_table_fixture();
    let engine = LudicrousSpeedEngine::new();

    let mut cold = snapshot_for(&fixture, CalcMode::Recalculate, None);
    cold.evaluate_data_tables = true;
    let cold_result = engine.run(&cold).unwrap();
    let cold_values = cold_result.writeback.data_table_cells.len();
    assert!(
        cold_values > 0,
        "cold run should publish data table cell values"
    );

    // Warm no-change run: the engine re-evaluates no tables at all. It must
    // still report current values for every data table cell, or a host
    // driving those cells live would watch them go stale on the second
    // recalc -- which is exactly the bug this guards.
    let mut warm = snapshot_for(&fixture, CalcMode::Recalculate, None);
    warm.evaluate_data_tables = true;
    warm.workbook_path = String::new();
    let warm_result = engine.run(&warm).unwrap();
    assert_eq!(
        warm_result.writeback.data_table_cells.len(),
        cold_values,
        "warm no-change run must still report every data table cell value"
    );

    // And a warm run with an edit, where only dirty tables are recomputed,
    // must still report the full set rather than only the recomputed ones.
    let mut edited = snapshot_for(&fixture, CalcMode::Recalculate, None);
    edited.evaluate_data_tables = true;
    edited.workbook_path = String::new();
    edited.changed_cells = vec![ChangedCell {
        sheet_name: "Sheet1".to_string(),
        row: 1,
        column: 1,
        address: "A1".to_string(),
        input: "7".to_string(),
        is_formula: false,
    }];
    let edited_result = engine.run(&edited).unwrap();
    assert_eq!(
        edited_result.writeback.data_table_cells.len(),
        cold_values,
        "warm edited run must report all data table cells, not only recomputed tables"
    );
}

#[test]
fn changed_cell_on_unknown_sheet_does_not_fail_the_run() {
    let fixture = create_basic_ma_fixture();
    let engine = LudicrousSpeedEngine::new();

    let mut cold = snapshot_for(&fixture, CalcMode::Recalculate, None);
    let cold_result = engine.run(&cold).unwrap();
    let expected_cells = cold_result.writeback.cells.len();

    // A warm run (no workbook path, so no reload is possible) carrying an
    // edit on a sheet the cached model has never seen. This happens whenever
    // a sheet is created after the last snapshot -- LudicrousSpeed's own
    // bookkeeping sheets being the obvious case. Such a cell cannot affect
    // anything the model computes, so it must be skipped rather than taking
    // the entire run down with "sheet not found for changed cell".
    cold.workbook_path = String::new();
    cold.changed_cells = vec![ChangedCell {
        sheet_name: "_LudicrousSpeed_DataTables".to_string(),
        row: 2,
        column: 1,
        address: "A2".to_string(),
        input: "Sheet1!C3:D4".to_string(),
        is_formula: false,
    }];

    let warm_result = engine
        .run(&cold)
        .expect("an edit on an unknown sheet must not fail the run");
    assert_eq!(
        warm_result.writeback.cells.len(),
        expected_cells,
        "the rest of the workbook should still be evaluated normally"
    );
}

/// Rewrites a data table's body the way the Excel host's "Convert to Live"
/// does: the {=TABLE()} array marker is gone and every body cell holds a
/// WS.LIVE formula pointing at itself.
fn convert_data_table_to_live_cells(source_path: &Path, output_path: &Path) {
    let source = File::open(source_path).unwrap();
    let mut archive = ZipArchive::new(source).unwrap();
    let output = File::create(output_path).unwrap();
    let mut writer = ZipWriter::new(output);

    for index in 0..archive.len() {
        let mut file = archive.by_index(index).unwrap();
        let name = file.name().to_string();
        let options = FileOptions::default()
            .compression_method(file.compression())
            .last_modified_time(file.last_modified());

        if name.ends_with('/') {
            writer.add_directory(name, options).unwrap();
            continue;
        }

        writer.start_file(name.clone(), options).unwrap();
        if name == "xl/worksheets/sheet1.xml" {
            let mut xml = String::new();
            file.read_to_string(&mut xml).unwrap();
            let mut xml = xml.replace(
                r#"<f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="A1" r2="A2" ca="1"/>"#,
                r#"<f>WS.LIVE("Sheet1!C3")</f>"#,
            );
            for address in ["D3", "C4", "D4"] {
                xml = xml.replace(
                    &format!(r#"<c r="{address}">"#),
                    &format!(r#"<c r="{address}"><f>WS.LIVE("Sheet1!{address}")</f>"#),
                );
            }
            writer.write_all(xml.as_bytes()).unwrap();
        } else {
            io::copy(&mut file, &mut writer).unwrap();
        }
    }

    writer.finish().unwrap();
}

#[test]
fn converted_data_table_is_still_computed_from_a_host_override() {
    let native = create_data_table_fixture();
    let converted = native.with_file_name("converted-data-table-model.xlsx");
    convert_data_table_to_live_cells(&native, &converted);

    // Without an override the engine has nothing to go on: the {=TABLE()}
    // marker it discovers tables from is gone.
    let mut blind = snapshot_for(&converted, CalcMode::Recalculate, None);
    blind.evaluate_data_tables = true;
    blind.workbook_id = Some("converted-blind".to_string());
    let blind_result = LudicrousSpeedEngine::new().run(&blind).unwrap();
    assert_eq!(
        blind_result.benchmark.data_tables.data_table_count, 0,
        "a converted table leaves no marker to discover"
    );

    // With the override the host persisted at conversion time, the engine
    // rebuilds the region and computes the table again -- this is the loop
    // that has to close for live data tables to keep updating.
    let mut declared = snapshot_for(&converted, CalcMode::Recalculate, None);
    declared.evaluate_data_tables = true;
    declared.workbook_id = Some("converted-override".to_string());
    declared.data_table_overrides = vec![DataTableOverride {
        sheet_name: "Sheet1".to_string(),
        range_address: "C3:D4".to_string(),
        anchor_address: "C3".to_string(),
        column_input_cell: Some("A1".to_string()),
        row_input_cell: Some("A2".to_string()),
        is_two_dimensional: true,
        dtr: true,
    }];

    let result = LudicrousSpeedEngine::new().run(&declared).unwrap();
    assert_eq!(
        result.benchmark.data_tables.data_table_count, 1,
        "the host-declared table must be picked up"
    );
    assert_eq!(
        result.writeback.data_table_cells.len(),
        4,
        "all four body cells must get computed values"
    );

    // The body now holds LudicrousSpeed's own previous output, so there is nothing
    // meaningful to validate against and the comparison must be reported as
    // absent rather than as a pass.
    assert!(
        result
            .writeback
            .data_table_cells
            .iter()
            .all(|cell| cell.matched_excel_cache.is_none()),
        "an override table has no Excel value to compare against"
    );

    // The WS.LIVE formulas must not become unsupported_formula fallbacks --
    // that would taint everything downstream and make coverage worse after
    // converting than before.
    assert!(
        !result
            .analysis
            .fallback_reasons
            .iter()
            .any(|reason| reason.message.to_ascii_uppercase().contains("WS.LIVE")),
        "WS.LIVE cells must not be reported as unsupported formulas: {:?}",
        result.analysis.fallback_reasons
    );
}
