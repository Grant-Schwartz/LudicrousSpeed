use std::fs::File;
use std::io::{self, Read, Write};
use std::path::{Path, PathBuf};

use ironcalc::{
    base::{expressions::utils::number_to_column, Model},
    export::save_to_xlsx,
};
use warpspeed_engine::{
    CalcMode, ChangedCell, DataTableEvaluationStatus, FormulaValueKind, WarpSpeedEngine,
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
    assert_eq!(result.writeback.attempted, 0);
    assert_eq!(result.writeback.written, 0);
    assert_eq!(result.writeback.failed, 0);
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
    };

    let err = WarpSpeedEngine::new().run(&snapshot).unwrap_err();
    assert!(err.to_string().contains("failed to load workbook"));
}

#[test]
fn handles_large_sensitivity_style_formula_block() {
    let fixture = create_sensitivity_fixture();
    let result = run_engine(&fixture, CalcMode::Recalculate, None);

    assert_eq!(result.analysis.coverage.formula_cells, 50);
    assert_eq!(result.analysis.coverage.supported_formula_cells, 50);
    assert_eq!(result.analysis.coverage.fallback_formula_cells, 0);
    assert!(result.benchmark.total_warpspeed_ms >= result.benchmark.ironcalc_ms);
}

#[test]
fn validates_openxml_data_table_in_parallel_mode() {
    let fixture = create_data_table_fixture();
    let mut snapshot = snapshot_for(&fixture, CalcMode::Benchmark, None);
    snapshot.evaluate_data_tables = true;

    let result = WarpSpeedEngine::new().run(&snapshot).unwrap();

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

    let result = WarpSpeedEngine::new().run(&snapshot).unwrap();

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
fn skips_evaluation_on_warm_noop_cache_hit() {
    let fixture = create_basic_ma_fixture();
    let engine = WarpSpeedEngine::new();
    let first = engine
        .run(&snapshot_for(&fixture, CalcMode::Benchmark, Some(1_000)))
        .unwrap();
    let second = engine
        .run(&snapshot_for(&fixture, CalcMode::Benchmark, Some(1_000)))
        .unwrap();

    assert_eq!(
        first.benchmark.strategy,
        warpspeed_engine::CalculationStrategy::ColdFull
    );
    assert_eq!(
        second.benchmark.strategy,
        warpspeed_engine::CalculationStrategy::WarmNoopCacheHit
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
    let engine = WarpSpeedEngine::new();
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
        warpspeed_engine::CalculationStrategy::WarmFullWithDirtyPlan
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
) -> warpspeed_engine::CalcResult {
    let snapshot = snapshot_for(workbook_path, mode, excel_baseline_ms);

    WarpSpeedEngine::new().run(&snapshot).unwrap()
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
