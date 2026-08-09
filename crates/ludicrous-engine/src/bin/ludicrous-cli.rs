use std::env;
use std::path::Path;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::thread;
use std::time::Duration;

use ludicrous_engine::{
    progress, CalcMode, CalcResult, CalculationStrategy, ChangedCell, FallbackReason,
    LudicrousSpeedEngine, WorkbookSnapshot,
};

fn main() {
    let args = env::args().skip(1).collect::<Vec<_>>();
    if args.is_empty() || args.iter().any(|arg| arg == "--help" || arg == "-h") {
        eprintln!(
            "usage: ludicrous-cli <workbook.xlsx> [analyze|recalculate|benchmark] [--json] [--no-analytics] [--eval-data-tables] [--progress] [--runs N] [--edit Sheet!A1=value]"
        );
        std::process::exit(2);
    };

    let workbook_path = args[0].clone();
    let mut mode = CalcMode::Analyze;
    let mut output_json = false;
    let mut evaluate_data_tables = false;
    let mut runs = 1_usize;
    let mut runs_was_set = false;
    let mut edit_cells = Vec::new();
    let mut include_analytics = true;
    let mut mode_was_set = false;
    let mut show_progress = false;

    let mut index = 1;
    while index < args.len() {
        match args[index].as_str() {
            "--json" => output_json = true,
            "--no-analytics" => include_analytics = false,
            "--eval-data-tables" => evaluate_data_tables = true,
            "--progress" => show_progress = true,
            "--runs" => {
                runs_was_set = true;
                index += 1;
                let Some(value) = args.get(index) else {
                    eprintln!("--runs requires a positive integer");
                    std::process::exit(2);
                };
                runs = match value.parse::<usize>() {
                    Ok(value) if value > 0 => value,
                    _ => {
                        eprintln!("--runs requires a positive integer");
                        std::process::exit(2);
                    }
                };
            }
            "--edit" | "--set" => {
                index += 1;
                let Some(value) = args.get(index) else {
                    eprintln!("--edit requires a value like Sheet!A1=125.0");
                    std::process::exit(2);
                };
                match parse_edit_arg(value) {
                    Ok(change) => edit_cells.push(change),
                    Err(err) => {
                        eprintln!("{err}");
                        std::process::exit(2);
                    }
                }
            }
            "benchmark" | "recalculate" | "analyze" if mode_was_set => {
                eprintln!("mode specified more than once");
                std::process::exit(2);
            }
            "benchmark" => {
                mode = CalcMode::Benchmark;
                mode_was_set = true;
            }
            "recalculate" => {
                mode = CalcMode::Recalculate;
                mode_was_set = true;
            }
            "analyze" => {
                mode = CalcMode::Analyze;
                mode_was_set = true;
            }
            other => {
                eprintln!("unknown argument: {other}");
                std::process::exit(2);
            }
        }
        index += 1;
    }

    if !edit_cells.is_empty() {
        if runs_was_set && runs < 2 {
            eprintln!("--edit requires at least two runs so the cache can be warmed first");
            std::process::exit(2);
        }
        if !runs_was_set {
            runs = 2;
        }
    }

    let workbook_name = Path::new(&workbook_path)
        .file_name()
        .map(|name| name.to_string_lossy().to_string());

    let snapshot = WorkbookSnapshot {
        workbook_path: workbook_path.clone(),
        workbook_name,
        workbook_id: Some(workbook_path.clone()),
        mode,
        excel_baseline_ms: None,
        force_reload: false,
        changed_cells: Vec::new(),
        evaluate_data_tables,
        locale: "en".to_string(),
        timezone: "UTC".to_string(),
        language: "en".to_string(),
        inline_workbook: None,
        data_table_overrides: Vec::new(),
        include_analytics,
    };

    let engine = LudicrousSpeedEngine::new();

    // Polls the same counters the Excel host polls, so this exercises the real
    // mechanism rather than a CLI-only imitation of it.
    let progress_stop = Arc::new(AtomicBool::new(false));
    let progress_thread = if show_progress {
        let stop = Arc::clone(&progress_stop);
        Some(thread::spawn(move || {
            while !stop.load(Ordering::Relaxed) {
                let (phase, done, total) = progress::snapshot();
                if phase != progress::PHASE_IDLE {
                    eprint!("\r\x1b[2K{}", format_progress(phase, done, total));
                }
                thread::sleep(Duration::from_millis(100));
            }
            eprint!("\r\x1b[2K");
        }))
    } else {
        None
    };

    let mut results = Vec::new();
    for run_index in 0..runs {
        let mut run_snapshot = snapshot.clone();
        if !edit_cells.is_empty() && run_index == 1 {
            run_snapshot.workbook_path = String::new();
            run_snapshot.changed_cells = edit_cells.clone();
        } else if !edit_cells.is_empty() && run_index > 1 {
            run_snapshot.workbook_path = String::new();
        }

        match engine.run(&run_snapshot) {
            Ok(result) => {
                if output_json {
                    results.push(result);
                } else {
                    if runs > 1 {
                        println!(
                            "Run {}/{} ({})",
                            run_index + 1,
                            runs,
                            run_label(run_index, !edit_cells.is_empty())
                        );
                        if !edit_cells.is_empty() && run_index == 1 {
                            print_edits(&edit_cells);
                        }
                    }
                    print_summary(&snapshot.workbook_path, &result);
                    if run_index + 1 < runs {
                        println!();
                    }
                }
            }
            Err(err) => {
                eprintln!("{err}");
                std::process::exit(1);
            }
        }
    }

    progress_stop.store(true, Ordering::Relaxed);
    if let Some(handle) = progress_thread {
        let _ = handle.join();
    }

    if output_json {
        if runs == 1 {
            println!("{}", serde_json::to_string_pretty(&results[0]).unwrap());
        } else {
            println!("{}", serde_json::to_string_pretty(&results).unwrap());
        }
    }
}

fn format_progress(phase: u32, done: u64, total: u64) -> String {
    let name = match phase {
        progress::PHASE_LOADING => "Loading workbook",
        progress::PHASE_ANALYZING => "Building dependency graph",
        progress::PHASE_EVALUATING => "Evaluating",
        progress::PHASE_DATA_TABLES => "Data tables",
        _ => "Working",
    };
    if total > 0 {
        format!("{name}: {done}/{total} ({}%)", done * 100 / total)
    } else {
        format!("{name}...")
    }
}

fn run_label(run_index: usize, has_edits: bool) -> &'static str {
    if !has_edits {
        return if run_index == 0 {
            "cold"
        } else {
            "warm no-change"
        };
    }

    match run_index {
        0 => "cold baseline",
        1 => "warm edit",
        _ => "warm no-change after edit",
    }
}

fn print_edits(edit_cells: &[ChangedCell]) {
    println!("Applied edits:");
    for edit in edit_cells {
        println!("  {}!{} = {}", edit.sheet_name, edit.address, edit.input);
    }
}

fn print_summary(workbook_path: &str, result: &CalcResult) {
    let workbook_name = result
        .analysis
        .workbook_name
        .as_deref()
        .or_else(|| {
            Path::new(workbook_path)
                .file_name()
                .and_then(|name| name.to_str())
        })
        .unwrap_or(workbook_path);
    let coverage = &result.analysis.coverage;
    let benchmark = &result.benchmark;

    println!("Workbook: {workbook_name}");
    println!("Strategy: {}", strategy_label(benchmark.strategy));
    println!(
        "Coverage: {}/{} supported ({:.1}%), {} fallback formulas",
        coverage.supported_formula_cells,
        coverage.formula_cells,
        coverage.coverage_ratio() * 100.0,
        coverage.fallback_formula_cells
    );
    println!(
        "Timings: total={} ms, load={} ms, graph={} ms, ironcalc={} ms, cache_lookup={} ms",
        benchmark.total_ludicrous_ms,
        benchmark.load_ms,
        benchmark.graph_build_ms,
        benchmark.ironcalc_ms,
        benchmark.cache_lookup_ms
    );
    println!(
        "Cache: model={}, graph={}, result={}, planned_reuse={:.1}%",
        benchmark.model_cache_hit,
        benchmark.graph_cache_hit,
        benchmark.result_cache_hit,
        benchmark.cache_hit_rate * 100.0
    );
    println!(
        "Dirty plan: dirty_formulas={}, reusable_formulas={}",
        benchmark.dirty_formula_cells, benchmark.planned_reusable_formula_cells
    );
    println!(
        "Data tables: status={}, tables={}, cells={}, dirty={}, reused={}, evaluated={}, validated={}, mismatched={} (stale_cache={}), unsupported={}, eval={} ms, workers={}",
        data_table_status_label(benchmark.data_tables.status),
        benchmark.data_tables.data_table_count,
        benchmark.data_tables.data_table_cells,
        benchmark.data_tables.dirty_data_tables,
        benchmark.data_tables.reused_data_table_cells,
        benchmark.data_tables.evaluated_data_table_cells,
        benchmark.data_tables.validated_data_table_cells,
        benchmark.data_tables.mismatched_data_table_cells,
        benchmark.data_tables.stale_cache_data_table_cells,
        benchmark.data_tables.unsupported_data_table_cells,
        benchmark.data_tables.data_table_eval_ms,
        benchmark.data_tables.data_table_parallelism
    );
    print_data_table_diagnostics(&benchmark.data_tables.diagnostics);

    if !result.analysis.fallback_reasons.is_empty() {
        print_fallback_summary(&result.analysis.fallback_reasons);
    }

    println!("Use --json to print the full engine response.");
}

fn parse_edit_arg(argument: &str) -> Result<ChangedCell, String> {
    let Some((target, input)) = argument.split_once('=') else {
        return Err("--edit requires a value like Sheet!A1=125.0".to_string());
    };
    let Some((sheet_ref, address_ref)) = target.rsplit_once('!') else {
        return Err(
            "--edit requires an explicit sheet, for example 'Assumptions!C12=125.0'".to_string(),
        );
    };

    let sheet_name = parse_sheet_name(sheet_ref)?;
    let (address, row, column) = parse_a1_address(address_ref)?;

    Ok(ChangedCell {
        sheet_name,
        row,
        column,
        address,
        input: input.to_string(),
        is_formula: input.trim_start().starts_with('='),
    })
}

fn parse_sheet_name(sheet_ref: &str) -> Result<String, String> {
    let trimmed = sheet_ref.trim();
    if trimmed.is_empty() {
        return Err("--edit sheet name cannot be empty".to_string());
    }

    if trimmed.starts_with('\'') || trimmed.ends_with('\'') {
        if !(trimmed.starts_with('\'') && trimmed.ends_with('\'')) || trimmed.len() < 2 {
            return Err(
                "--edit quoted sheet names must look like 'Sheet Name'!A1=125.0".to_string(),
            );
        }
        let unquoted = &trimmed[1..trimmed.len() - 1];
        let sheet_name = unquoted.replace("''", "'");
        if sheet_name.is_empty() {
            return Err("--edit sheet name cannot be empty".to_string());
        }
        return Ok(sheet_name);
    }

    Ok(trimmed.to_string())
}

fn parse_a1_address(address_ref: &str) -> Result<(String, i32, i32), String> {
    let normalized = address_ref.trim().replace('$', "").to_ascii_uppercase();
    if normalized.is_empty() {
        return Err("--edit cell address cannot be empty".to_string());
    }

    let mut split_at = 0;
    for ch in normalized.chars() {
        if ch.is_ascii_alphabetic() {
            split_at += 1;
        } else {
            break;
        }
    }

    let (column_letters, row_digits) = normalized.split_at(split_at);
    if column_letters.is_empty()
        || row_digits.is_empty()
        || !row_digits.chars().all(|ch| ch.is_ascii_digit())
    {
        return Err(format!(
            "--edit cell address must be A1-style, got {address_ref}"
        ));
    }

    let mut column = 0_i32;
    for ch in column_letters.bytes() {
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add((ch - b'A' + 1) as i32))
            .ok_or_else(|| format!("--edit column is too large: {address_ref}"))?;
    }

    let row = row_digits
        .parse::<i32>()
        .map_err(|_| format!("--edit row is too large: {address_ref}"))?;
    if row <= 0 || column <= 0 {
        return Err(format!(
            "--edit cell address must be positive, got {address_ref}"
        ));
    }

    Ok((format!("{column_letters}{row}"), row, column))
}

fn print_fallback_summary(fallback_reasons: &[FallbackReason]) {
    let mut counts = std::collections::BTreeMap::<&str, usize>::new();
    for reason in fallback_reasons {
        *counts.entry(reason.code.as_str()).or_default() += 1;
    }

    println!("Fallback reasons: {}", fallback_reasons.len());
    for (code, count) in counts {
        println!("  {code}: {count}");
    }

    println!("Sample fallbacks:");
    for reason in fallback_reasons.iter().take(10) {
        let location = reason.location.as_deref().unwrap_or("<unknown>");
        println!("  - {} at {}: {}", reason.code, location, reason.message);
    }
}

fn print_data_table_diagnostics(diagnostics: &[ludicrous_engine::DataTableDiagnostic]) {
    if diagnostics.is_empty() {
        return;
    }

    let mut counts = std::collections::BTreeMap::<&str, usize>::new();
    for diagnostic in diagnostics {
        *counts.entry(diagnostic.code.as_str()).or_default() += 1;
    }

    println!("Data table diagnostics: {}", diagnostics.len());
    for (code, count) in counts {
        println!("  {code}: {count}");
    }

    println!("Sample data table diagnostics:");
    for diagnostic in diagnostics.iter().take(10) {
        let formula_cell = diagnostic.formula_cell.as_deref().unwrap_or("<unknown>");
        println!(
            "  - {} at {} {} ({} cells): {}",
            diagnostic.code,
            diagnostic.table_id,
            formula_cell,
            diagnostic.affected_cells,
            diagnostic.message
        );
        if let Some(formula) = diagnostic.formula.as_deref() {
            println!("    formula: {formula}");
        }
    }
}

fn strategy_label(strategy: CalculationStrategy) -> &'static str {
    match strategy {
        CalculationStrategy::ColdFull => "cold_full",
        CalculationStrategy::WarmFullWithDirtyPlan => "warm_full_with_dirty_plan",
        CalculationStrategy::WarmNoopCacheHit => "warm_noop_cache_hit",
        CalculationStrategy::ForcedReload => "forced_reload",
    }
}

fn data_table_status_label(status: ludicrous_engine::DataTableEvaluationStatus) -> &'static str {
    match status {
        ludicrous_engine::DataTableEvaluationStatus::None => "none",
        ludicrous_engine::DataTableEvaluationStatus::MetadataOnly => "metadata_only",
        ludicrous_engine::DataTableEvaluationStatus::Reused => "reused",
        ludicrous_engine::DataTableEvaluationStatus::Validated => "validated",
        ludicrous_engine::DataTableEvaluationStatus::Mismatch => "mismatch",
        ludicrous_engine::DataTableEvaluationStatus::Unsupported => "unsupported",
        ludicrous_engine::DataTableEvaluationStatus::Partial => "partial",
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_value_edit() {
        let edit = parse_edit_arg("Assumptions!C12=125.0").unwrap();

        assert_eq!(edit.sheet_name, "Assumptions");
        assert_eq!(edit.address, "C12");
        assert_eq!(edit.row, 12);
        assert_eq!(edit.column, 3);
        assert_eq!(edit.input, "125.0");
        assert!(!edit.is_formula);
    }

    #[test]
    fn parses_formula_edit_with_quoted_sheet_name() {
        let edit = parse_edit_arg("'Case Inputs'!$AA$7==SUM(A1:A3)").unwrap();

        assert_eq!(edit.sheet_name, "Case Inputs");
        assert_eq!(edit.address, "AA7");
        assert_eq!(edit.row, 7);
        assert_eq!(edit.column, 27);
        assert_eq!(edit.input, "=SUM(A1:A3)");
        assert!(edit.is_formula);
    }

    #[test]
    fn rejects_edit_without_sheet_name() {
        assert!(parse_edit_arg("A1=10").is_err());
    }
}
