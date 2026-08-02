use std::collections::{HashMap, HashSet};
use std::thread;
use std::time::Instant;

use ironcalc::base::{
    cell::CellValue,
    expressions::{
        parser::Node,
        token::{OpCompare, OpProduct, OpSum, OpUnary},
        utils::number_to_column,
    },
    Model,
};

use crate::graph::CellId;
use crate::model::{DataTableBenchmarkSummary, DataTableDiagnostic, DataTableEvaluationStatus};

const NUMERIC_TOLERANCE: f64 = 1e-7;
const ITERATIVE_MAX_ITERATIONS: usize = 100;
const ITERATIVE_MAX_CHANGE: f64 = 1e-7;

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SheetCellRef {
    pub sheet_name: String,
    pub row: i32,
    pub column: i32,
    pub address: String,
}

impl SheetCellRef {
    pub(crate) fn new(sheet_name: &str, row: i32, column: i32) -> Option<Self> {
        Some(Self {
            sheet_name: sheet_name.to_string(),
            row,
            column,
            address: format!("{}{}", number_to_column(column)?, row),
        })
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SheetRange {
    pub start_row: i32,
    pub start_column: i32,
    pub end_row: i32,
    pub end_column: i32,
}

impl SheetRange {
    pub(crate) fn contains(&self, row: i32, column: i32) -> bool {
        self.start_row <= row
            && row <= self.end_row
            && self.start_column <= column
            && column <= self.end_column
    }

    pub(crate) fn cell_count(&self) -> usize {
        if self.end_row < self.start_row || self.end_column < self.start_column {
            return 0;
        }

        ((self.end_row - self.start_row + 1) as usize)
            * ((self.end_column - self.start_column + 1) as usize)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct DataTableRegion {
    pub id: String,
    pub sheet_name: String,
    pub sheet_part: String,
    pub anchor_address: String,
    pub range_address: String,
    pub range: SheetRange,
    pub formula_cell: Option<SheetCellRef>,
    pub column_input_cell: Option<SheetCellRef>,
    pub row_input_cell: Option<SheetCellRef>,
    pub is_two_dimensional: bool,
    pub dtr: bool,
    pub unsupported_reason: Option<String>,
}

impl DataTableRegion {
    pub(crate) fn cell_count(&self) -> usize {
        self.range.cell_count()
    }

    fn is_parallel_validation_eligible(&self) -> bool {
        if self.unsupported_reason.is_some() || self.formula_cell.is_none() {
            return false;
        }

        if self.is_two_dimensional {
            self.column_input_cell.is_some()
                && self.row_input_cell.is_some()
                && self.range.start_row > 1
                && self.range.start_column > 1
        } else if self.dtr {
            self.column_input_cell.is_some() && self.range.start_row > 1
        } else {
            self.column_input_cell.is_some() && self.range.start_column > 1
        }
    }
}

#[derive(Clone, Debug)]
struct DataTableScenario {
    column_axis: Option<SheetCellRef>,
    row_axis: Option<SheetCellRef>,
    output_cell: SheetCellRef,
}

#[derive(Clone, Debug, Default)]
struct ScenarioCounts {
    evaluated: usize,
    validated: usize,
    mismatched: usize,
    unsupported: usize,
    diagnostics: Vec<DataTableDiagnostic>,
}

#[derive(Clone, Debug, PartialEq)]
enum ComparableValue {
    None,
    Number(f64),
    String(String),
    Boolean(bool),
}

pub(crate) fn parse_sheet_range(reference: &str) -> Option<SheetRange> {
    let mut parts = reference.split(':');
    let left = parts.next()?;
    let right = parts.next().unwrap_or(left);
    if parts.next().is_some() {
        return None;
    }

    let (_, start_row, start_column) = parse_a1_cell(left, "")?;
    let (_, end_row, end_column) = parse_a1_cell(right, "")?;
    Some(SheetRange {
        start_row: start_row.min(end_row),
        start_column: start_column.min(end_column),
        end_row: start_row.max(end_row),
        end_column: start_column.max(end_column),
    })
}

pub(crate) fn parse_sheet_cell_ref(reference: &str, default_sheet: &str) -> Option<SheetCellRef> {
    let (sheet_name, row, column) = parse_a1_cell(reference, default_sheet)?;
    SheetCellRef::new(&sheet_name, row, column)
}

pub(crate) fn build_data_table_region(
    sheet_name: &str,
    sheet_part: &str,
    anchor_address: &str,
    range_address: &str,
    r1: Option<&str>,
    r2: Option<&str>,
    is_two_dimensional: bool,
    dtr: bool,
) -> DataTableRegion {
    let range = parse_sheet_range(range_address).unwrap_or(SheetRange {
        start_row: 0,
        start_column: 0,
        end_row: -1,
        end_column: -1,
    });
    let formula_cell = if is_two_dimensional && range.start_row > 1 && range.start_column > 1 {
        SheetCellRef::new(sheet_name, range.start_row - 1, range.start_column - 1)
    } else if !is_two_dimensional && dtr && range.start_column > 1 {
        SheetCellRef::new(sheet_name, range.start_row, range.start_column - 1)
    } else if !is_two_dimensional && !dtr && range.start_row > 1 {
        SheetCellRef::new(sheet_name, range.start_row - 1, range.start_column)
    } else {
        None
    };
    let column_input_cell = r1.and_then(|reference| parse_sheet_cell_ref(reference, sheet_name));
    let row_input_cell = r2.and_then(|reference| parse_sheet_cell_ref(reference, sheet_name));
    let unsupported_reason = if range.cell_count() == 0 {
        Some("Data table range could not be parsed.".to_string())
    } else if formula_cell.is_none() {
        Some("Data table source formula could not be inferred.".to_string())
    } else if is_two_dimensional && (column_input_cell.is_none() || row_input_cell.is_none()) {
        Some("Data table input cells could not be parsed.".to_string())
    } else if !is_two_dimensional && column_input_cell.is_none() {
        Some("One-variable data table input cell could not be parsed.".to_string())
    } else {
        None
    };

    DataTableRegion {
        id: format!("{sheet_name}!{range_address}"),
        sheet_name: sheet_name.to_string(),
        sheet_part: sheet_part.to_string(),
        anchor_address: anchor_address.to_string(),
        range_address: range_address.to_string(),
        range,
        formula_cell,
        column_input_cell,
        row_input_cell,
        is_two_dimensional,
        dtr,
        unsupported_reason,
    }
}

pub(crate) fn summarize_data_tables(
    model: &Model<'_>,
    data_tables: &[DataTableRegion],
    changed_cells: &[CellId],
    requested_evaluation: bool,
) -> DataTableBenchmarkSummary {
    let data_table_count = data_tables.len();
    let data_table_cells = data_tables.iter().map(DataTableRegion::cell_count).sum();
    if data_table_count == 0 {
        return DataTableBenchmarkSummary::default();
    }

    let dirty_table_indexes = dirty_table_indexes(model, data_tables, changed_cells);
    let dirty_data_tables = dirty_table_indexes.len();
    let reused_data_table_cells = data_tables
        .iter()
        .enumerate()
        .filter(|(index, _)| !dirty_table_indexes.contains(index))
        .map(|(_, table)| table.cell_count())
        .sum();
    let unsupported_data_table_cells = data_tables
        .iter()
        .filter(|table| !table.is_parallel_validation_eligible())
        .map(DataTableRegion::cell_count)
        .sum();
    let diagnostics = if requested_evaluation {
        data_tables
            .iter()
            .filter(|table| !table.is_parallel_validation_eligible())
            .map(|table| {
                data_table_diagnostic(
                    model,
                    table,
                    table.cell_count(),
                    &KernelError::detail(
                        "ineligible_data_table",
                        table
                            .unsupported_reason
                            .clone()
                            .unwrap_or_else(|| "Data table shape is not eligible for kernel evaluation.".to_string()),
                    ),
                )
            })
            .collect()
    } else {
        Vec::new()
    };

    DataTableBenchmarkSummary {
        data_table_count,
        data_table_cells,
        dirty_data_tables,
        reused_data_table_cells,
        unsupported_data_table_cells,
        diagnostics,
        status: if requested_evaluation {
            DataTableEvaluationStatus::Unsupported
        } else {
            DataTableEvaluationStatus::MetadataOnly
        },
        ..DataTableBenchmarkSummary::default()
    }
}

pub(crate) fn evaluate_data_tables(
    model: &Model<'_>,
    data_tables: &[DataTableRegion],
    changed_cells: &[CellId],
    requested_evaluation: bool,
) -> DataTableBenchmarkSummary {
    let mut summary =
        summarize_data_tables(model, data_tables, changed_cells, requested_evaluation);
    if data_tables.is_empty() {
        return summary;
    }

    if !requested_evaluation {
        return summary;
    }

    let dirty_table_indexes = dirty_table_indexes(model, data_tables, changed_cells);
    let should_evaluate_table =
        |index: usize| changed_cells.is_empty() || dirty_table_indexes.contains(&index);
    let table_indexes = data_tables
        .iter()
        .enumerate()
        .filter(|(index, table)| {
            should_evaluate_table(*index) && table.is_parallel_validation_eligible()
        })
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    summary.reused_data_table_cells = data_tables
        .iter()
        .enumerate()
        .filter(|(index, _)| !should_evaluate_table(*index))
        .map(|(_, table)| table.cell_count())
        .sum();
    if table_indexes.is_empty() {
        summary.status = DataTableEvaluationStatus::Unsupported;
        return summary;
    }

    let started = Instant::now();
    let parallelism = thread::available_parallelism()
        .map(|count| count.get())
        .unwrap_or(1)
        .min(table_indexes.len())
        .max(1);
    let chunk_size = table_indexes.len().div_ceil(parallelism);
    let workbook = model.workbook.clone();
    let language = model.get_language();
    let counts = thread::scope(|scope| {
        let mut handles = Vec::new();
        for chunk in table_indexes.chunks(chunk_size) {
            let worker_workbook = workbook.clone();
            let worker_language = language.clone();
            handles.push(scope.spawn(move || {
                let Ok(worker_model) = Model::from_workbook(worker_workbook, &worker_language)
                else {
                    let mut counts = ScenarioCounts {
                        unsupported: chunk.iter().map(|index| data_tables[*index].cell_count()).sum(),
                        ..ScenarioCounts::default()
                    };
                    for table_index in chunk {
                        let table = &data_tables[*table_index];
                        counts.diagnostics.push(data_table_diagnostic_without_model(
                            table,
                            table.cell_count(),
                            &KernelError::detail(
                                "worker_model_load_failed",
                                "Data table worker could not clone the workbook model for parallel evaluation.",
                            ),
                        ));
                    }
                    return counts;
                };
                let mut counts = ScenarioCounts::default();
                for table_index in chunk {
                    counts.add(evaluate_table_kernel(
                        &worker_model,
                        &data_tables[*table_index],
                    ));
                }
                counts
            }));
        }

        let mut total = ScenarioCounts::default();
        for handle in handles {
            match handle.join() {
                Ok(counts) => total.add(counts),
                Err(_) => total.unsupported += chunk_size,
            }
        }
        total
    });

    summary.data_table_eval_ms = started.elapsed().as_millis();
    summary.data_table_parallelism = parallelism;
    summary.evaluated_data_table_cells = counts.evaluated;
    summary.validated_data_table_cells = counts.validated;
    summary.mismatched_data_table_cells = counts.mismatched;
    summary.unsupported_data_table_cells += counts.unsupported;
    summary.diagnostics.extend(counts.diagnostics);
    summary.status = data_table_status(&summary);
    summary
}

fn build_table_scenarios(
    model: &Model<'_>,
    table: &DataTableRegion,
) -> Result<Vec<DataTableScenario>, KernelError> {
    let mut scenarios = Vec::new();

    let Some(formula_cell) = table.formula_cell.clone() else {
        return Err(KernelError::Unsupported);
    };
    let Some(column_input) = table.column_input_cell.clone() else {
        return Err(KernelError::Unsupported);
    };

    if resolve_cell(model, &formula_cell).is_none() || resolve_cell(model, &column_input).is_none()
    {
        return Err(KernelError::Unsupported);
    }
    if table.is_two_dimensional
        && !table
            .row_input_cell
            .as_ref()
            .is_some_and(|row_input| resolve_cell(model, row_input).is_some())
    {
        return Err(KernelError::Unsupported);
    }

    if table.is_two_dimensional {
        for row in table.range.start_row..=table.range.end_row {
            let Some(row_axis) =
                SheetCellRef::new(&table.sheet_name, row, table.range.start_column - 1)
            else {
                return Err(KernelError::Unsupported);
            };
            for column in table.range.start_column..=table.range.end_column {
                let Some(column_axis) =
                    SheetCellRef::new(&table.sheet_name, table.range.start_row - 1, column)
                else {
                    return Err(KernelError::Unsupported);
                };
                let Some(output_cell) = SheetCellRef::new(&table.sheet_name, row, column) else {
                    return Err(KernelError::Unsupported);
                };
                scenarios.push(DataTableScenario {
                    column_axis: Some(column_axis),
                    row_axis: Some(row_axis.clone()),
                    output_cell,
                });
            }
        }
    } else if table.dtr {
        for column in table.range.start_column..=table.range.end_column {
            let Some(column_axis) =
                SheetCellRef::new(&table.sheet_name, table.range.start_row - 1, column)
            else {
                return Err(KernelError::Unsupported);
            };
            let Some(output_cell) =
                SheetCellRef::new(&table.sheet_name, table.range.start_row, column)
            else {
                return Err(KernelError::Unsupported);
            };
            scenarios.push(DataTableScenario {
                column_axis: Some(column_axis),
                row_axis: None,
                output_cell,
            });
        }
    } else {
        for row in table.range.start_row..=table.range.end_row {
            let Some(row_axis) =
                SheetCellRef::new(&table.sheet_name, row, table.range.start_column - 1)
            else {
                return Err(KernelError::Unsupported);
            };
            let Some(output_cell) =
                SheetCellRef::new(&table.sheet_name, row, table.range.start_column)
            else {
                return Err(KernelError::Unsupported);
            };
            scenarios.push(DataTableScenario {
                column_axis: None,
                row_axis: Some(row_axis),
                output_cell,
            });
        }
    }

    Ok(scenarios)
}

fn evaluate_table_kernel(model: &Model<'_>, table: &DataTableRegion) -> ScenarioCounts {
    let scenarios = match build_table_scenarios(model, table) {
        Ok(scenarios) => scenarios,
        Err(_) => {
            return ScenarioCounts {
                unsupported: table.cell_count(),
                ..ScenarioCounts::default()
            }
        }
    };

    let scenario_count = scenarios.len();
    let Some(formula_cell) = table
        .formula_cell
        .as_ref()
        .and_then(|cell| resolve_cell(model, cell))
    else {
        return ScenarioCounts {
            unsupported: table.cell_count(),
            ..ScenarioCounts::default()
        };
    };
    let Some(column_input_cell) = table
        .column_input_cell
        .as_ref()
        .and_then(|cell| resolve_cell(model, cell))
    else {
        return ScenarioCounts {
            unsupported: table.cell_count(),
            ..ScenarioCounts::default()
        };
    };
    let row_input_cell = table
        .row_input_cell
        .as_ref()
        .and_then(|cell| resolve_cell(model, cell));

    let column_inputs = match comparable_values_for_optional_cells(
        model,
        scenarios
            .iter()
            .map(|scenario| scenario.column_axis.as_ref()),
    ) {
        Ok(values) => values,
        Err(_) => {
            return ScenarioCounts {
                unsupported: table.cell_count(),
                ..ScenarioCounts::default()
            }
        }
    };
    let row_inputs = match comparable_values_for_optional_cells(
        model,
        scenarios.iter().map(|scenario| scenario.row_axis.as_ref()),
    ) {
        Ok(values) => values,
        Err(_) => {
            return ScenarioCounts {
                unsupported: table.cell_count(),
                ..ScenarioCounts::default()
            }
        }
    };
    let expected_values = match comparable_values_for_cells(
        model,
        scenarios.iter().map(|scenario| &scenario.output_cell),
    ) {
        Ok(values) => values,
        Err(_) => {
            return ScenarioCounts {
                unsupported: table.cell_count(),
                ..ScenarioCounts::default()
            }
        }
    };

    let mut overrides = HashMap::new();
    if (table.is_two_dimensional || table.dtr) && column_inputs.iter().any(Option::is_some) {
        let Some(values) = collect_optional_values(column_inputs) else {
            return ScenarioCounts {
                unsupported: table.cell_count(),
                ..ScenarioCounts::default()
            };
        };
        overrides.insert(
            column_input_cell,
            KernelValue::from_comparable_values(values),
        );
    }
    if row_inputs.iter().any(Option::is_some) {
        let Some(values) = collect_optional_values(row_inputs) else {
            return ScenarioCounts {
                unsupported: table.cell_count(),
                ..ScenarioCounts::default()
            };
        };
        let input_cell = if table.is_two_dimensional {
            let Some(row_input_cell) = row_input_cell else {
                return ScenarioCounts {
                    unsupported: table.cell_count(),
                    ..ScenarioCounts::default()
                };
            };
            row_input_cell
        } else {
            column_input_cell
        };
        overrides.insert(input_cell, KernelValue::from_comparable_values(values));
    }
    let mut context = KernelContext {
        model,
        scenario_count,
        overrides,
        memo: HashMap::new(),
        visiting: HashSet::new(),
        iterative_values: HashMap::new(),
        used_iteration: false,
        iteration_max_delta: 0.0,
    };

    let actual_value = match context.eval_cell_with_iteration(formula_cell) {
        Ok(value) => value,
        Err(_) => {
            return ScenarioCounts {
                unsupported: table.cell_count(),
                ..ScenarioCounts::default()
            }
        }
    };
    let actual_values = match actual_value.into_comparable_values(&mut context) {
        Ok(values) if values.len() == scenario_count => values,
        _ => {
            return ScenarioCounts {
                unsupported: table.cell_count(),
                ..ScenarioCounts::default()
            }
        }
    };

    let mut counts = ScenarioCounts::default();
    counts.evaluated = scenario_count;
    for (actual, expected) in actual_values.iter().zip(expected_values.iter()) {
        if values_match(actual, expected) {
            counts.validated += 1;
        } else {
            counts.mismatched += 1;
        }
    }
    counts
}

#[derive(Clone, Debug)]
enum KernelValue {
    Values(Vec<ComparableValue>),
    Range(KernelRange),
}

#[derive(Clone, Debug)]
struct KernelRange {
    cells: Vec<CellId>,
    rows: usize,
    columns: usize,
}

struct CriteriaSet {
    values_by_cell: Vec<Vec<ComparableValue>>,
    criteria: Vec<ComparableValue>,
}

#[derive(Debug)]
enum KernelError {
    Unsupported,
}

struct KernelContext<'a, 'm> {
    model: &'a Model<'m>,
    scenario_count: usize,
    overrides: HashMap<CellId, KernelValue>,
    memo: HashMap<CellId, KernelValue>,
    visiting: HashSet<CellId>,
    iterative_values: HashMap<CellId, KernelValue>,
    used_iteration: bool,
    iteration_max_delta: f64,
}

impl<'a, 'm> KernelContext<'a, 'm> {
    fn eval_cell_with_iteration(&mut self, cell: CellId) -> Result<KernelValue, KernelError> {
        let mut value = self.eval_cell(cell)?;
        if !self.used_iteration {
            return Ok(value);
        }

        for _ in 1..ITERATIVE_MAX_ITERATIONS {
            self.memo.clear();
            self.visiting.clear();
            self.used_iteration = false;
            self.iteration_max_delta = 0.0;
            value = self.eval_cell(cell)?;
            if self.iteration_max_delta <= ITERATIVE_MAX_CHANGE {
                break;
            }
        }

        Ok(value)
    }

    fn eval_cell(&mut self, cell: CellId) -> Result<KernelValue, KernelError> {
        if let Some(value) = self.overrides.get(&cell) {
            return Ok(value.clone());
        }
        if let Some(value) = self.memo.get(&cell) {
            return Ok(value.clone());
        }
        if !self.visiting.insert(cell) {
            self.used_iteration = true;
            return Ok(self.iterative_seed(cell));
        }

        let value = if let Some(node) = self.formula_node_for_cell(cell) {
            self.eval_node(&node, cell)?
        } else {
            self.static_cell_value(cell)?
        };

        self.visiting.remove(&cell);
        if self.used_iteration {
            self.update_iterative_value(cell, &value);
        }
        self.memo.insert(cell, value.clone());
        Ok(value)
    }

    fn eval_node(&mut self, node: &Node, context: CellId) -> Result<KernelValue, KernelError> {
        match node {
            Node::BooleanKind(value) => Ok(KernelValue::boolean(*value, self.scenario_count)),
            Node::NumberKind(value) => Ok(KernelValue::number(*value, self.scenario_count)),
            Node::StringKind(value) => Ok(KernelValue::string(value.clone(), self.scenario_count)),
            Node::ReferenceKind {
                sheet_index,
                absolute_row,
                absolute_column,
                row,
                column,
                ..
            } => self.eval_cell(CellId::new(
                *sheet_index,
                absolute_coord(*row, context.row, *absolute_row),
                absolute_coord(*column, context.column, *absolute_column),
            )),
            Node::RangeKind { .. } => self.range_from_node(node, context).map(KernelValue::Range),
            Node::OpSumKind { kind, left, right } => {
                let left = self.eval_node(left, context)?;
                let right = self.eval_node(right, context)?;
                let left = coerce_numbers(self, left)?;
                let right = coerce_numbers(self, right)?;
                let values = left
                    .iter()
                    .zip(right.iter())
                    .map(|(left, right)| match kind {
                        OpSum::Add => left + right,
                        OpSum::Minus => left - right,
                    })
                    .map(ComparableValue::Number)
                    .collect();
                Ok(KernelValue::Values(values))
            }
            Node::OpProductKind { kind, left, right } => {
                let left = self.eval_node(left, context)?;
                let right = self.eval_node(right, context)?;
                let left = coerce_numbers(self, left)?;
                let right = coerce_numbers(self, right)?;
                let values = left
                    .iter()
                    .zip(right.iter())
                    .map(|(left, right)| match kind {
                        OpProduct::Times => left * right,
                        OpProduct::Divide => left / right,
                    })
                    .map(ComparableValue::Number)
                    .collect();
                Ok(KernelValue::Values(values))
            }
            Node::OpPowerKind { left, right } => {
                let left = self.eval_node(left, context)?;
                let right = self.eval_node(right, context)?;
                let left = coerce_numbers(self, left)?;
                let right = coerce_numbers(self, right)?;
                Ok(KernelValue::Values(
                    left.iter()
                        .zip(right.iter())
                        .map(|(left, right)| ComparableValue::Number(left.powf(*right)))
                        .collect(),
                ))
            }
            Node::CompareKind { kind, left, right } => {
                let left = self
                    .eval_node(left, context)?
                    .into_comparable_values(self)?;
                let right = self
                    .eval_node(right, context)?
                    .into_comparable_values(self)?;
                Ok(KernelValue::Values(
                    left.iter()
                        .zip(right.iter())
                        .map(|(left, right)| {
                            ComparableValue::Boolean(compare_values(left, right, kind))
                        })
                        .collect(),
                ))
            }
            Node::UnaryKind { kind, right } => {
                let right_value = self.eval_node(right, context)?;
                let numbers = coerce_numbers(self, right_value)?;
                Ok(KernelValue::Values(
                    numbers
                        .into_iter()
                        .map(|value| match kind {
                            OpUnary::Minus => -value,
                            OpUnary::Percentage => value / 100.0,
                        })
                        .map(ComparableValue::Number)
                        .collect(),
                ))
            }
            Node::OpConcatenateKind { left, right } => {
                let left = self
                    .eval_node(left, context)?
                    .into_comparable_values(self)?;
                let right = self
                    .eval_node(right, context)?
                    .into_comparable_values(self)?;
                Ok(KernelValue::Values(
                    left.iter()
                        .zip(right.iter())
                        .map(|(left, right)| {
                            ComparableValue::String(format!(
                                "{}{}",
                                comparable_to_text(left),
                                comparable_to_text(right)
                            ))
                        })
                        .collect(),
                ))
            }
            Node::FunctionKind { kind, args } => {
                self.eval_function(&format!("{kind:?}"), args, context)
            }
            Node::ImplicitIntersection { child, .. } => match self.eval_node(child, context)? {
                KernelValue::Range(range) => self.eval_implicit_intersection(range, context),
                value => Ok(value),
            },
            Node::EmptyArgKind => Ok(KernelValue::empty(self.scenario_count)),
            Node::WrongReferenceKind { .. }
            | Node::WrongRangeKind { .. }
            | Node::OpRangeKind { .. }
            | Node::InvalidFunctionKind { .. }
            | Node::ArrayKind(_)
            | Node::DefinedNameKind(_)
            | Node::TableNameKind(_)
            | Node::WrongVariableKind(_)
            | Node::ErrorKind(_)
            | Node::ParseErrorKind { .. } => Err(KernelError::Unsupported),
        }
    }

    fn eval_function(
        &mut self,
        function_name: &str,
        args: &[Node],
        context: CellId,
    ) -> Result<KernelValue, KernelError> {
        match function_name {
            "Sum" => self.eval_numeric_aggregate(args, context, NumericAggregate::Sum),
            "Product" => self.eval_numeric_aggregate(args, context, NumericAggregate::Product),
            "Min" => self.eval_numeric_aggregate(args, context, NumericAggregate::Min),
            "Max" => self.eval_numeric_aggregate(args, context, NumericAggregate::Max),
            "Average" => self.eval_average(args, context),
            "Count" => self.eval_count(args, context, CountMode::Numbers),
            "Counta" => self.eval_count(args, context, CountMode::NonEmpty),
            "Countif" => self.eval_countif(args, context),
            "Countifs" => self.eval_countifs(args, context),
            "Averageif" => self.eval_averageif(args, context),
            "Averageifs" => self.eval_ifs_aggregate(args, context, IfsAggregate::Average),
            "Abs" => self.eval_single_number_function(args, context, f64::abs),
            "Round" => self.eval_round(args, context, RoundMode::Nearest),
            "Roundup" => self.eval_round(args, context, RoundMode::Up),
            "Rounddown" => self.eval_round(args, context, RoundMode::Down),
            "Power" => self.eval_power(args, context),
            "And" => self.eval_boolean_aggregate(args, context, BooleanAggregate::And),
            "Not" => self.eval_not(args, context),
            "Or" => self.eval_boolean_aggregate(args, context, BooleanAggregate::Or),
            "If" => self.eval_if(args, context),
            "Choose" => self.eval_choose(args, context),
            "Index" => self.eval_index(args, context),
            "Match" => self.eval_match(args, context),
            "Offset" => self.eval_offset(args, context),
            "Vlookup" => self.eval_table_lookup(args, context, LookupDirection::Vertical),
            "Hlookup" => self.eval_table_lookup(args, context, LookupDirection::Horizontal),
            "Concat" | "Concatenate" => self.eval_concatenate(args, context),
            "Sumif" => self.eval_sumif(args, context),
            "Sumifs" => self.eval_ifs_aggregate(args, context, IfsAggregate::Sum),
            "Minifs" => self.eval_ifs_aggregate(args, context, IfsAggregate::Min),
            "Maxifs" => self.eval_ifs_aggregate(args, context, IfsAggregate::Max),
            "Npv" => self.eval_npv(args, context),
            "Irr" => self.eval_irr(args, context),
            "Pmt" => self.eval_pmt(args, context),
            _ => Err(KernelError::Unsupported),
        }
    }

    fn eval_numeric_aggregate(
        &mut self,
        args: &[Node],
        context: CellId,
        aggregate: NumericAggregate,
    ) -> Result<KernelValue, KernelError> {
        let mut result = match aggregate {
            NumericAggregate::Sum => vec![0.0; self.scenario_count],
            NumericAggregate::Product => vec![1.0; self.scenario_count],
            NumericAggregate::Min => vec![f64::INFINITY; self.scenario_count],
            NumericAggregate::Max => vec![f64::NEG_INFINITY; self.scenario_count],
        };
        let mut saw_value = false;

        for arg in args {
            let value = self.eval_node(arg, context)?;
            for numbers in flatten_numbers(self, value)? {
                saw_value = true;
                for (slot, value) in result.iter_mut().zip(numbers.iter()) {
                    match aggregate {
                        NumericAggregate::Sum => *slot += value,
                        NumericAggregate::Product => *slot *= value,
                        NumericAggregate::Min => *slot = slot.min(*value),
                        NumericAggregate::Max => *slot = slot.max(*value),
                    }
                }
            }
        }

        if !saw_value {
            result.fill(0.0);
        }

        Ok(KernelValue::Values(
            result.into_iter().map(ComparableValue::Number).collect(),
        ))
    }

    fn eval_average(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.is_empty() {
            return Err(KernelError::Unsupported);
        }

        let mut totals = vec![0.0; self.scenario_count];
        let mut counts = vec![0_usize; self.scenario_count];
        for arg in args {
            let value = self.eval_node(arg, context)?;
            for numbers in flatten_numbers(self, value)? {
                for (index, number) in numbers.iter().enumerate() {
                    totals[index] += number;
                    counts[index] += 1;
                }
            }
        }

        if counts.iter().any(|count| *count == 0) {
            return Err(KernelError::Unsupported);
        }

        Ok(KernelValue::Values(
            totals
                .iter()
                .zip(counts.iter())
                .map(|(total, count)| ComparableValue::Number(total / *count as f64))
                .collect(),
        ))
    }

    fn eval_count(
        &mut self,
        args: &[Node],
        context: CellId,
        mode: CountMode,
    ) -> Result<KernelValue, KernelError> {
        let mut counts = vec![0_usize; self.scenario_count];
        for arg in args {
            let value = self.eval_node(arg, context)?;
            for values in flatten_comparable_values(self, value)? {
                for (index, value) in values.iter().enumerate() {
                    let should_count = match mode {
                        CountMode::Numbers => matches!(value, ComparableValue::Number(_)),
                        CountMode::NonEmpty => !matches!(value, ComparableValue::None),
                    };
                    if should_count {
                        counts[index] += 1;
                    }
                }
            }
        }

        Ok(KernelValue::Values(
            counts
                .into_iter()
                .map(|count| ComparableValue::Number(count as f64))
                .collect(),
        ))
    }

    fn eval_countif(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.len() != 2 {
            return Err(KernelError::Unsupported);
        }
        self.eval_countifs(args, context)
    }

    fn eval_countifs(
        &mut self,
        args: &[Node],
        context: CellId,
    ) -> Result<KernelValue, KernelError> {
        if args.len() < 2 || args.len() % 2 != 0 {
            return Err(KernelError::Unsupported);
        }

        let first_range = self.range_from_node(&args[0], context)?;
        let criteria_sets =
            self.collect_criteria_sets(args, context, first_range.rows, first_range.columns)?;
        let mut counts = vec![0_usize; self.scenario_count];
        for offset in 0..first_range.cells.len() {
            for scenario_index in 0..self.scenario_count {
                if criteria_sets_match(&criteria_sets, offset, scenario_index)? {
                    counts[scenario_index] += 1;
                }
            }
        }

        Ok(KernelValue::Values(
            counts
                .into_iter()
                .map(|count| ComparableValue::Number(count as f64))
                .collect(),
        ))
    }

    fn eval_sumif(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        let normalized = match args.len() {
            2 => vec![args[0].clone(), args[0].clone(), args[1].clone()],
            3 => vec![args[2].clone(), args[0].clone(), args[1].clone()],
            _ => return Err(KernelError::Unsupported),
        };
        self.eval_ifs_aggregate(&normalized, context, IfsAggregate::Sum)
    }

    fn eval_averageif(
        &mut self,
        args: &[Node],
        context: CellId,
    ) -> Result<KernelValue, KernelError> {
        let normalized = match args.len() {
            2 => vec![args[0].clone(), args[0].clone(), args[1].clone()],
            3 => vec![args[2].clone(), args[0].clone(), args[1].clone()],
            _ => return Err(KernelError::Unsupported),
        };
        self.eval_ifs_aggregate(&normalized, context, IfsAggregate::Average)
    }

    fn eval_ifs_aggregate(
        &mut self,
        args: &[Node],
        context: CellId,
        aggregate: IfsAggregate,
    ) -> Result<KernelValue, KernelError> {
        if args.len() < 3 || args.len() % 2 == 0 {
            return Err(KernelError::Unsupported);
        }

        let value_range = self.range_from_node(&args[0], context)?;
        let values_by_cell = value_range
            .cells
            .iter()
            .map(|cell| self.eval_cell(*cell)?.into_comparable_values(self))
            .collect::<Result<Vec<_>, _>>()?;
        let criteria_sets =
            self.collect_criteria_sets(&args[1..], context, value_range.rows, value_range.columns)?;

        let mut totals = vec![0.0; self.scenario_count];
        let mut counts = vec![0_usize; self.scenario_count];
        let mut mins = vec![f64::INFINITY; self.scenario_count];
        let mut maxes = vec![f64::NEG_INFINITY; self.scenario_count];
        for (offset, values) in values_by_cell.iter().enumerate() {
            for scenario_index in 0..self.scenario_count {
                if !criteria_sets_match(&criteria_sets, offset, scenario_index)? {
                    continue;
                }
                let ComparableValue::Number(value) = values[scenario_index] else {
                    continue;
                };
                totals[scenario_index] += value;
                counts[scenario_index] += 1;
                mins[scenario_index] = mins[scenario_index].min(value);
                maxes[scenario_index] = maxes[scenario_index].max(value);
            }
        }

        let mut result = Vec::with_capacity(self.scenario_count);
        for scenario_index in 0..self.scenario_count {
            let value = match aggregate {
                IfsAggregate::Sum => totals[scenario_index],
                IfsAggregate::Average => {
                    if counts[scenario_index] == 0 {
                        return Err(KernelError::Unsupported);
                    }
                    totals[scenario_index] / counts[scenario_index] as f64
                }
                IfsAggregate::Min => {
                    if counts[scenario_index] == 0 {
                        0.0
                    } else {
                        mins[scenario_index]
                    }
                }
                IfsAggregate::Max => {
                    if counts[scenario_index] == 0 {
                        0.0
                    } else {
                        maxes[scenario_index]
                    }
                }
            };
            result.push(ComparableValue::Number(value));
        }

        Ok(KernelValue::Values(result))
    }

    fn collect_criteria_sets(
        &mut self,
        args: &[Node],
        context: CellId,
        expected_rows: usize,
        expected_columns: usize,
    ) -> Result<Vec<CriteriaSet>, KernelError> {
        if args.len() < 2 || args.len() % 2 != 0 {
            return Err(KernelError::Unsupported);
        }

        let mut criteria_sets = Vec::new();
        for pair in args.chunks_exact(2) {
            let range = self.range_from_node(&pair[0], context)?;
            if range.rows != expected_rows || range.columns != expected_columns {
                return Err(KernelError::Unsupported);
            }
            let criteria = self
                .eval_node(&pair[1], context)?
                .into_comparable_values(self)?;
            let values_by_cell = range
                .cells
                .iter()
                .map(|cell| self.eval_cell(*cell)?.into_comparable_values(self))
                .collect::<Result<Vec<_>, _>>()?;
            criteria_sets.push(CriteriaSet {
                values_by_cell,
                criteria,
            });
        }
        Ok(criteria_sets)
    }

    fn eval_boolean_aggregate(
        &mut self,
        args: &[Node],
        context: CellId,
        aggregate: BooleanAggregate,
    ) -> Result<KernelValue, KernelError> {
        if args.is_empty() {
            return Err(KernelError::Unsupported);
        }

        let mut result = match aggregate {
            BooleanAggregate::And => vec![true; self.scenario_count],
            BooleanAggregate::Or => vec![false; self.scenario_count],
        };
        let mut saw_value = false;
        for arg in args {
            let value = self.eval_node(arg, context)?;
            for values in flatten_comparable_values(self, value)? {
                saw_value = true;
                for (slot, value) in result.iter_mut().zip(values.iter()) {
                    let boolean = coerce_comparable_to_boolean(value)?;
                    match aggregate {
                        BooleanAggregate::And => *slot = *slot && boolean,
                        BooleanAggregate::Or => *slot = *slot || boolean,
                    }
                }
            }
        }

        if !saw_value {
            return Err(KernelError::Unsupported);
        }

        Ok(KernelValue::Values(
            result.into_iter().map(ComparableValue::Boolean).collect(),
        ))
    }

    fn eval_not(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.len() != 1 {
            return Err(KernelError::Unsupported);
        }
        let value = self.eval_node(&args[0], context)?;
        Ok(KernelValue::Values(
            coerce_booleans(self, value)?
                .into_iter()
                .map(|value| ComparableValue::Boolean(!value))
                .collect(),
        ))
    }

    fn eval_concatenate(
        &mut self,
        args: &[Node],
        context: CellId,
    ) -> Result<KernelValue, KernelError> {
        if args.is_empty() {
            return Err(KernelError::Unsupported);
        }

        let mut result = vec![String::new(); self.scenario_count];
        for arg in args {
            let value = self.eval_node(arg, context)?;
            for values in flatten_comparable_values(self, value)? {
                for (slot, value) in result.iter_mut().zip(values.iter()) {
                    slot.push_str(&comparable_to_text(value));
                }
            }
        }

        Ok(KernelValue::Values(
            result.into_iter().map(ComparableValue::String).collect(),
        ))
    }

    fn eval_single_number_function(
        &mut self,
        args: &[Node],
        context: CellId,
        function: impl Fn(f64) -> f64,
    ) -> Result<KernelValue, KernelError> {
        if args.len() != 1 {
            return Err(KernelError::Unsupported);
        }
        let value = self.eval_node(&args[0], context)?;
        let numbers = coerce_numbers(self, value)?;
        Ok(KernelValue::Values(
            numbers
                .into_iter()
                .map(function)
                .map(ComparableValue::Number)
                .collect(),
        ))
    }

    fn eval_round(
        &mut self,
        args: &[Node],
        context: CellId,
        mode: RoundMode,
    ) -> Result<KernelValue, KernelError> {
        if args.len() != 2 {
            return Err(KernelError::Unsupported);
        }
        let number_value = self.eval_node(&args[0], context)?;
        let digits_value = self.eval_node(&args[1], context)?;
        let numbers = coerce_numbers(self, number_value)?;
        let digits = coerce_numbers(self, digits_value)?;
        Ok(KernelValue::Values(
            numbers
                .iter()
                .zip(digits.iter())
                .map(|(number, digits)| {
                    let factor = 10_f64.powi(*digits as i32);
                    let rounded = match mode {
                        RoundMode::Nearest => (number * factor).round() / factor,
                        RoundMode::Up => {
                            if *number >= 0.0 {
                                (number * factor).ceil() / factor
                            } else {
                                (number * factor).floor() / factor
                            }
                        }
                        RoundMode::Down => {
                            if *number >= 0.0 {
                                (number * factor).floor() / factor
                            } else {
                                (number * factor).ceil() / factor
                            }
                        }
                    };
                    ComparableValue::Number(rounded)
                })
                .collect(),
        ))
    }

    fn eval_power(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.len() != 2 {
            return Err(KernelError::Unsupported);
        }
        let left = self.eval_node(&args[0], context)?;
        let right = self.eval_node(&args[1], context)?;
        let left = coerce_numbers(self, left)?;
        let right = coerce_numbers(self, right)?;
        Ok(KernelValue::Values(
            left.iter()
                .zip(right.iter())
                .map(|(left, right)| ComparableValue::Number(left.powf(*right)))
                .collect(),
        ))
    }

    fn eval_if(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.len() < 2 || args.len() > 3 {
            return Err(KernelError::Unsupported);
        }
        let condition_value = self.eval_node(&args[0], context)?;
        let conditions = coerce_booleans(self, condition_value)?;
        let true_values = self
            .eval_node(&args[1], context)?
            .into_comparable_values(self)?;
        let false_values = if args.len() == 3 {
            self.eval_node(&args[2], context)?
                .into_comparable_values(self)?
        } else {
            vec![ComparableValue::Boolean(false); self.scenario_count]
        };

        Ok(KernelValue::Values(
            conditions
                .iter()
                .enumerate()
                .map(|(index, condition)| {
                    if *condition {
                        true_values[index].clone()
                    } else {
                        false_values[index].clone()
                    }
                })
                .collect(),
        ))
    }

    fn eval_choose(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.len() < 2 {
            return Err(KernelError::Unsupported);
        }

        let index_value = self.eval_node(&args[0], context)?;
        let indexes = coerce_numbers(self, index_value)?
            .into_iter()
            .map(|value| value.trunc() as isize)
            .collect::<Vec<_>>();

        let first_index = indexes[0];
        if indexes.iter().all(|index| *index == first_index) {
            let branch_index = first_index - 1;
            if branch_index < 0 || branch_index as usize >= args.len() - 1 {
                return Err(KernelError::Unsupported);
            }
            return self.eval_node(&args[branch_index as usize + 1], context);
        }

        let branches = args[1..]
            .iter()
            .map(|arg| self.eval_node(arg, context)?.into_comparable_values(self))
            .collect::<Result<Vec<_>, _>>()?;
        let mut result = Vec::with_capacity(self.scenario_count);
        for (scenario_index, index) in indexes.iter().enumerate() {
            let branch_index = *index - 1;
            if branch_index < 0 || branch_index as usize >= branches.len() {
                return Err(KernelError::Unsupported);
            }
            result.push(branches[branch_index as usize][scenario_index].clone());
        }
        Ok(KernelValue::Values(result))
    }

    fn eval_index(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.len() < 2 || args.len() > 3 {
            return Err(KernelError::Unsupported);
        }
        let range = self.range_from_node(&args[0], context)?;
        let row_numbers = {
            let value = self.eval_node(&args[1], context)?;
            coerce_numbers(self, value)?
        };
        let column_numbers = if args.len() == 3 {
            let value = self.eval_node(&args[2], context)?;
            coerce_numbers(self, value)?
        } else {
            vec![1.0; self.scenario_count]
        };

        let mut result = Vec::with_capacity(self.scenario_count);
        for index in 0..self.scenario_count {
            let row_index = row_numbers[index].round() as isize - 1;
            let column_index = column_numbers[index].round() as isize - 1;
            if row_index < 0
                || column_index < 0
                || row_index as usize >= range.rows
                || column_index as usize >= range.columns
            {
                return Err(KernelError::Unsupported);
            }
            let cell = range.cells[row_index as usize * range.columns + column_index as usize];
            let value = self.eval_cell(cell)?.into_comparable_values(self)?;
            result.push(value[index].clone());
        }
        Ok(KernelValue::Values(result))
    }

    fn eval_match(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.len() < 2 || args.len() > 3 {
            return Err(KernelError::Unsupported);
        }
        let lookup_values = self
            .eval_node(&args[0], context)?
            .into_comparable_values(self)?;
        let range = self.range_from_node(&args[1], context)?;
        if range.rows != 1 && range.columns != 1 {
            return Err(KernelError::Unsupported);
        }
        let match_types = if args.len() == 3 {
            let value = self.eval_node(&args[2], context)?;
            coerce_numbers(self, value)?
        } else {
            vec![1.0; self.scenario_count]
        };

        let range_values = range
            .cells
            .iter()
            .map(|cell| self.eval_cell(*cell)?.into_comparable_values(self))
            .collect::<Result<Vec<_>, _>>()?;
        let mut result = Vec::with_capacity(self.scenario_count);
        for scenario_index in 0..self.scenario_count {
            let lookup = &lookup_values[scenario_index];
            let match_type = match_types[scenario_index].round() as i32;
            let position = match match_type {
                0 => range_values
                    .iter()
                    .position(|candidate| values_match(&candidate[scenario_index], lookup)),
                1 => approximate_match_position(&range_values, lookup, scenario_index, true),
                -1 => approximate_match_position(&range_values, lookup, scenario_index, false),
                _ => None,
            };
            let Some(position) = position else {
                return Err(KernelError::Unsupported);
            };
            result.push(ComparableValue::Number(position as f64 + 1.0));
        }
        Ok(KernelValue::Values(result))
    }

    fn eval_table_lookup(
        &mut self,
        args: &[Node],
        context: CellId,
        direction: LookupDirection,
    ) -> Result<KernelValue, KernelError> {
        if args.len() < 3 || args.len() > 4 {
            return Err(KernelError::Unsupported);
        }

        let lookup_values = self
            .eval_node(&args[0], context)?
            .into_comparable_values(self)?;
        let range = self.range_from_node(&args[1], context)?;
        let index_value = self.eval_node(&args[2], context)?;
        let indexes = coerce_numbers(self, index_value)?;
        let sorted_flags = if args.len() == 4 {
            let sorted_value = self.eval_node(&args[3], context)?;
            coerce_booleans(self, sorted_value)?
        } else {
            vec![true; self.scenario_count]
        };
        let range_values = range
            .cells
            .iter()
            .map(|cell| self.eval_cell(*cell)?.into_comparable_values(self))
            .collect::<Result<Vec<_>, _>>()?;

        let lookup_count = match direction {
            LookupDirection::Vertical => range.rows,
            LookupDirection::Horizontal => range.columns,
        };
        let result_limit = match direction {
            LookupDirection::Vertical => range.columns,
            LookupDirection::Horizontal => range.rows,
        };
        if lookup_count == 0 || result_limit == 0 {
            return Err(KernelError::Unsupported);
        }

        let mut result = Vec::with_capacity(self.scenario_count);
        for scenario_index in 0..self.scenario_count {
            let result_index = indexes[scenario_index].floor() as isize - 1;
            if result_index < 0 || result_index as usize >= result_limit {
                return Err(KernelError::Unsupported);
            }

            let position = match direction {
                LookupDirection::Vertical => lookup_position(
                    &range_values,
                    (0..range.rows).map(|row| row * range.columns),
                    &lookup_values[scenario_index],
                    scenario_index,
                    sorted_flags[scenario_index],
                ),
                LookupDirection::Horizontal => lookup_position(
                    &range_values,
                    0..range.columns,
                    &lookup_values[scenario_index],
                    scenario_index,
                    sorted_flags[scenario_index],
                ),
            };
            let Some(position) = position else {
                return Err(KernelError::Unsupported);
            };
            let cell_index = match direction {
                LookupDirection::Vertical => position * range.columns + result_index as usize,
                LookupDirection::Horizontal => result_index as usize * range.columns + position,
            };
            result.push(range_values[cell_index][scenario_index].clone());
        }

        Ok(KernelValue::Values(result))
    }

    fn eval_offset(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.len() < 3 || args.len() > 5 {
            return Err(KernelError::Unsupported);
        }
        let base = self.range_from_node(&args[0], context)?;
        let rows_value = self.eval_node(&args[1], context)?;
        let columns_value = self.eval_node(&args[2], context)?;
        let row_offset = constant_i32(coerce_numbers(self, rows_value)?)?;
        let column_offset = constant_i32(coerce_numbers(self, columns_value)?)?;
        let height = if args.len() >= 4 {
            let height_value = self.eval_node(&args[3], context)?;
            constant_i32(coerce_numbers(self, height_value)?)?
        } else {
            base.rows as i32
        };
        let width = if args.len() == 5 {
            let width_value = self.eval_node(&args[4], context)?;
            constant_i32(coerce_numbers(self, width_value)?)?
        } else {
            base.columns as i32
        };
        if height <= 0 || width <= 0 {
            return Err(KernelError::Unsupported);
        }

        let start = base
            .cells
            .first()
            .copied()
            .ok_or(KernelError::Unsupported)?;
        Ok(KernelValue::Range(cells_for_rect(
            start.sheet,
            start.row + row_offset,
            start.column + column_offset,
            start.row + row_offset + height - 1,
            start.column + column_offset + width - 1,
        )))
    }

    fn eval_npv(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.len() < 2 {
            return Err(KernelError::Unsupported);
        }
        let rate_value = self.eval_node(&args[0], context)?;
        let rates = coerce_numbers(self, rate_value)?;
        let mut cash_flows = Vec::new();
        for arg in &args[1..] {
            let value = self.eval_node(arg, context)?;
            cash_flows.extend(flatten_numbers(self, value)?);
        }
        if cash_flows.is_empty() {
            return Err(KernelError::Unsupported);
        }

        let mut result = vec![0.0; self.scenario_count];
        for scenario_index in 0..self.scenario_count {
            let rate = rates[scenario_index];
            if rate <= -1.0 {
                return Err(KernelError::Unsupported);
            }
            let mut total = 0.0;
            for (period_index, values) in cash_flows.iter().enumerate() {
                total += values[scenario_index] / (1.0 + rate).powi(period_index as i32 + 1);
            }
            result[scenario_index] = total;
        }

        Ok(KernelValue::Values(
            result.into_iter().map(ComparableValue::Number).collect(),
        ))
    }

    fn eval_irr(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.is_empty() || args.len() > 2 {
            return Err(KernelError::Unsupported);
        }
        let values = self.eval_node(&args[0], context)?;
        let cash_flows = flatten_numbers(self, values)?;
        if cash_flows.is_empty() {
            return Err(KernelError::Unsupported);
        }
        let guesses = if args.len() == 2 {
            let guess_value = self.eval_node(&args[1], context)?;
            coerce_numbers(self, guess_value)?
        } else {
            vec![0.1; self.scenario_count]
        };

        let mut result = Vec::with_capacity(self.scenario_count);
        for scenario_index in 0..self.scenario_count {
            let flows = cash_flows
                .iter()
                .map(|values| values[scenario_index])
                .collect::<Vec<_>>();
            let Some(value) = solve_irr(&flows, guesses[scenario_index]) else {
                return Err(KernelError::Unsupported);
            };
            result.push(ComparableValue::Number(value));
        }
        Ok(KernelValue::Values(result))
    }

    fn eval_pmt(&mut self, args: &[Node], context: CellId) -> Result<KernelValue, KernelError> {
        if args.len() < 3 || args.len() > 5 {
            return Err(KernelError::Unsupported);
        }
        let rate_value = self.eval_node(&args[0], context)?;
        let nper_value = self.eval_node(&args[1], context)?;
        let pv_value = self.eval_node(&args[2], context)?;
        let rates = coerce_numbers(self, rate_value)?;
        let npers = coerce_numbers(self, nper_value)?;
        let pvs = coerce_numbers(self, pv_value)?;
        let fvs = if args.len() >= 4 {
            let fv_value = self.eval_node(&args[3], context)?;
            coerce_numbers(self, fv_value)?
        } else {
            vec![0.0; self.scenario_count]
        };
        let types = if args.len() == 5 {
            let type_value = self.eval_node(&args[4], context)?;
            coerce_numbers(self, type_value)?
        } else {
            vec![0.0; self.scenario_count]
        };

        let mut result = Vec::with_capacity(self.scenario_count);
        for index in 0..self.scenario_count {
            let rate = rates[index];
            let nper = npers[index];
            let pv = pvs[index];
            let fv = fvs[index];
            let payment_type = types[index];
            if nper == 0.0 || (payment_type != 0.0 && payment_type != 1.0) {
                return Err(KernelError::Unsupported);
            }
            let value = if rate == 0.0 {
                -(pv + fv) / nper
            } else {
                let factor = (1.0 + rate).powf(nper);
                -(pv * factor + fv) * rate / ((1.0 + rate * payment_type) * (factor - 1.0))
            };
            result.push(ComparableValue::Number(value));
        }
        Ok(KernelValue::Values(result))
    }

    fn eval_implicit_intersection(
        &mut self,
        range: KernelRange,
        context: CellId,
    ) -> Result<KernelValue, KernelError> {
        if range.cells.len() == 1 {
            return self.eval_cell(range.cells[0]);
        }
        if range.rows == 1 {
            if let Some(cell) = range
                .cells
                .iter()
                .find(|cell| cell.sheet == context.sheet && cell.column == context.column)
            {
                return self.eval_cell(*cell);
            }
        }
        if range.columns == 1 {
            if let Some(cell) = range
                .cells
                .iter()
                .find(|cell| cell.sheet == context.sheet && cell.row == context.row)
            {
                return self.eval_cell(*cell);
            }
        }
        Err(KernelError::Unsupported)
    }

    fn range_from_node(
        &mut self,
        node: &Node,
        context: CellId,
    ) -> Result<KernelRange, KernelError> {
        match node {
            Node::ReferenceKind {
                sheet_index,
                absolute_row,
                absolute_column,
                row,
                column,
                ..
            } => Ok(cells_for_rect(
                *sheet_index,
                absolute_coord(*row, context.row, *absolute_row),
                absolute_coord(*column, context.column, *absolute_column),
                absolute_coord(*row, context.row, *absolute_row),
                absolute_coord(*column, context.column, *absolute_column),
            )),
            Node::RangeKind {
                sheet_index,
                absolute_row1,
                absolute_column1,
                row1,
                column1,
                absolute_row2,
                absolute_column2,
                row2,
                column2,
                ..
            } => {
                let start_row = absolute_coord(*row1, context.row, *absolute_row1);
                let end_row = absolute_coord(*row2, context.row, *absolute_row2);
                let start_column = absolute_coord(*column1, context.column, *absolute_column1);
                let end_column = absolute_coord(*column2, context.column, *absolute_column2);
                Ok(cells_for_rect(
                    *sheet_index,
                    start_row,
                    start_column,
                    end_row,
                    end_column,
                ))
            }
            Node::FunctionKind { kind, .. } if format!("{kind:?}") == "Offset" => {
                match self.eval_node(node, context)? {
                    KernelValue::Range(range) => Ok(range),
                    KernelValue::Values(_) => Err(KernelError::Unsupported),
                }
            }
            Node::FunctionKind { kind, args } if format!("{kind:?}") == "Choose" => {
                self.range_from_choose(args, context)
            }
            _ => Err(KernelError::Unsupported),
        }
    }

    fn range_from_choose(
        &mut self,
        args: &[Node],
        context: CellId,
    ) -> Result<KernelRange, KernelError> {
        if args.len() < 2 {
            return Err(KernelError::Unsupported);
        }
        let index_value = self.eval_node(&args[0], context)?;
        let index = constant_i32(coerce_numbers(self, index_value)?)? - 1;
        if index < 0 || index as usize >= args.len() - 1 {
            return Err(KernelError::Unsupported);
        }
        self.range_from_node(&args[index as usize + 1], context)
    }

    fn formula_node_for_cell(&self, cell: CellId) -> Option<Node> {
        let formula_index = self
            .model
            .workbook
            .worksheet(cell.sheet)
            .ok()?
            .cell(cell.row, cell.column)?
            .get_formula()?;
        self.model
            .parsed_formulas
            .get(cell.sheet as usize)?
            .get(formula_index as usize)
            .cloned()
    }

    fn static_cell_value(&self, cell: CellId) -> Result<KernelValue, KernelError> {
        let value = self
            .model
            .get_cell_value_by_index(cell.sheet, cell.row, cell.column)
            .map_err(|_| KernelError::Unsupported)?;
        Ok(KernelValue::from_comparable_value(
            ComparableValue::from(value),
            self.scenario_count,
        ))
    }

    fn iterative_seed(&mut self, cell: CellId) -> KernelValue {
        if let Some(value) = self.iterative_values.get(&cell) {
            return value.clone();
        }

        let value = self.static_cell_value(cell).unwrap_or_else(|_| {
            KernelValue::from_comparable_value(ComparableValue::Number(0.0), self.scenario_count)
        });
        let value = match value {
            KernelValue::Values(values)
                if values
                    .iter()
                    .all(|value| matches!(value, ComparableValue::None))
                    || values
                        .iter()
                        .any(|value| matches!(value, ComparableValue::String(_))) =>
            {
                KernelValue::from_comparable_value(
                    ComparableValue::Number(0.0),
                    self.scenario_count,
                )
            }
            value => value,
        };
        self.iterative_values.insert(cell, value.clone());
        value
    }

    fn update_iterative_value(&mut self, cell: CellId, value: &KernelValue) {
        let previous = self.iterative_values.insert(cell, value.clone());
        let delta = previous
            .as_ref()
            .map(|previous| kernel_value_delta(previous, value))
            .unwrap_or(f64::INFINITY);
        self.iteration_max_delta = self.iteration_max_delta.max(delta);
    }
}

#[derive(Clone, Copy)]
enum NumericAggregate {
    Sum,
    Product,
    Min,
    Max,
}

#[derive(Clone, Copy)]
enum CountMode {
    Numbers,
    NonEmpty,
}

#[derive(Clone, Copy)]
enum BooleanAggregate {
    And,
    Or,
}

#[derive(Clone, Copy)]
enum IfsAggregate {
    Sum,
    Average,
    Min,
    Max,
}

#[derive(Clone, Copy)]
enum LookupDirection {
    Vertical,
    Horizontal,
}

#[derive(Clone, Copy)]
enum RoundMode {
    Nearest,
    Up,
    Down,
}

#[derive(Clone, Copy)]
enum CriteriaOperator {
    Equal,
    NotEqual,
    LessThan,
    LessOrEqual,
    GreaterThan,
    GreaterOrEqual,
}

impl KernelValue {
    fn number(value: f64, scenario_count: usize) -> Self {
        Self::from_comparable_value(ComparableValue::Number(value), scenario_count)
    }

    fn boolean(value: bool, scenario_count: usize) -> Self {
        Self::from_comparable_value(ComparableValue::Boolean(value), scenario_count)
    }

    fn string(value: String, scenario_count: usize) -> Self {
        Self::from_comparable_value(ComparableValue::String(value), scenario_count)
    }

    fn empty(scenario_count: usize) -> Self {
        Self::from_comparable_value(ComparableValue::None, scenario_count)
    }

    fn from_comparable_value(value: ComparableValue, scenario_count: usize) -> Self {
        Self::Values(vec![value; scenario_count])
    }

    fn from_comparable_values(values: Vec<ComparableValue>) -> Self {
        Self::Values(values)
    }

    fn into_comparable_values(
        self,
        context: &mut KernelContext<'_, '_>,
    ) -> Result<Vec<ComparableValue>, KernelError> {
        match self {
            KernelValue::Values(values) => Ok(values),
            KernelValue::Range(range) if range.cells.len() == 1 => context
                .eval_cell(range.cells[0])?
                .into_comparable_values(context),
            KernelValue::Range(_) => Err(KernelError::Unsupported),
        }
    }
}

fn comparable_values_for_cells<'a>(
    model: &Model<'_>,
    cells: impl Iterator<Item = &'a SheetCellRef>,
) -> Result<Vec<ComparableValue>, KernelError> {
    cells
        .map(|cell| read_comparable_value(model, cell).ok_or(KernelError::Unsupported))
        .collect()
}

fn comparable_values_for_optional_cells<'a>(
    model: &Model<'_>,
    cells: impl Iterator<Item = Option<&'a SheetCellRef>>,
) -> Result<Vec<Option<ComparableValue>>, KernelError> {
    cells
        .map(|cell| {
            cell.map(|cell| read_comparable_value(model, cell).ok_or(KernelError::Unsupported))
                .transpose()
        })
        .collect()
}

fn collect_optional_values(values: Vec<Option<ComparableValue>>) -> Option<Vec<ComparableValue>> {
    values.into_iter().collect()
}

fn coerce_numbers(
    context: &mut KernelContext<'_, '_>,
    value: KernelValue,
) -> Result<Vec<f64>, KernelError> {
    value
        .into_comparable_values(context)?
        .into_iter()
        .map(|value| match value {
            ComparableValue::None => Ok(0.0),
            ComparableValue::Number(value) => Ok(value),
            ComparableValue::Boolean(value) => Ok(if value { 1.0 } else { 0.0 }),
            ComparableValue::String(value) => {
                value.parse::<f64>().map_err(|_| KernelError::Unsupported)
            }
        })
        .collect()
}

fn coerce_comparable_to_boolean(value: &ComparableValue) -> Result<bool, KernelError> {
    match value {
        ComparableValue::None => Ok(false),
        ComparableValue::Number(value) => Ok(*value != 0.0),
        ComparableValue::Boolean(value) => Ok(*value),
        ComparableValue::String(_) => Err(KernelError::Unsupported),
    }
}

fn coerce_booleans(
    context: &mut KernelContext<'_, '_>,
    value: KernelValue,
) -> Result<Vec<bool>, KernelError> {
    value
        .into_comparable_values(context)?
        .into_iter()
        .map(|value| match value {
            ComparableValue::None => Ok(false),
            ComparableValue::Number(value) => Ok(value != 0.0),
            ComparableValue::Boolean(value) => Ok(value),
            ComparableValue::String(_) => Err(KernelError::Unsupported),
        })
        .collect()
}

fn flatten_numbers(
    context: &mut KernelContext<'_, '_>,
    value: KernelValue,
) -> Result<Vec<Vec<f64>>, KernelError> {
    match value {
        KernelValue::Values(_) => Ok(vec![coerce_numbers(context, value)?]),
        KernelValue::Range(range) => range
            .cells
            .iter()
            .map(|cell| {
                let value = context.eval_cell(*cell)?;
                coerce_numbers(context, value)
            })
            .collect(),
    }
}

fn flatten_comparable_values(
    context: &mut KernelContext<'_, '_>,
    value: KernelValue,
) -> Result<Vec<Vec<ComparableValue>>, KernelError> {
    match value {
        KernelValue::Values(values) => Ok(vec![values]),
        KernelValue::Range(range) => range
            .cells
            .iter()
            .map(|cell| context.eval_cell(*cell)?.into_comparable_values(context))
            .collect(),
    }
}

fn constant_i32(values: Vec<f64>) -> Result<i32, KernelError> {
    let Some(first) = values.first() else {
        return Err(KernelError::Unsupported);
    };
    let first = first.round() as i32;
    if values.iter().all(|value| value.round() as i32 == first) {
        Ok(first)
    } else {
        Err(KernelError::Unsupported)
    }
}

fn comparable_to_text(value: &ComparableValue) -> String {
    match value {
        ComparableValue::None => String::new(),
        ComparableValue::Number(value) => value.to_string(),
        ComparableValue::String(value) => value.clone(),
        ComparableValue::Boolean(value) => {
            if *value {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
    }
}

fn kernel_value_delta(left: &KernelValue, right: &KernelValue) -> f64 {
    match (left, right) {
        (KernelValue::Values(left), KernelValue::Values(right)) if left.len() == right.len() => {
            left.iter()
                .zip(right.iter())
                .map(|(left, right)| comparable_delta(left, right))
                .fold(0.0, f64::max)
        }
        (KernelValue::Range(left), KernelValue::Range(right))
            if left.cells == right.cells
                && left.rows == right.rows
                && left.columns == right.columns =>
        {
            0.0
        }
        _ => f64::INFINITY,
    }
}

fn comparable_delta(left: &ComparableValue, right: &ComparableValue) -> f64 {
    match (left, right) {
        (ComparableValue::Number(left), ComparableValue::Number(right)) => (left - right).abs(),
        (ComparableValue::Boolean(left), ComparableValue::Boolean(right)) if left == right => 0.0,
        (ComparableValue::String(left), ComparableValue::String(right)) if left == right => 0.0,
        (ComparableValue::None, ComparableValue::None) => 0.0,
        _ => f64::INFINITY,
    }
}

fn solve_irr(cash_flows: &[f64], guess: f64) -> Option<f64> {
    if cash_flows.is_empty()
        || !cash_flows.iter().any(|value| *value > 0.0)
        || !cash_flows.iter().any(|value| *value < 0.0)
    {
        return None;
    }

    let mut rate = guess.max(-0.999_999);
    for _ in 0..100 {
        let (value, derivative) = irr_value_and_derivative(cash_flows, rate)?;
        if value.abs() <= 1e-10 {
            return Some(rate);
        }
        if derivative == 0.0 {
            break;
        }
        let next = rate - value / derivative;
        if !next.is_finite() || next <= -0.999_999 {
            break;
        }
        if (next - rate).abs() <= 1e-10 {
            return Some(next);
        }
        rate = next;
    }

    solve_irr_bisection(cash_flows)
}

fn irr_value_and_derivative(cash_flows: &[f64], rate: f64) -> Option<(f64, f64)> {
    if rate <= -1.0 {
        return None;
    }
    let base = 1.0 + rate;
    let mut value = 0.0;
    let mut derivative = 0.0;
    for (period, cash_flow) in cash_flows.iter().enumerate() {
        value += cash_flow / base.powi(period as i32);
        if period > 0 {
            derivative -= period as f64 * cash_flow / base.powi(period as i32 + 1);
        }
    }
    if value.is_finite() && derivative.is_finite() {
        Some((value, derivative))
    } else {
        None
    }
}

fn solve_irr_bisection(cash_flows: &[f64]) -> Option<f64> {
    let mut previous_rate = -0.999_9;
    let mut previous_value = irr_value(cash_flows, previous_rate)?;
    let max_rate = 10.0;
    let steps = 512;
    for step in 1..=steps {
        let rate = -0.999_9 + (max_rate + 0.999_9) * step as f64 / steps as f64;
        let value = irr_value(cash_flows, rate)?;
        if previous_value == 0.0 {
            return Some(previous_rate);
        }
        if value == 0.0 {
            return Some(rate);
        }
        if previous_value.signum() != value.signum() {
            return bisect_irr(cash_flows, previous_rate, rate);
        }
        previous_rate = rate;
        previous_value = value;
    }
    None
}

fn irr_value(cash_flows: &[f64], rate: f64) -> Option<f64> {
    irr_value_and_derivative(cash_flows, rate).map(|(value, _)| value)
}

fn bisect_irr(cash_flows: &[f64], mut low: f64, mut high: f64) -> Option<f64> {
    let mut low_value = irr_value(cash_flows, low)?;
    for _ in 0..100 {
        let mid = (low + high) / 2.0;
        let mid_value = irr_value(cash_flows, mid)?;
        if mid_value.abs() <= 1e-10 || (high - low).abs() <= 1e-10 {
            return Some(mid);
        }
        if low_value.signum() == mid_value.signum() {
            low = mid;
            low_value = mid_value;
        } else {
            high = mid;
        }
    }
    Some((low + high) / 2.0)
}

fn approximate_match_position(
    range_values: &[Vec<ComparableValue>],
    lookup: &ComparableValue,
    scenario_index: usize,
    ascending: bool,
) -> Option<usize> {
    let ComparableValue::Number(lookup) = lookup else {
        return None;
    };
    let mut best = None;
    for (index, candidate_values) in range_values.iter().enumerate() {
        let ComparableValue::Number(candidate) = candidate_values.get(scenario_index)? else {
            return None;
        };
        if ascending {
            if candidate <= lookup {
                best = Some(index);
            }
        } else if candidate >= lookup {
            best = Some(index);
        }
    }
    best
}

fn lookup_position(
    table_values: &[Vec<ComparableValue>],
    positions: impl Iterator<Item = usize>,
    lookup: &ComparableValue,
    scenario_index: usize,
    approximate: bool,
) -> Option<usize> {
    let mut best = None;
    for (position_index, cell_index) in positions.enumerate() {
        let candidate = table_values.get(cell_index)?.get(scenario_index)?;
        if lookup_values_match(candidate, lookup) {
            return Some(position_index);
        }
        if approximate && compare_order(candidate, lookup).is_some_and(|order| order.is_le()) {
            best = Some(position_index);
        }
    }
    if approximate {
        best
    } else {
        None
    }
}

fn lookup_values_match(candidate: &ComparableValue, lookup: &ComparableValue) -> bool {
    match (candidate, lookup) {
        (ComparableValue::String(candidate), ComparableValue::String(lookup)) => {
            if contains_wildcard(lookup) {
                wildcard_matches(lookup, candidate)
            } else {
                candidate.eq_ignore_ascii_case(lookup)
            }
        }
        _ => values_match(candidate, lookup),
    }
}

fn criteria_sets_match(
    criteria_sets: &[CriteriaSet],
    offset: usize,
    scenario_index: usize,
) -> Result<bool, KernelError> {
    for criteria_set in criteria_sets {
        let Some(values) = criteria_set.values_by_cell.get(offset) else {
            return Err(KernelError::Unsupported);
        };
        let Some(value) = values.get(scenario_index) else {
            return Err(KernelError::Unsupported);
        };
        let Some(criteria) = criteria_set.criteria.get(scenario_index) else {
            return Err(KernelError::Unsupported);
        };
        if !criteria_matches(value, criteria)? {
            return Ok(false);
        }
    }
    Ok(true)
}

fn criteria_matches(
    value: &ComparableValue,
    criteria: &ComparableValue,
) -> Result<bool, KernelError> {
    match criteria {
        ComparableValue::String(criteria) => {
            let (operator, operand) = parse_criteria(criteria);
            if operand.is_empty() {
                return Ok(compare_criteria_values(
                    value,
                    &ComparableValue::None,
                    operator,
                ));
            }
            if matches!(
                operator,
                CriteriaOperator::Equal | CriteriaOperator::NotEqual
            ) && contains_wildcard(operand)
            {
                let matched = wildcard_matches(operand, &comparable_to_text(value));
                return Ok(match operator {
                    CriteriaOperator::Equal => matched,
                    CriteriaOperator::NotEqual => !matched,
                    _ => unreachable!(),
                });
            }
            Ok(compare_criteria_values(
                value,
                &parse_criteria_operand(operand),
                operator,
            ))
        }
        _ => Ok(compare_criteria_values(
            value,
            criteria,
            CriteriaOperator::Equal,
        )),
    }
}

fn parse_criteria(criteria: &str) -> (CriteriaOperator, &str) {
    let criteria = criteria.trim();
    for (prefix, operator) in [
        (">=", CriteriaOperator::GreaterOrEqual),
        ("<=", CriteriaOperator::LessOrEqual),
        ("<>", CriteriaOperator::NotEqual),
        (">", CriteriaOperator::GreaterThan),
        ("<", CriteriaOperator::LessThan),
        ("=", CriteriaOperator::Equal),
    ] {
        if let Some(operand) = criteria.strip_prefix(prefix) {
            return (operator, operand.trim());
        }
    }
    (CriteriaOperator::Equal, criteria)
}

fn parse_criteria_operand(operand: &str) -> ComparableValue {
    if operand.eq_ignore_ascii_case("true") {
        ComparableValue::Boolean(true)
    } else if operand.eq_ignore_ascii_case("false") {
        ComparableValue::Boolean(false)
    } else if let Ok(value) = operand.parse::<f64>() {
        ComparableValue::Number(value)
    } else {
        ComparableValue::String(operand.to_string())
    }
}

fn compare_criteria_values(
    value: &ComparableValue,
    criteria: &ComparableValue,
    operator: CriteriaOperator,
) -> bool {
    match operator {
        CriteriaOperator::Equal => criteria_equal(value, criteria),
        CriteriaOperator::NotEqual => !criteria_equal(value, criteria),
        CriteriaOperator::LessThan => {
            compare_order(value, criteria).is_some_and(|order| order.is_lt())
        }
        CriteriaOperator::LessOrEqual => {
            compare_order(value, criteria).is_some_and(|order| order.is_le())
        }
        CriteriaOperator::GreaterThan => {
            compare_order(value, criteria).is_some_and(|order| order.is_gt())
        }
        CriteriaOperator::GreaterOrEqual => {
            compare_order(value, criteria).is_some_and(|order| order.is_ge())
        }
    }
}

fn criteria_equal(left: &ComparableValue, right: &ComparableValue) -> bool {
    match (left, right) {
        (ComparableValue::String(left), ComparableValue::String(right)) => {
            left.eq_ignore_ascii_case(right)
        }
        _ => values_match(left, right),
    }
}

fn contains_wildcard(pattern: &str) -> bool {
    pattern.chars().any(|ch| matches!(ch, '*' | '?'))
}

fn wildcard_matches(pattern: &str, value: &str) -> bool {
    let pattern = pattern.to_ascii_lowercase();
    let value = value.to_ascii_lowercase();
    wildcard_matches_bytes(pattern.as_bytes(), value.as_bytes())
}

fn wildcard_matches_bytes(pattern: &[u8], value: &[u8]) -> bool {
    let mut pattern_index = 0;
    let mut value_index = 0;
    let mut star_index = None;
    let mut match_index = 0;

    while value_index < value.len() {
        if pattern_index < pattern.len()
            && (pattern[pattern_index] == b'?' || pattern[pattern_index] == value[value_index])
        {
            pattern_index += 1;
            value_index += 1;
        } else if pattern_index < pattern.len() && pattern[pattern_index] == b'*' {
            star_index = Some(pattern_index);
            match_index = value_index;
            pattern_index += 1;
        } else if let Some(star) = star_index {
            pattern_index = star + 1;
            match_index += 1;
            value_index = match_index;
        } else {
            return false;
        }
    }

    while pattern_index < pattern.len() && pattern[pattern_index] == b'*' {
        pattern_index += 1;
    }
    pattern_index == pattern.len()
}

fn compare_values(left: &ComparableValue, right: &ComparableValue, kind: &OpCompare) -> bool {
    match kind {
        OpCompare::Equal => values_match(left, right),
        OpCompare::NonEqual => !values_match(left, right),
        OpCompare::LessThan => compare_order(left, right).is_some_and(|order| order.is_lt()),
        OpCompare::GreaterThan => compare_order(left, right).is_some_and(|order| order.is_gt()),
        OpCompare::LessOrEqualThan => compare_order(left, right).is_some_and(|order| order.is_le()),
        OpCompare::GreaterOrEqualThan => {
            compare_order(left, right).is_some_and(|order| order.is_ge())
        }
    }
}

fn compare_order(left: &ComparableValue, right: &ComparableValue) -> Option<std::cmp::Ordering> {
    match (left, right) {
        (ComparableValue::Number(left), ComparableValue::Number(right)) => left.partial_cmp(right),
        (ComparableValue::String(left), ComparableValue::String(right)) => Some(left.cmp(right)),
        (ComparableValue::Boolean(left), ComparableValue::Boolean(right)) => Some(left.cmp(right)),
        _ => None,
    }
}

fn cells_for_rect(
    sheet: u32,
    start_row: i32,
    start_column: i32,
    end_row: i32,
    end_column: i32,
) -> KernelRange {
    let top = start_row.min(end_row);
    let bottom = start_row.max(end_row);
    let left = start_column.min(end_column);
    let right = start_column.max(end_column);
    let rows = (bottom - top + 1).max(0) as usize;
    let columns = (right - left + 1).max(0) as usize;
    let mut cells = Vec::with_capacity(rows.saturating_mul(columns));
    for row in top..=bottom {
        for column in left..=right {
            cells.push(CellId::new(sheet, row, column));
        }
    }
    KernelRange {
        cells,
        rows,
        columns,
    }
}

fn absolute_coord(value: i32, context: i32, is_absolute: bool) -> i32 {
    if is_absolute {
        value
    } else {
        value + context
    }
}

impl ScenarioCounts {
    fn add(&mut self, other: ScenarioCounts) {
        self.evaluated += other.evaluated;
        self.validated += other.validated;
        self.mismatched += other.mismatched;
        self.unsupported += other.unsupported;
        self.diagnostics.extend(other.diagnostics);
    }
}

fn data_table_status(summary: &DataTableBenchmarkSummary) -> DataTableEvaluationStatus {
    if summary.data_table_count == 0 {
        DataTableEvaluationStatus::None
    } else if summary.evaluated_data_table_cells == 0 {
        DataTableEvaluationStatus::Unsupported
    } else if summary.mismatched_data_table_cells > 0 {
        DataTableEvaluationStatus::Mismatch
    } else if summary.unsupported_data_table_cells > 0 {
        DataTableEvaluationStatus::Partial
    } else {
        DataTableEvaluationStatus::Validated
    }
}

fn dirty_table_indexes(
    model: &Model<'_>,
    data_tables: &[DataTableRegion],
    changed_cells: &[CellId],
) -> HashSet<usize> {
    if changed_cells.is_empty() {
        return HashSet::new();
    }

    let mut dirty = HashSet::new();
    for (index, table) in data_tables.iter().enumerate() {
        for changed_cell in changed_cells {
            if data_table_contains_cell(model, table, *changed_cell) {
                dirty.insert(index);
                break;
            }
        }
    }
    dirty
}

fn data_table_contains_cell(model: &Model<'_>, table: &DataTableRegion, cell: CellId) -> bool {
    let Some(sheet_index) = sheet_index_by_name(model, &table.sheet_name) else {
        return false;
    };
    if cell.sheet == sheet_index {
        if table.range.contains(cell.row, cell.column) {
            return true;
        }
        if cell.row == table.range.start_row - 1
            && table.range.start_column <= cell.column
            && cell.column <= table.range.end_column
        {
            return true;
        }
        if cell.column == table.range.start_column - 1
            && table.range.start_row <= cell.row
            && cell.row <= table.range.end_row
        {
            return true;
        }
    }

    [
        table.formula_cell.as_ref(),
        table.column_input_cell.as_ref(),
        table.row_input_cell.as_ref(),
    ]
    .iter()
    .flatten()
    .filter_map(|cell_ref| resolve_cell(model, cell_ref))
    .any(|resolved| resolved == cell)
}

fn resolve_cell(model: &Model<'_>, cell_ref: &SheetCellRef) -> Option<CellId> {
    Some(CellId::new(
        sheet_index_by_name(model, &cell_ref.sheet_name)?,
        cell_ref.row,
        cell_ref.column,
    ))
}

fn sheet_index_by_name(model: &Model<'_>, name: &str) -> Option<u32> {
    model
        .workbook
        .worksheets
        .iter()
        .position(|worksheet| worksheet.name.eq_ignore_ascii_case(name))
        .map(|index| index as u32)
}

fn read_comparable_value(model: &Model<'_>, cell_ref: &SheetCellRef) -> Option<ComparableValue> {
    let resolved = resolve_cell(model, cell_ref)?;
    model
        .get_cell_value_by_index(resolved.sheet, resolved.row, resolved.column)
        .ok()
        .map(ComparableValue::from)
}

fn values_match(actual: &ComparableValue, expected: &ComparableValue) -> bool {
    match (actual, expected) {
        (ComparableValue::None, ComparableValue::None) => true,
        (ComparableValue::Boolean(left), ComparableValue::Boolean(right)) => left == right,
        (ComparableValue::String(left), ComparableValue::String(right)) => left == right,
        (ComparableValue::Number(left), ComparableValue::Number(right)) => {
            (left - right).abs() <= NUMERIC_TOLERANCE * right.abs().max(1.0)
        }
        _ => false,
    }
}

impl From<CellValue> for ComparableValue {
    fn from(value: CellValue) -> Self {
        match value {
            CellValue::None => ComparableValue::None,
            CellValue::Number(value) => ComparableValue::Number(value),
            CellValue::String(value) => ComparableValue::String(value),
            CellValue::Boolean(value) => ComparableValue::Boolean(value),
        }
    }
}

fn parse_a1_cell(reference: &str, default_sheet: &str) -> Option<(String, i32, i32)> {
    let cleaned = reference.trim().replace('$', "");
    let (sheet_name, cell_address) = match cleaned.rsplit_once('!') {
        Some((sheet, cell)) => (unquote_sheet_name(sheet), cell),
        None => (default_sheet.to_string(), cleaned.as_str()),
    };

    let mut column_name = String::new();
    let mut row_name = String::new();
    for ch in cell_address.chars() {
        if ch.is_ascii_alphabetic() && row_name.is_empty() {
            column_name.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            row_name.push(ch);
        } else {
            return None;
        }
    }

    if column_name.is_empty() || row_name.is_empty() {
        return None;
    }

    Some((
        sheet_name,
        row_name.parse().ok()?,
        column_name_to_number(&column_name)?,
    ))
}

fn column_name_to_number(column_name: &str) -> Option<i32> {
    let mut column = 0_i32;
    for ch in column_name.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        column = column
            .checked_mul(26)?
            .checked_add(ch.to_ascii_uppercase() as i32 - 'A' as i32 + 1)?;
    }
    Some(column)
}

fn unquote_sheet_name(sheet_name: &str) -> String {
    sheet_name
        .trim()
        .strip_prefix('\'')
        .and_then(|value| value.strip_suffix('\''))
        .unwrap_or(sheet_name.trim())
        .replace("''", "'")
}

#[cfg(test)]
mod tests {
    use ironcalc::base::Model;

    use super::*;

    #[test]
    fn parses_a1_ranges_and_sheet_qualified_cells() {
        let range = parse_sheet_range("$E$121:$I$125").unwrap();
        assert_eq!(range.start_row, 121);
        assert_eq!(range.start_column, 5);
        assert_eq!(range.end_row, 125);
        assert_eq!(range.end_column, 9);

        let cell = parse_sheet_cell_ref("'Target DCF'!$G$75", "Sheet1").unwrap();
        assert_eq!(cell.sheet_name, "Target DCF");
        assert_eq!(cell.address, "G75");
        assert_eq!(cell.row, 75);
        assert_eq!(cell.column, 7);
    }

    #[test]
    fn builds_conventional_two_dimensional_region() {
        let table = build_data_table_region(
            "Target DCF",
            "xl/worksheets/sheet12.xml",
            "E121",
            "E121:I125",
            Some("G75"),
            Some("I69"),
            true,
            true,
        );

        assert_eq!(
            table
                .formula_cell
                .as_ref()
                .map(|cell| cell.address.as_str()),
            Some("D120")
        );
        assert!(table.is_parallel_validation_eligible());
        assert_eq!(table.cell_count(), 25);
    }

    #[test]
    fn validates_conventional_two_dimensional_table() {
        let mut model = Model::new_empty("table", "en", "UTC", "en").unwrap();
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

        let table = build_data_table_region(
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            "C3",
            "C3:D4",
            Some("A1"),
            Some("A2"),
            true,
            true,
        );
        let summary = evaluate_data_tables(&model, &[table], &[], true);

        assert_eq!(summary.status, DataTableEvaluationStatus::Validated);
        assert_eq!(summary.validated_data_table_cells, 4);
        assert_eq!(summary.mismatched_data_table_cells, 0);
    }

    #[test]
    fn validates_vectorized_formula_cone_with_common_banker_functions() {
        let mut model = Model::new_empty("table", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "1".to_string()).unwrap();
        model.set_user_input(0, 2, 1, "10".to_string()).unwrap();
        model.set_user_input(0, 1, 3, "100".to_string()).unwrap();
        model.set_user_input(0, 1, 4, "200".to_string()).unwrap();
        model.set_user_input(0, 1, 5, "300".to_string()).unwrap();
        model
            .set_user_input(
                0,
                2,
                2,
                "=IF(A1>2,SUM(A1:A2),INDEX($C$1:$E$1,1,MATCH(A1,$C$2:$E$2,0))+A2)".to_string(),
            )
            .unwrap();
        model.set_user_input(0, 2, 3, "1".to_string()).unwrap();
        model.set_user_input(0, 2, 4, "2".to_string()).unwrap();
        model.set_user_input(0, 2, 5, "3".to_string()).unwrap();
        model.set_user_input(0, 3, 2, "10".to_string()).unwrap();
        model.set_user_input(0, 4, 2, "20".to_string()).unwrap();
        model.set_user_input(0, 3, 3, "110".to_string()).unwrap();
        model.set_user_input(0, 3, 4, "210".to_string()).unwrap();
        model.set_user_input(0, 3, 5, "13".to_string()).unwrap();
        model.set_user_input(0, 4, 3, "120".to_string()).unwrap();
        model.set_user_input(0, 4, 4, "220".to_string()).unwrap();
        model.set_user_input(0, 4, 5, "23".to_string()).unwrap();
        model.evaluate();

        let table = build_data_table_region(
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            "C3",
            "C3:E4",
            Some("A1"),
            Some("A2"),
            true,
            true,
        );
        let summary = evaluate_data_tables(&model, &[table], &[], true);

        assert_eq!(summary.status, DataTableEvaluationStatus::Validated);
        assert_eq!(summary.evaluated_data_table_cells, 6);
        assert_eq!(summary.validated_data_table_cells, 6);
        assert_eq!(summary.mismatched_data_table_cells, 0);
    }

    #[test]
    fn validates_one_variable_row_data_table() {
        let mut model = Model::new_empty("one-row", "en", "UTC", "en").unwrap();
        set(&mut model, 1, 1, "0");
        set(&mut model, 3, 2, "=A1*2");
        set(&mut model, 2, 3, "10");
        set(&mut model, 2, 4, "20");
        set(&mut model, 2, 5, "30");
        set(&mut model, 3, 3, "20");
        set(&mut model, 3, 4, "40");
        set(&mut model, 3, 5, "60");

        let table = build_data_table_region(
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            "C3",
            "C3:E3",
            Some("A1"),
            None,
            false,
            true,
        );
        let summary = evaluate_data_tables(&model, &[table], &[], true);

        assert_eq!(summary.status, DataTableEvaluationStatus::Validated);
        assert_eq!(summary.validated_data_table_cells, 3);
        assert_eq!(summary.mismatched_data_table_cells, 0);
    }

    #[test]
    fn validates_one_variable_column_data_table() {
        let mut model = Model::new_empty("one-column", "en", "UTC", "en").unwrap();
        set(&mut model, 1, 1, "0");
        set(&mut model, 2, 2, "=A1*3");
        set(&mut model, 3, 1, "10");
        set(&mut model, 4, 1, "20");
        set(&mut model, 5, 1, "30");
        set(&mut model, 3, 2, "30");
        set(&mut model, 4, 2, "60");
        set(&mut model, 5, 2, "90");

        let table = build_data_table_region(
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            "B3",
            "B3:B5",
            Some("A1"),
            None,
            false,
            false,
        );
        let summary = evaluate_data_tables(&model, &[table], &[], true);

        assert_eq!(summary.status, DataTableEvaluationStatus::Validated);
        assert_eq!(summary.validated_data_table_cells, 3);
        assert_eq!(summary.mismatched_data_table_cells, 0);
    }

    #[test]
    fn validates_choose_scalar_and_range_branches() {
        let mut model = Model::new_empty("choose", "en", "UTC", "en").unwrap();
        set(&mut model, 1, 1, "2");
        set(&mut model, 2, 2, "=CHOOSE(A1,100,200,300)+A2");
        set(&mut model, 2, 3, "1");
        set(&mut model, 2, 4, "2");
        set(&mut model, 2, 5, "3");
        set(&mut model, 3, 2, "10");
        set(&mut model, 4, 2, "20");
        set(&mut model, 3, 3, "110");
        set(&mut model, 3, 4, "210");
        set(&mut model, 3, 5, "310");
        set(&mut model, 4, 3, "120");
        set(&mut model, 4, 4, "220");
        set(&mut model, 4, 5, "320");

        set(&mut model, 1, 6, "10");
        set(&mut model, 1, 7, "20");
        set(&mut model, 1, 8, "30");
        set(&mut model, 2, 6, "40");
        set(&mut model, 2, 7, "50");
        set(&mut model, 2, 8, "60");
        set(
            &mut model,
            7,
            2,
            "=SUM(CHOOSE($A$1,$F$1:$H$1,$F$2:$H$2))+A2",
        );
        set(&mut model, 6, 3, "1");
        set(&mut model, 6, 4, "2");
        set(&mut model, 7, 3, "151");
        set(&mut model, 7, 4, "152");

        let scalar_table = build_data_table_region(
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            "C3",
            "C3:E4",
            Some("A1"),
            Some("A2"),
            true,
            true,
        );
        let range_table = build_data_table_region(
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            "C7",
            "C7:D7",
            Some("A2"),
            None,
            false,
            true,
        );
        let summary = evaluate_data_tables(&model, &[scalar_table, range_table], &[], true);

        assert_eq!(summary.status, DataTableEvaluationStatus::Validated);
        assert_eq!(summary.validated_data_table_cells, 8);
        assert_eq!(summary.mismatched_data_table_cells, 0);
    }

    #[test]
    fn validates_iterative_formula_cone() {
        let mut model = Model::new_empty("iterative", "en", "UTC", "en").unwrap();
        set(&mut model, 1, 1, "0");
        set(&mut model, 1, 3, "=(C1+A1)/2");
        set(&mut model, 3, 2, "=C1");
        set(&mut model, 2, 3, "10");
        set(&mut model, 2, 4, "20");
        set(&mut model, 3, 3, "10");
        set(&mut model, 3, 4, "20");

        let table = build_data_table_region(
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            "C3",
            "C3:D3",
            Some("A1"),
            None,
            false,
            true,
        );
        let summary = evaluate_data_tables(&model, &[table], &[], true);

        assert_eq!(summary.status, DataTableEvaluationStatus::Validated);
        assert_eq!(summary.validated_data_table_cells, 2);
        assert_eq!(summary.mismatched_data_table_cells, 0);
    }

    #[test]
    fn validates_finance_criteria_lookup_and_text_functions() {
        let mut model = Model::new_empty("ma-functions", "en", "UTC", "en").unwrap();
        set(&mut model, 1, 6, "Base");
        set(&mut model, 2, 6, "Upside");
        set(&mut model, 3, 6, "Base");
        set(&mut model, 4, 6, "Downside");
        set(&mut model, 1, 7, "10");
        set(&mut model, 2, 7, "20");
        set(&mut model, 3, 7, "30");
        set(&mut model, 4, 7, "40");
        set(&mut model, 1, 8, "1");
        set(&mut model, 2, 8, "2");
        set(&mut model, 3, 8, "3");
        set(&mut model, 1, 9, "100");
        set(&mut model, 2, 9, "200");
        set(&mut model, 3, 9, "300");
        set(&mut model, 6, 8, "1");
        set(&mut model, 6, 9, "2");
        set(&mut model, 6, 10, "3");
        set(&mut model, 7, 8, "10");
        set(&mut model, 7, 9, "20");
        set(&mut model, 7, 10, "30");
        set(&mut model, 1, 11, "11");
        set(&mut model, 1, 12, "12");
        set(&mut model, 1, 13, "13");
        set(&mut model, 2, 11, "-100");
        set(&mut model, 2, 12, "60");
        set(&mut model, 2, 13, "60");

        set(
            &mut model,
            2,
            2,
            concat!(
                "=SUMIF($F$1:$F$4,\"Base\",$G$1:$G$4)",
                "+COUNTIF($G$1:$G$4,\">=20\")",
                "+AVERAGEIF($F$1:$F$4,\"<>Downside\",$G$1:$G$4)",
                "+MINIFS($G$1:$G$4,$F$1:$F$4,\"Base\")",
                "+MAXIFS($G$1:$G$4,$F$1:$F$4,\"Base\")",
                "+VLOOKUP(A1,$H$1:$I$3,2,0)",
                "+HLOOKUP(A1,$H$6:$J$7,2,0)",
                "+NPV(0.1,$K$1:$M$1)",
                "+IRR($K$2:$M$2)",
                "+PMT(0.1,2,-100)",
                "+A2"
            ),
        );
        set(&mut model, 2, 3, "1");
        set(&mut model, 2, 4, "2");
        set(&mut model, 3, 2, "5");
        set(&mut model, 4, 2, "15");

        let npv = 11.0 / 1.1_f64 + 12.0 / 1.1_f64.powi(2) + 13.0 / 1.1_f64.powi(3);
        let irr = (60.0 + 27_600.0_f64.sqrt()) / 200.0 - 1.0;
        let pmt = 100.0 * 1.1_f64.powi(2) * 0.1 / (1.1_f64.powi(2) - 1.0);
        let base = 40.0 + 3.0 + 20.0 + 10.0 + 30.0 + npv + irr + pmt;
        set(&mut model, 3, 3, &(base + 100.0 + 10.0 + 5.0).to_string());
        set(&mut model, 3, 4, &(base + 200.0 + 20.0 + 5.0).to_string());
        set(&mut model, 4, 3, &(base + 100.0 + 10.0 + 15.0).to_string());
        set(&mut model, 4, 4, &(base + 200.0 + 20.0 + 15.0).to_string());

        set(&mut model, 8, 2, "=CONCATENATE(\"Premium \",A1)&\"x\"");
        set(&mut model, 7, 3, "10");
        set(&mut model, 7, 4, "20");
        set(&mut model, 8, 3, "Premium 10x");
        set(&mut model, 8, 4, "Premium 20x");

        let numeric_table = build_data_table_region(
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            "C3",
            "C3:D4",
            Some("A1"),
            Some("A2"),
            true,
            true,
        );
        let text_table = build_data_table_region(
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            "C8",
            "C8:D8",
            Some("A1"),
            None,
            false,
            true,
        );
        let summary = evaluate_data_tables(&model, &[numeric_table, text_table], &[], true);

        assert_eq!(summary.status, DataTableEvaluationStatus::Validated);
        assert_eq!(summary.validated_data_table_cells, 6);
        assert_eq!(summary.mismatched_data_table_cells, 0);
    }

    #[test]
    fn keeps_unsupported_formula_cones_as_fallbacks() {
        let mut model = Model::new_empty("table", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "1".to_string()).unwrap();
        model.set_user_input(0, 2, 1, "10".to_string()).unwrap();
        model
            .set_user_input(0, 2, 2, "=INDIRECT(\"A1\")*A2".to_string())
            .unwrap();
        model.set_user_input(0, 2, 3, "1".to_string()).unwrap();
        model.set_user_input(0, 2, 4, "2".to_string()).unwrap();
        model.set_user_input(0, 3, 2, "10".to_string()).unwrap();
        model.set_user_input(0, 4, 2, "20".to_string()).unwrap();
        model.set_user_input(0, 3, 3, "10".to_string()).unwrap();
        model.set_user_input(0, 3, 4, "20".to_string()).unwrap();
        model.set_user_input(0, 4, 3, "20".to_string()).unwrap();
        model.set_user_input(0, 4, 4, "40".to_string()).unwrap();
        model.evaluate();

        let table = build_data_table_region(
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            "C3",
            "C3:D4",
            Some("A1"),
            Some("A2"),
            true,
            true,
        );
        let summary = evaluate_data_tables(&model, &[table], &[], true);

        assert_eq!(summary.status, DataTableEvaluationStatus::Unsupported);
        assert_eq!(summary.evaluated_data_table_cells, 0);
        assert_eq!(summary.unsupported_data_table_cells, 4);
    }

    fn set(model: &mut Model, row: i32, column: i32, input: &str) {
        model
            .set_user_input(0, row, column, input.to_string())
            .unwrap();
    }
}
