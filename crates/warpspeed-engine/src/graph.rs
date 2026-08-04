use std::collections::{HashMap, HashSet, VecDeque};

use ironcalc::base::{
    expressions::{
        parser::Node, token::OpUnary, types::CellReferenceIndex, utils::number_to_column,
    },
    Model,
};
use sha2::{Digest, Sha256};

use crate::model::{FallbackDetail, FallbackReason, FormulaCoverage};

#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash, Ord, PartialOrd)]
pub(crate) struct CellId {
    pub sheet: u32,
    pub row: i32,
    pub column: i32,
}

impl CellId {
    pub(crate) fn new(sheet: u32, row: i32, column: i32) -> Self {
        Self { sheet, row, column }
    }
}

impl From<CellReferenceIndex> for CellId {
    fn from(value: CellReferenceIndex) -> Self {
        Self {
            sheet: value.sheet,
            row: value.row,
            column: value.column,
        }
    }
}

#[derive(Clone, Debug)]
struct RangeRef {
    sheet: u32,
    start_row: i32,
    start_column: i32,
    end_row: i32,
    end_column: i32,
}

impl RangeRef {
    fn contains(&self, cell: CellId) -> bool {
        self.sheet == cell.sheet
            && self.start_row <= cell.row
            && cell.row <= self.end_row
            && self.start_column <= cell.column
            && cell.column <= self.end_column
    }
}

#[derive(Clone, Debug)]
struct RangeDependency {
    range: RangeRef,
    dependent: usize,
}

#[derive(Clone, Debug)]
pub(crate) struct DependencyGraph {
    formula_cells: Vec<CellId>,
    formula_node_by_cell: HashMap<CellId, usize>,
    dependents_by_cell: HashMap<CellId, Vec<usize>>,
    range_dependents_by_sheet: HashMap<u32, Vec<RangeDependency>>,
    sccs: Vec<Vec<usize>>,
    scc_by_node: Vec<usize>,
    formula_has_fallback: Vec<bool>,
    formula_fallback_codes: Vec<HashSet<String>>,
    fallback_reasons: Vec<FallbackReason>,
    fallback_details: Vec<FallbackDetail>,
    circular_sccs: HashSet<usize>,
    has_cached_circular_allowances: bool,
    structure_hash: String,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(crate) struct DirtySummary {
    pub dirty_formula_cells: usize,
    pub planned_reusable_formula_cells: usize,
}

#[derive(Default)]
struct FormulaDependencies {
    cells: Vec<CellId>,
    ranges: Vec<RangeRef>,
    fallbacks: Vec<FormulaFallback>,
}

#[derive(Clone, Debug)]
struct FormulaFallback {
    code: &'static str,
    message: String,
}

pub(crate) fn build_dependency_graph(model: &Model<'_>) -> DependencyGraph {
    let mut formula_entries = Vec::new();
    let mut formula_node_by_cell = HashMap::new();

    for cell_ref in model.get_all_cells() {
        let Some(formula_index) =
            formula_index_for_cell(model, cell_ref.index, cell_ref.row, cell_ref.column)
        else {
            continue;
        };
        let cell = CellId::new(cell_ref.index, cell_ref.row, cell_ref.column);
        let node_index = formula_entries.len();
        formula_entries.push((cell, formula_index));
        formula_node_by_cell.insert(cell, node_index);
    }

    let mut dependents_by_cell: HashMap<CellId, Vec<usize>> = HashMap::new();
    let mut range_dependents_by_sheet: HashMap<u32, Vec<RangeDependency>> = HashMap::new();
    let mut adjacency = vec![Vec::new(); formula_entries.len()];
    let mut formula_has_fallback = vec![false; formula_entries.len()];
    let mut formula_fallback_codes = vec![HashSet::new(); formula_entries.len()];
    let mut fallback_reasons = Vec::new();
    let mut fallback_details = Vec::new();
    let mut fallback_keys = HashSet::new();

    for (dependent_node, (cell, formula_index)) in formula_entries.iter().enumerate() {
        let Some(node) = model
            .parsed_formulas
            .get(cell.sheet as usize)
            .and_then(|formulas| formulas.get(*formula_index as usize))
        else {
            formula_has_fallback[dependent_node] = true;
            formula_fallback_codes[dependent_node].insert("missing_parsed_formula".to_string());
            let location = Some(format_cell_location(model, *cell));
            add_fallback(
                &mut fallback_reasons,
                &mut fallback_keys,
                &mut fallback_details,
                "missing_parsed_formula",
                "Formula cell does not have a parsed IronCalc node.".to_string(),
                location,
                formula_text_for_cell(model, *cell),
                None,
                None,
            );
            continue;
        };

        let mut dependencies = FormulaDependencies::default();
        collect_node_dependencies(model, node, *cell, &mut dependencies);

        if !dependencies.fallbacks.is_empty() {
            formula_has_fallback[dependent_node] = true;
            for fallback in dependencies.fallbacks {
                formula_fallback_codes[dependent_node].insert(fallback.code.to_string());
                let location = Some(format_cell_location(model, *cell));
                add_fallback(
                    &mut fallback_reasons,
                    &mut fallback_keys,
                    &mut fallback_details,
                    fallback.code,
                    fallback.message,
                    location,
                    formula_text_for_cell(model, *cell),
                    None,
                    None,
                );
            }
        }

        for precedent in dependencies.cells {
            dependents_by_cell
                .entry(precedent)
                .or_default()
                .push(dependent_node);
            if let Some(precedent_node) = formula_node_by_cell.get(&precedent) {
                adjacency[*precedent_node].push(dependent_node);
            }
        }

        for range in dependencies.ranges {
            for (precedent_node, (precedent_cell, _)) in formula_entries.iter().enumerate() {
                if range.contains(*precedent_cell) {
                    adjacency[precedent_node].push(dependent_node);
                }
            }
            range_dependents_by_sheet
                .entry(range.sheet)
                .or_default()
                .push(RangeDependency {
                    range,
                    dependent: dependent_node,
                });
        }
    }

    for edges in &mut adjacency {
        edges.sort_unstable();
        edges.dedup();
    }

    let sccs = tarjan_scc(&adjacency);
    let mut scc_by_node = vec![0; formula_entries.len()];
    for (scc_index, scc) in sccs.iter().enumerate() {
        for node in scc {
            scc_by_node[*node] = scc_index;
        }
    }

    let mut circular_sccs = HashSet::new();
    for (scc_index, scc) in sccs.iter().enumerate() {
        let is_cycle = scc.len() > 1 || scc.iter().any(|node| adjacency[*node].contains(node));
        if is_cycle {
            circular_sccs.insert(scc_index);
            for node in scc {
                formula_has_fallback[*node] = true;
                formula_fallback_codes[*node].insert("circular_reference".to_string());
                let cell = formula_entries[*node].0;
                let location = Some(format_cell_location(model, cell));
                add_fallback(
                    &mut fallback_reasons,
                    &mut fallback_keys,
                    &mut fallback_details,
                    "circular_reference",
                    "Formula participates in a circular dependency component.".to_string(),
                    location,
                    formula_text_for_cell(model, cell),
                    Some(scc_index),
                    Some(scc.len()),
                );
            }
        }
    }

    let structure_hash = structure_hash(model, &formula_entries);
    let formula_cells = formula_entries.into_iter().map(|(cell, _)| cell).collect();

    DependencyGraph {
        formula_cells,
        formula_node_by_cell,
        dependents_by_cell,
        range_dependents_by_sheet,
        sccs,
        scc_by_node,
        formula_has_fallback,
        formula_fallback_codes,
        fallback_reasons,
        fallback_details,
        circular_sccs,
        has_cached_circular_allowances: false,
        structure_hash,
    }
}

impl DependencyGraph {
    pub(crate) fn formula_count(&self) -> usize {
        self.formula_cells.len()
    }

    pub(crate) fn structure_hash(&self) -> &str {
        &self.structure_hash
    }

    pub(crate) fn is_result_cache_safe(&self) -> bool {
        self.fallback_reasons.is_empty()
    }

    pub(crate) fn has_cached_circular_allowances(&self) -> bool {
        self.has_cached_circular_allowances
    }

    pub(crate) fn fallback_reasons(&self) -> Vec<FallbackReason> {
        self.fallback_reasons.clone()
    }

    pub(crate) fn fallback_details(&self) -> Vec<FallbackDetail> {
        self.fallback_details.clone()
    }

    pub(crate) fn coverage(&self) -> FormulaCoverage {
        let fallback_formula_cells = self
            .formula_has_fallback
            .iter()
            .filter(|has_fallback| **has_fallback)
            .count();
        FormulaCoverage {
            formula_cells: self.formula_count(),
            supported_formula_cells: self.formula_count().saturating_sub(fallback_formula_cells),
            fallback_formula_cells,
        }
    }

    pub(crate) fn supported_formula_cells(&self) -> impl Iterator<Item = CellId> + '_ {
        self.formula_cells
            .iter()
            .copied()
            .enumerate()
            .filter_map(|(index, cell)| (!self.formula_has_fallback[index]).then_some(cell))
    }

    pub(crate) fn circular_formula_cells(&self) -> impl Iterator<Item = CellId> + '_ {
        self.circular_sccs
            .iter()
            .flat_map(|scc_index| self.sccs[*scc_index].iter())
            .map(|node| self.formula_cells[*node])
    }

    pub(crate) fn allow_cached_circular_values(
        &mut self,
        cached_value_cells: &HashSet<CellId>,
        changed_cells: &[CellId],
    ) {
        if self.circular_sccs.is_empty() {
            return;
        }

        let dirty_nodes = self.dirty_nodes(changed_cells);
        let allowed_sccs = self
            .circular_sccs
            .iter()
            .copied()
            .filter(|scc_index| {
                let scc = &self.sccs[*scc_index];
                let component_is_dirty = scc.iter().any(|node| dirty_nodes[*node]);
                let component_has_only_circular_fallbacks = scc.iter().all(|node| {
                    self.formula_fallback_codes[*node].len() == 1
                        && self.formula_fallback_codes[*node].contains("circular_reference")
                });
                let component_has_cached_values = scc
                    .iter()
                    .all(|node| cached_value_cells.contains(&self.formula_cells[*node]));
                !component_is_dirty
                    && component_has_only_circular_fallbacks
                    && component_has_cached_values
            })
            .collect::<HashSet<_>>();

        if allowed_sccs.is_empty() {
            return;
        }

        for scc_index in &allowed_sccs {
            for node in &self.sccs[*scc_index] {
                self.formula_fallback_codes[*node].remove("circular_reference");
                self.formula_has_fallback[*node] = !self.formula_fallback_codes[*node].is_empty();
            }
        }

        let allowed_circular_locations = self
            .fallback_details
            .iter()
            .filter(|detail| {
                detail.code == "circular_reference"
                    && detail
                        .circular_component
                        .map(|component| allowed_sccs.contains(&component))
                        .unwrap_or(false)
            })
            .filter_map(|detail| detail.location.clone())
            .collect::<HashSet<_>>();

        self.fallback_reasons.retain(|reason| {
            reason.code != "circular_reference"
                || reason
                    .location
                    .as_ref()
                    .map(|location| !allowed_circular_locations.contains(location))
                    .unwrap_or(true)
        });
        self.fallback_details.retain(|detail| {
            detail.code != "circular_reference"
                || detail
                    .circular_component
                    .map(|component| !allowed_sccs.contains(&component))
                    .unwrap_or(true)
        });
        self.has_cached_circular_allowances = true;
    }

    pub(crate) fn dirty_summary(&self, changed_cells: &[CellId]) -> DirtySummary {
        if self.formula_cells.is_empty() || changed_cells.is_empty() {
            return DirtySummary {
                dirty_formula_cells: 0,
                planned_reusable_formula_cells: self.formula_count(),
            };
        }

        let mut dirty_nodes = vec![false; self.formula_cells.len()];
        let mut seen_cells = HashSet::new();
        let mut queue = VecDeque::new();

        for cell in changed_cells {
            queue.push_back(*cell);
            if let Some(node) = self.formula_node_by_cell.get(cell) {
                self.mark_node_and_scc(*node, &mut dirty_nodes, &mut queue);
            }
        }

        while let Some(cell) = queue.pop_front() {
            if !seen_cells.insert(cell) {
                continue;
            }

            if let Some(dependents) = self.dependents_by_cell.get(&cell) {
                for dependent in dependents {
                    self.mark_node_and_scc(*dependent, &mut dirty_nodes, &mut queue);
                }
            }

            if let Some(range_dependents) = self.range_dependents_by_sheet.get(&cell.sheet) {
                for dependency in range_dependents {
                    if dependency.range.contains(cell) {
                        self.mark_node_and_scc(dependency.dependent, &mut dirty_nodes, &mut queue);
                    }
                }
            }
        }

        let dirty_formula_cells = dirty_nodes.iter().filter(|dirty| **dirty).count();
        DirtySummary {
            dirty_formula_cells,
            planned_reusable_formula_cells: self
                .formula_count()
                .saturating_sub(dirty_formula_cells),
        }
    }

    fn dirty_nodes(&self, changed_cells: &[CellId]) -> Vec<bool> {
        if self.formula_cells.is_empty() || changed_cells.is_empty() {
            return vec![false; self.formula_cells.len()];
        }

        let mut dirty_nodes = vec![false; self.formula_cells.len()];
        let mut seen_cells = HashSet::new();
        let mut queue = VecDeque::new();

        for cell in changed_cells {
            queue.push_back(*cell);
            if let Some(node) = self.formula_node_by_cell.get(cell) {
                self.mark_node_and_scc(*node, &mut dirty_nodes, &mut queue);
            }
        }

        while let Some(cell) = queue.pop_front() {
            if !seen_cells.insert(cell) {
                continue;
            }

            if let Some(dependents) = self.dependents_by_cell.get(&cell) {
                for dependent in dependents {
                    self.mark_node_and_scc(*dependent, &mut dirty_nodes, &mut queue);
                }
            }

            if let Some(range_dependents) = self.range_dependents_by_sheet.get(&cell.sheet) {
                for dependency in range_dependents {
                    if dependency.range.contains(cell) {
                        self.mark_node_and_scc(dependency.dependent, &mut dirty_nodes, &mut queue);
                    }
                }
            }
        }

        dirty_nodes
    }

    fn mark_node_and_scc(
        &self,
        node: usize,
        dirty_nodes: &mut [bool],
        queue: &mut VecDeque<CellId>,
    ) {
        let scc_index = self.scc_by_node[node];
        for member in &self.sccs[scc_index] {
            if !dirty_nodes[*member] {
                dirty_nodes[*member] = true;
                queue.push_back(self.formula_cells[*member]);
            }
        }
    }
}

fn formula_index_for_cell(model: &Model<'_>, sheet: u32, row: i32, column: i32) -> Option<i32> {
    let worksheet = model.workbook.worksheet(sheet).ok()?;
    worksheet.cell(row, column)?.get_formula()
}

fn collect_node_dependencies(
    model: &Model<'_>,
    node: &Node,
    context: CellId,
    dependencies: &mut FormulaDependencies,
) {
    match node {
        Node::ReferenceKind {
            sheet_index,
            absolute_row,
            absolute_column,
            row,
            column,
            ..
        } => dependencies.cells.push(CellId::new(
            *sheet_index,
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
            dependencies.ranges.push(RangeRef {
                sheet: *sheet_index,
                start_row: start_row.min(end_row),
                start_column: start_column.min(end_column),
                end_row: start_row.max(end_row),
                end_column: start_column.max(end_column),
            });
        }
        Node::WrongReferenceKind { .. } | Node::WrongRangeKind { .. } => {
            dependencies.fallbacks.push(FormulaFallback {
                code: "unsupported_reference",
                message: "Formula contains a reference IronCalc marked invalid.".to_string(),
            });
        }
        Node::OpRangeKind { left, right }
        | Node::OpConcatenateKind { left, right }
        | Node::OpSumKind { left, right, .. }
        | Node::OpProductKind { left, right, .. }
        | Node::OpPowerKind { left, right }
        | Node::CompareKind { left, right, .. } => {
            collect_node_dependencies(model, left, context, dependencies);
            collect_node_dependencies(model, right, context, dependencies);
            if matches!(node, Node::OpRangeKind { .. }) {
                dependencies.fallbacks.push(FormulaFallback {
                    code: "dynamic_reference",
                    message: "Formula builds a range dynamically.".to_string(),
                });
            }
        }
        Node::FunctionKind { kind, args } => {
            let function_name = format!("{kind:?}");
            match function_name.as_str() {
                "Offset" => {
                    // A literal-offset OFFSET (e.g. OFFSET(A1, 0, 3)) always
                    // resolves to the same cell regardless of any other
                    // cell's value, so it's just an ordinary range
                    // dependency in disguise. Only OFFSET calls whose shift
                    // amount itself depends on another cell (e.g.
                    // OFFSET(A1, 0, B1)) are genuinely dynamic.
                    if let Some(range) = try_resolve_static_offset(args, context) {
                        dependencies.ranges.push(range);
                        return;
                    }
                    dependencies.fallbacks.push(FormulaFallback {
                        code: "dynamic_reference",
                        message: "Offset can change dependencies at runtime.".to_string(),
                    });
                }
                "Indirect" => dependencies.fallbacks.push(FormulaFallback {
                    code: "dynamic_reference",
                    message: "Indirect can change dependencies at runtime.".to_string(),
                }),
                "Rand" | "Randbetween" | "Now" | "Today" => {
                    dependencies.fallbacks.push(FormulaFallback {
                        code: "volatile_formula",
                        message: format!(
                            "{function_name} is volatile and is not result-cache safe."
                        ),
                    });
                }
                _ => {}
            }
            for arg in args {
                collect_node_dependencies(model, arg, context, dependencies);
            }
        }
        Node::InvalidFunctionKind { name, args } if name.eq_ignore_ascii_case("sumproduct") => {
            for arg in args {
                collect_node_dependencies(model, arg, context, dependencies);
            }
        }
        Node::InvalidFunctionKind { name, args } => {
            dependencies.fallbacks.push(FormulaFallback {
                code: "unsupported_formula",
                message: format!("Formula uses unsupported function {name}."),
            });
            for arg in args {
                collect_node_dependencies(model, arg, context, dependencies);
            }
        }
        Node::DefinedNameKind((name, scope, formula)) => {
            collect_defined_name_dependencies(model, name, *scope, formula, context, dependencies);
        }
        Node::TableNameKind(name) => dependencies.fallbacks.push(FormulaFallback {
            code: "unsupported_reference",
            message: format!("Table reference {name} is treated as a graph boundary in V1."),
        }),
        Node::WrongVariableKind(name) => dependencies.fallbacks.push(FormulaFallback {
            code: "unsupported_reference",
            message: format!("Variable name {name} could not be resolved."),
        }),
        Node::ImplicitIntersection { child, .. } | Node::UnaryKind { right: child, .. } => {
            collect_node_dependencies(model, child, context, dependencies);
        }
        Node::ParseErrorKind { message, .. } => dependencies.fallbacks.push(FormulaFallback {
            code: "unsupported_formula",
            message: format!("Formula parse error: {message}."),
        }),
        Node::ArrayKind(_)
        | Node::BooleanKind(_)
        | Node::NumberKind(_)
        | Node::StringKind(_)
        | Node::ErrorKind(_)
        | Node::EmptyArgKind => {}
    }
}

fn collect_defined_name_dependencies(
    model: &Model<'_>,
    name: &str,
    scope: Option<u32>,
    formula: &str,
    context: CellId,
    dependencies: &mut FormulaDependencies,
) {
    let reference = formula.trim().trim_start_matches('=').trim();
    if reference.is_empty() || reference.starts_with('#') {
        dependencies.fallbacks.push(FormulaFallback {
            code: "unsupported_reference",
            message: format!("Defined name {name} points to an invalid or empty reference."),
        });
        return;
    }

    if reference.parse::<f64>().is_ok()
        || reference.eq_ignore_ascii_case("TRUE")
        || reference.eq_ignore_ascii_case("FALSE")
        || (reference.starts_with('"') && reference.ends_with('"'))
    {
        return;
    }

    let default_sheet = scope
        .or(Some(context.sheet))
        .and_then(|sheet| model.workbook.worksheet(sheet).ok())
        .map(|worksheet| worksheet.name.as_str())
        .unwrap_or_default();

    if let Some(cell) = cell_id_from_a1_reference(model, reference, default_sheet) {
        dependencies.cells.push(cell);
        return;
    }

    if let Some(range) = range_ref_from_a1_reference(model, reference, default_sheet) {
        dependencies.ranges.push(range);
        return;
    }

    dependencies.fallbacks.push(FormulaFallback {
        code: "unsupported_reference",
        message: format!(
            "Defined name {name} uses a formula or reference shape WarpSpeed cannot graph yet."
        ),
    });
}

fn absolute_coord(value: i32, context: i32, is_absolute: bool) -> i32 {
    if is_absolute {
        value
    } else {
        value + context
    }
}

struct StaticOffsetBase {
    sheet: u32,
    start_row: i32,
    start_column: i32,
    rows: i32,
    columns: i32,
}

/// Resolves OFFSET(reference, rows, cols[, height[, width]]) to a fixed
/// RangeRef when every argument that affects the result is a compile-time
/// constant: a literal reference/range base, and literal numeric (optionally
/// unary-minus) shift/height/width arguments. Returns None the moment any
/// argument could depend on another cell's value, since then the true target
/// can change on any recalc and must stay a `dynamic_reference` fallback.
fn try_resolve_static_offset(args: &[Node], context: CellId) -> Option<RangeRef> {
    if args.len() < 3 || args.len() > 5 {
        return None;
    }

    let base = try_static_offset_base(&args[0], context)?;
    let row_offset = try_constant_i32(&args[1])?;
    let column_offset = try_constant_i32(&args[2])?;
    let height = match args.get(3) {
        None | Some(Node::EmptyArgKind) => base.rows,
        Some(node) => try_constant_i32(node)?,
    };
    let width = match args.get(4) {
        None | Some(Node::EmptyArgKind) => base.columns,
        Some(node) => try_constant_i32(node)?,
    };
    if height <= 0 || width <= 0 {
        return None;
    }

    let start_row = base.start_row + row_offset;
    let start_column = base.start_column + column_offset;
    Some(RangeRef {
        sheet: base.sheet,
        start_row,
        start_column,
        end_row: start_row + height - 1,
        end_column: start_column + width - 1,
    })
}

fn try_static_offset_base(node: &Node, context: CellId) -> Option<StaticOffsetBase> {
    match node {
        Node::ReferenceKind {
            sheet_index,
            absolute_row,
            absolute_column,
            row,
            column,
            ..
        } => Some(StaticOffsetBase {
            sheet: *sheet_index,
            start_row: absolute_coord(*row, context.row, *absolute_row),
            start_column: absolute_coord(*column, context.column, *absolute_column),
            rows: 1,
            columns: 1,
        }),
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
            let row_a = absolute_coord(*row1, context.row, *absolute_row1);
            let row_b = absolute_coord(*row2, context.row, *absolute_row2);
            let column_a = absolute_coord(*column1, context.column, *absolute_column1);
            let column_b = absolute_coord(*column2, context.column, *absolute_column2);
            let (start_row, end_row) = (row_a.min(row_b), row_a.max(row_b));
            let (start_column, end_column) = (column_a.min(column_b), column_a.max(column_b));
            Some(StaticOffsetBase {
                sheet: *sheet_index,
                start_row,
                start_column,
                rows: end_row - start_row + 1,
                columns: end_column - start_column + 1,
            })
        }
        _ => None,
    }
}

/// Only a bare numeric literal or its unary negation counts as a compile-time
/// constant here -- anything else (a cell reference, a nested formula) could
/// change value on a future recalc, which is exactly the case that must stay
/// a fallback rather than be silently baked into the dependency graph.
fn try_constant_i32(node: &Node) -> Option<i32> {
    match node {
        Node::NumberKind(value) => Some(*value as i32),
        Node::UnaryKind {
            kind: OpUnary::Minus,
            right,
        } => try_constant_i32(right).map(|value| -value),
        _ => None,
    }
}

fn cell_id_from_a1_reference(
    model: &Model<'_>,
    reference: &str,
    default_sheet: &str,
) -> Option<CellId> {
    let (sheet_name, row, column) = parse_a1_cell(reference, default_sheet)?;
    Some(CellId::new(
        sheet_index_by_name(model, &sheet_name)?,
        row,
        column,
    ))
}

fn range_ref_from_a1_reference(
    model: &Model<'_>,
    reference: &str,
    default_sheet: &str,
) -> Option<RangeRef> {
    let cleaned = reference.trim();
    let mut parts = cleaned.split(':');
    let left = parts.next()?;
    let right = parts.next().unwrap_or(left);
    if parts.next().is_some() {
        return None;
    }

    let (left_sheet, start_row, start_column) = parse_a1_cell(left, default_sheet)?;
    let (right_sheet, end_row, end_column) = parse_a1_cell(right, &left_sheet)?;
    if !left_sheet.eq_ignore_ascii_case(&right_sheet) {
        return None;
    }

    Some(RangeRef {
        sheet: sheet_index_by_name(model, &left_sheet)?,
        start_row: start_row.min(end_row),
        start_column: start_column.min(end_column),
        end_row: start_row.max(end_row),
        end_column: start_column.max(end_column),
    })
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
    sheet_name.trim().trim_matches('\'').replace("''", "'")
}

fn sheet_index_by_name(model: &Model<'_>, name: &str) -> Option<u32> {
    model
        .workbook
        .worksheets
        .iter()
        .position(|worksheet| worksheet.name.eq_ignore_ascii_case(name))
        .map(|index| index as u32)
}

fn structure_hash(model: &Model<'_>, formula_entries: &[(CellId, i32)]) -> String {
    let mut hasher = Sha256::new();
    for worksheet in &model.workbook.worksheets {
        hasher.update(worksheet.name.as_bytes());
        hasher.update([0_u8]);
    }

    let mut sorted_entries = formula_entries.to_vec();
    sorted_entries.sort_by_key(|(cell, _)| *cell);
    for (cell, formula_index) in sorted_entries {
        hasher.update(cell.sheet.to_le_bytes());
        hasher.update(cell.row.to_le_bytes());
        hasher.update(cell.column.to_le_bytes());
        if let Some(formula) = model
            .workbook
            .worksheets
            .get(cell.sheet as usize)
            .and_then(|worksheet| worksheet.shared_formulas.get(formula_index as usize))
        {
            hasher.update(formula.as_bytes());
        }
        hasher.update([0_u8]);
    }

    format!("{:x}", hasher.finalize())
}

fn tarjan_scc(adjacency: &[Vec<usize>]) -> Vec<Vec<usize>> {
    struct Tarjan<'a> {
        adjacency: &'a [Vec<usize>],
        next_index: usize,
        stack: Vec<usize>,
        indices: Vec<Option<usize>>,
        lowlink: Vec<usize>,
        on_stack: Vec<bool>,
        sccs: Vec<Vec<usize>>,
    }

    impl<'a> Tarjan<'a> {
        fn strong_connect(&mut self, node: usize) {
            self.indices[node] = Some(self.next_index);
            self.lowlink[node] = self.next_index;
            self.next_index += 1;
            self.stack.push(node);
            self.on_stack[node] = true;

            for dependent in &self.adjacency[node] {
                if self.indices[*dependent].is_none() {
                    self.strong_connect(*dependent);
                    self.lowlink[node] = self.lowlink[node].min(self.lowlink[*dependent]);
                } else if self.on_stack[*dependent] {
                    self.lowlink[node] =
                        self.lowlink[node].min(self.indices[*dependent].unwrap_or(0));
                }
            }

            if self.lowlink[node] == self.indices[node].unwrap_or(0) {
                let mut scc = Vec::new();
                while let Some(member) = self.stack.pop() {
                    self.on_stack[member] = false;
                    scc.push(member);
                    if member == node {
                        break;
                    }
                }
                self.sccs.push(scc);
            }
        }
    }

    let mut tarjan = Tarjan {
        adjacency,
        next_index: 0,
        stack: Vec::new(),
        indices: vec![None; adjacency.len()],
        lowlink: vec![0; adjacency.len()],
        on_stack: vec![false; adjacency.len()],
        sccs: Vec::new(),
    };

    for node in 0..adjacency.len() {
        if tarjan.indices[node].is_none() {
            tarjan.strong_connect(node);
        }
    }

    tarjan.sccs
}

fn add_fallback(
    fallback_reasons: &mut Vec<FallbackReason>,
    fallback_keys: &mut HashSet<(String, Option<String>)>,
    fallback_details: &mut Vec<FallbackDetail>,
    code: &str,
    message: String,
    location: Option<String>,
    formula: Option<String>,
    circular_component: Option<usize>,
    circular_component_size: Option<usize>,
) {
    let key = (code.to_string(), location.clone());
    if fallback_keys.insert(key) {
        fallback_reasons.push(FallbackReason {
            code: code.to_string(),
            message: message.clone(),
            location: location.clone(),
        });
    }

    fallback_details.push(FallbackDetail {
        code: code.to_string(),
        message,
        location,
        formula,
        circular_component,
        circular_component_size,
    });
}

fn format_cell_location(model: &Model<'_>, cell: CellId) -> String {
    let sheet = model
        .workbook
        .worksheets
        .get(cell.sheet as usize)
        .map(|worksheet| worksheet.name.as_str())
        .unwrap_or("<unknown>");
    let column = number_to_column(cell.column).unwrap_or_else(|| format!("C{}", cell.column));
    format!("{sheet}!{column}{}", cell.row)
}

fn formula_text_for_cell(model: &Model<'_>, cell: CellId) -> Option<String> {
    model
        .get_cell_formula(cell.sheet, cell.row, cell.column)
        .ok()
        .flatten()
}

#[cfg(test)]
mod tests {
    use ironcalc::base::Model;

    use super::*;

    #[test]
    fn extracts_direct_dependencies_and_dirty_closure() {
        let mut model = Model::new_empty("direct", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "10".to_string()).unwrap();
        model.set_user_input(0, 1, 2, "=A1*2".to_string()).unwrap();
        model.set_user_input(0, 1, 3, "=B1+1".to_string()).unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(graph.coverage().formula_cells, 2);
        assert_eq!(
            graph.dirty_summary(&[CellId::new(0, 1, 1)]),
            DirtySummary {
                dirty_formula_cells: 2,
                planned_reusable_formula_cells: 0,
            }
        );
    }

    #[test]
    fn keeps_range_dependencies_compressed_but_queryable() {
        let mut model = Model::new_empty("range", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "10".to_string()).unwrap();
        model.set_user_input(0, 2, 1, "20".to_string()).unwrap();
        model.set_user_input(0, 3, 1, "30".to_string()).unwrap();
        model
            .set_user_input(0, 1, 2, "=SUM(A1:A3)".to_string())
            .unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(
            graph.dirty_summary(&[CellId::new(0, 2, 1)]),
            DirtySummary {
                dirty_formula_cells: 1,
                planned_reusable_formula_cells: 0,
            }
        );
    }

    #[test]
    fn treats_sumproduct_as_supported_range_dependency() {
        let mut model = Model::new_empty("sumproduct", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "10".to_string()).unwrap();
        model.set_user_input(0, 1, 2, "20".to_string()).unwrap();
        model.set_user_input(0, 2, 1, "2".to_string()).unwrap();
        model.set_user_input(0, 2, 2, "3".to_string()).unwrap();
        model
            .set_user_input(0, 3, 1, "=SUMPRODUCT(A1:B1,A2:B2)".to_string())
            .unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(graph.coverage().formula_cells, 1);
        assert_eq!(graph.coverage().fallback_formula_cells, 0);
        assert_eq!(
            graph
                .dirty_summary(&[CellId::new(0, 1, 2)])
                .dirty_formula_cells,
            1
        );
    }

    #[test]
    fn extracts_cross_sheet_dependencies() {
        let mut model = Model::new_empty("cross-sheet", "en", "UTC", "en").unwrap();
        model.add_sheet("Debt").unwrap();
        model.set_user_input(0, 1, 1, "10".to_string()).unwrap();
        model
            .set_user_input(1, 1, 1, "=Sheet1!A1*2".to_string())
            .unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(
            graph
                .dirty_summary(&[CellId::new(0, 1, 1)])
                .dirty_formula_cells,
            1
        );
    }

    #[test]
    fn extracts_dependencies_through_workbook_defined_names() {
        let mut model = Model::new_empty("defined-name", "en", "UTC", "en").unwrap();
        model.add_sheet("Debt").unwrap();
        model
            .new_defined_name("DebtBalance", None, "Debt!$C$5")
            .unwrap();
        model.set_user_input(1, 5, 3, "10".to_string()).unwrap();
        model
            .set_user_input(0, 1, 1, "=DebtBalance*2".to_string())
            .unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(graph.coverage().fallback_formula_cells, 0);
        assert_eq!(
            graph
                .dirty_summary(&[CellId::new(1, 5, 3)])
                .dirty_formula_cells,
            1
        );
    }

    #[test]
    fn extracts_dependencies_through_sheet_scoped_range_names() {
        let mut model = Model::new_empty("scoped-name", "en", "UTC", "en").unwrap();
        model.add_sheet("Debt").unwrap();
        model
            .new_defined_name("DebtRows", Some(1), "Debt!$A$1:$A$3")
            .unwrap();
        model.set_user_input(1, 1, 1, "10".to_string()).unwrap();
        model.set_user_input(1, 2, 1, "20".to_string()).unwrap();
        model.set_user_input(1, 3, 1, "30".to_string()).unwrap();
        model
            .set_user_input(1, 1, 2, "=SUM(DebtRows)".to_string())
            .unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(graph.coverage().fallback_formula_cells, 0);
        assert_eq!(
            graph
                .dirty_summary(&[CellId::new(1, 2, 1)])
                .dirty_formula_cells,
            1
        );
    }

    #[test]
    fn marks_dynamic_and_volatile_formulas_as_fallbacks() {
        let mut model = Model::new_empty("fallbacks", "en", "UTC", "en").unwrap();
        model
            .set_user_input(0, 1, 1, "=INDIRECT(\"A2\")".to_string())
            .unwrap();
        model
            .set_user_input(0, 2, 1, "=RAND()".to_string())
            .unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(graph.coverage().formula_cells, 2);
        assert_eq!(graph.coverage().fallback_formula_cells, 2);
        assert!(!graph.is_result_cache_safe());
        assert!(graph
            .fallback_details()
            .iter()
            .any(|detail| detail.formula.as_deref() == Some("=INDIRECT(\"A2\")")));
    }

    #[test]
    fn treats_literal_offset_as_a_supported_range_dependency() {
        let mut model = Model::new_empty("static-offset", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "10".to_string()).unwrap();
        model.set_user_input(0, 1, 2, "20".to_string()).unwrap();
        model
            .set_user_input(0, 1, 3, "=OFFSET(A1,0,1)".to_string())
            .unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(graph.coverage().formula_cells, 1);
        assert_eq!(graph.coverage().fallback_formula_cells, 0);
        assert!(graph.is_result_cache_safe());
        // The formula's real dependency is B1 (A1 shifted right by one
        // column), not A1 itself -- OFFSET only uses A1's location.
        assert_eq!(
            graph
                .dirty_summary(&[CellId::new(0, 1, 2)])
                .dirty_formula_cells,
            1
        );
    }

    #[test]
    fn treats_offset_with_negative_shift_and_explicit_size_as_supported() {
        let mut model = Model::new_empty("static-offset-size", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 5, 5, "1".to_string()).unwrap();
        model.set_user_input(0, 3, 3, "2".to_string()).unwrap();
        model
            .set_user_input(0, 1, 1, "=SUM(OFFSET(E5,-2,-2,1,1))".to_string())
            .unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(graph.coverage().fallback_formula_cells, 0);
        assert_eq!(
            graph
                .dirty_summary(&[CellId::new(0, 3, 3)])
                .dirty_formula_cells,
            1
        );
    }

    #[test]
    fn treats_offset_with_cell_dependent_shift_as_dynamic_reference_fallback() {
        let mut model = Model::new_empty("dynamic-offset", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "10".to_string()).unwrap();
        model.set_user_input(0, 1, 2, "1".to_string()).unwrap();
        model
            .set_user_input(0, 1, 3, "=OFFSET(A1,0,B1)".to_string())
            .unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(graph.coverage().fallback_formula_cells, 1);
        assert!(graph
            .fallback_reasons()
            .iter()
            .any(|reason| reason.code == "dynamic_reference"
                && reason.message.contains("Offset")));
    }

    #[test]
    fn marks_circular_sccs_as_fallbacks_and_dirty_together() {
        let mut model = Model::new_empty("cycle", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "=B1".to_string()).unwrap();
        model.set_user_input(0, 1, 2, "=A1".to_string()).unwrap();

        let graph = build_dependency_graph(&model);
        assert_eq!(graph.coverage().fallback_formula_cells, 2);
        assert_eq!(
            graph
                .dirty_summary(&[CellId::new(0, 1, 1)])
                .dirty_formula_cells,
            2
        );
        let details = graph.fallback_details();
        assert_eq!(details.len(), 2);
        assert!(details
            .iter()
            .all(|detail| detail.code == "circular_reference"));
        assert!(details
            .iter()
            .all(|detail| detail.circular_component_size == Some(2)));
        assert!(details
            .iter()
            .any(|detail| detail.formula.as_deref() == Some("=A1")));
    }

    #[test]
    fn allows_clean_cached_circular_components() {
        let mut model = Model::new_empty("cached-cycle", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "=B1".to_string()).unwrap();
        model.set_user_input(0, 1, 2, "=A1".to_string()).unwrap();

        let mut graph = build_dependency_graph(&model);
        let cached_cells = [CellId::new(0, 1, 1), CellId::new(0, 1, 2)]
            .into_iter()
            .collect::<HashSet<_>>();
        graph.allow_cached_circular_values(&cached_cells, &[]);

        assert_eq!(graph.coverage().fallback_formula_cells, 0);
        assert!(graph.fallback_reasons().is_empty());
        assert!(graph.fallback_details().is_empty());
        assert_eq!(graph.supported_formula_cells().count(), 2);
    }

    #[test]
    fn keeps_dirty_cached_circular_components_as_fallbacks() {
        let mut model = Model::new_empty("dirty-cycle", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "1".to_string()).unwrap();
        model.set_user_input(0, 1, 2, "=A1+C1".to_string()).unwrap();
        model.set_user_input(0, 1, 3, "=B1".to_string()).unwrap();

        let mut graph = build_dependency_graph(&model);
        let cached_cells = [CellId::new(0, 1, 2), CellId::new(0, 1, 3)]
            .into_iter()
            .collect::<HashSet<_>>();
        graph.allow_cached_circular_values(&cached_cells, &[CellId::new(0, 1, 1)]);

        assert_eq!(graph.coverage().fallback_formula_cells, 2);
        assert!(graph
            .fallback_reasons()
            .iter()
            .all(|reason| reason.code == "circular_reference"));
    }
}
