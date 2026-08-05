use std::collections::HashMap;
use std::fs::{self, File};
use std::io::{self, Read, Write};
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use regex::Regex;
use zip::result::ZipError;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::data_tables::{build_data_table_region, DataTableRegion};
use crate::model::{EngineError, FallbackDetail, FallbackReason};

#[derive(Clone, Debug, Default)]
pub(crate) struct ImportFallbacks {
    pub fallback_formula_cells: usize,
    pub fallback_reasons: Vec<FallbackReason>,
    pub fallback_details: Vec<FallbackDetail>,
}

impl ImportFallbacks {
    pub(crate) fn is_empty(&self) -> bool {
        self.fallback_formula_cells == 0
            && self.fallback_reasons.is_empty()
            && self.fallback_details.is_empty()
    }
}

#[derive(Debug)]
pub(crate) struct SanitizedWorkbook {
    pub path: PathBuf,
    pub fallbacks: ImportFallbacks,
    pub data_tables: Vec<DataTableRegion>,
}

#[derive(Debug)]
pub(crate) struct SanitizedSheet {
    xml: String,
    fallbacks: ImportFallbacks,
    data_tables: Vec<DataTableRegion>,
    /// WS.LIVE cells reduced to their cached values. These are not fallbacks
    /// -- the kernel recomputes them -- but they still mean the sheet XML was
    /// rewritten and must replace the original.
    stripped_live_cells: usize,
}

#[derive(Debug)]
struct SheetRecord {
    name: String,
    relationship_id: String,
}

pub(crate) fn sanitize_data_table_formulas(
    workbook_path: &str,
    workbook_hash: &str,
) -> Result<Option<SanitizedWorkbook>, EngineError> {
    let source =
        File::open(workbook_path).map_err(|err| EngineError::WorkbookLoad(err.to_string()))?;
    let mut archive = ZipArchive::new(source).map_err(zip_error)?;
    let sheet_names_by_part = sheet_names_by_part(&mut archive)?;
    let mut replacements = HashMap::new();
    let mut fallbacks = ImportFallbacks::default();
    let mut data_tables = Vec::new();

    for index in 0..archive.len() {
        let mut file = archive.by_index(index).map_err(zip_error)?;
        let part_name = file.name().to_string();
        if !is_worksheet_xml_part(&part_name) {
            continue;
        }

        let mut xml = String::new();
        file.read_to_string(&mut xml)
            .map_err(|err| EngineError::WorkbookLoad(err.to_string()))?;

        // WS.LIVE cells need stripping too, and a workbook whose tables
        // have all been converted contains no dataTable markers at all.
        if !xml.contains("dataTable") && !contains_ignore_ascii_case(&xml, LIVE_MARKER) {
            continue;
        }

        let sheet_name = sheet_names_by_part
            .get(&part_name)
            .map(String::as_str)
            .unwrap_or(part_name.as_str());
        let sanitized = strip_data_table_formulas_from_sheet_xml(sheet_name, &part_name, &xml);

        if sanitized.fallbacks.fallback_formula_cells > 0 || sanitized.stripped_live_cells > 0 {
            fallbacks.fallback_formula_cells += sanitized.fallbacks.fallback_formula_cells;
            fallbacks
                .fallback_reasons
                .extend(sanitized.fallbacks.fallback_reasons);
            fallbacks
                .fallback_details
                .extend(sanitized.fallbacks.fallback_details);
            data_tables.extend(sanitized.data_tables);
            replacements.insert(part_name, sanitized.xml.into_bytes());
        }
    }

    // A workbook whose tables are all converted produces no fallbacks, but
    // its WS.LIVE cells still had to be stripped, so the rewrite is required.
    if fallbacks.is_empty() && replacements.is_empty() {
        return Ok(None);
    }

    let sanitized_path = sanitized_workbook_path(workbook_path, workbook_hash)?;
    rewrite_zip(workbook_path, &sanitized_path, replacements)?;

    Ok(Some(SanitizedWorkbook {
        path: sanitized_path,
        fallbacks,
        data_tables,
    }))
}

pub(crate) fn strip_data_table_formulas_from_sheet_xml(
    sheet_name: &str,
    part_name: &str,
    xml: &str,
) -> SanitizedSheet {
    let cell_regex =
        Regex::new(r#"(?s)<c\b(?P<attrs>[^>]*)>(?P<body>.*?)</c>"#).expect("cell regex is valid");
    let mut sanitized = String::with_capacity(xml.len());
    let mut fallbacks = ImportFallbacks::default();
    let mut data_tables = Vec::new();
    let mut stripped_live_cells = 0;
    let mut cursor = 0;

    for captures in cell_regex.captures_iter(xml) {
        let Some(full_match) = captures.get(0) else {
            continue;
        };
        let Some(body_match) = captures.name("body") else {
            continue;
        };
        let Some(attrs_match) = captures.name("attrs") else {
            continue;
        };

        let body_text = body_match.as_str();
        if !body_text.contains("dataTable") && !contains_ignore_ascii_case(body_text, LIVE_MARKER) {
            continue;
        }

        let cell_attrs = parse_xml_attributes(attrs_match.as_str());
        let cell_address = cell_attrs.get("r").cloned();
        let stripped = strip_data_table_formulas_from_cell_body(body_text);
        let body_rewritten = stripped.body != body_text;
        if stripped.formulas.is_empty() && !body_rewritten {
            continue;
        }
        if stripped.formulas.is_empty() {
            stripped_live_cells += 1;
        }

        sanitized.push_str(&xml[cursor..body_match.start()]);
        sanitized.push_str(&stripped.body);
        cursor = body_match.end();

        for formula in stripped.formulas {
            fallbacks.fallback_formula_cells += 1;
            let formula_ref = formula.get("ref").cloned().or_else(|| cell_address.clone());
            let location = formula_ref
                .map(|reference| format!("{sheet_name}!{reference}"))
                .or_else(|| Some(part_name.to_string()));
            fallbacks.fallback_reasons.push(FallbackReason {
                code: "data_table_formula".to_string(),
                message:
                    "Excel native data table formula was preserved as its cached value for import."
                        .to_string(),
                location: location.clone(),
            });
            fallbacks.fallback_details.push(FallbackDetail {
                code: "data_table_formula".to_string(),
                message:
                    "Excel native data table formula was preserved as its cached value for import."
                        .to_string(),
                location,
                formula: Some(data_table_formula_summary(&formula)),
                circular_component: None,
                circular_component_size: None,
            });

            if let Some(range_address) = formula.get("ref") {
                data_tables.push(build_data_table_region(
                    sheet_name,
                    part_name,
                    cell_address.as_deref().unwrap_or("<unknown>"),
                    range_address,
                    formula.get("r1").map(String::as_str),
                    formula.get("r2").map(String::as_str),
                    formula_bool(&formula, "dt2D"),
                    formula_bool(&formula, "dtr"),
                ));
            }
        }

        let _ = full_match;
    }

    if fallbacks.fallback_formula_cells == 0 && stripped_live_cells == 0 {
        return SanitizedSheet {
            xml: xml.to_string(),
            fallbacks,
            data_tables,
            stripped_live_cells,
        };
    }

    sanitized.push_str(&xml[cursor..]);
    SanitizedSheet {
        xml: sanitized,
        fallbacks,
        data_tables,
        stripped_live_cells,
    }
}

/// Uppercase form of the live-cell function name, for case-insensitive
/// scanning. Held as a constant so the scan never has to allocate.
const LIVE_MARKER: &[u8] = b"WS.LIVE";

/// Case-insensitive substring search that allocates nothing. The naive
/// version -- uppercasing the haystack -- copied every sheet's XML and every
/// cell body, which on a large workbook cost more than the whole sanitize
/// pass it was guarding.
fn contains_ignore_ascii_case(haystack: &str, needle_upper: &[u8]) -> bool {
    let bytes = haystack.as_bytes();
    if bytes.len() < needle_upper.len() {
        return false;
    }

    bytes
        .windows(needle_upper.len())
        .any(|window| window.eq_ignore_ascii_case(needle_upper))
}

/// True when the `<f>` element starting at `open_end` contains a WS.LIVE
/// call. Self-closing formula tags have no text and never qualify.
fn formula_body_is_live_cell(body: &str, open_end: usize, open_tag: &str) -> bool {
    if open_tag.trim_end().ends_with("/>") {
        return false;
    }

    let Some(relative_close) = body[open_end..].find("</f>") else {
        return false;
    };

    contains_ignore_ascii_case(&body[open_end..open_end + relative_close], LIVE_MARKER)
}

fn strip_data_table_formulas_from_cell_body(body: &str) -> StrippedCellBody {
    let mut stripped = String::with_capacity(body.len());
    let mut formulas = Vec::new();
    let mut cursor = 0;

    while let Some(relative_start) = body[cursor..].find("<f") {
        let start = cursor + relative_start;
        let Some(relative_open_end) = body[start..].find('>') else {
            break;
        };
        let open_end = start + relative_open_end + 1;
        let open_tag = &body[start..open_end];
        let attrs = parse_xml_attributes(open_tag);

        let is_data_table = attrs.get("t").map(String::as_str) == Some("dataTable");

        // A WS.LIVE cell is one WarpSpeed already drives. Its formula is a
        // function IronCalc has never heard of, so leaving it in place would
        // make the cell an unsupported_formula fallback and -- through
        // dependency tainting -- drag everything downstream out of the
        // writeback-safe set, making coverage *worse* after conversion.
        // Strip it to its cached value, exactly as a native data table cell
        // is stripped; the value it should display is recomputed by the
        // kernel from the host's table override.
        let is_live_cell = !is_data_table && formula_body_is_live_cell(body, open_end, open_tag);

        if !is_data_table && !is_live_cell {
            stripped.push_str(&body[cursor..open_end]);
            cursor = open_end;
            continue;
        }

        stripped.push_str(&body[cursor..start]);
        if is_data_table {
            formulas.push(attrs);
        }
        if open_tag.trim_end().ends_with("/>") {
            cursor = open_end;
        } else if let Some(relative_close_start) = body[open_end..].find("</f>") {
            cursor = open_end + relative_close_start + "</f>".len();
        } else {
            cursor = open_end;
        }
    }

    stripped.push_str(&body[cursor..]);
    StrippedCellBody {
        body: stripped,
        formulas,
    }
}

#[derive(Debug)]
struct StrippedCellBody {
    body: String,
    formulas: Vec<HashMap<String, String>>,
}

fn sheet_names_by_part(
    archive: &mut ZipArchive<File>,
) -> Result<HashMap<String, String>, EngineError> {
    let Some(workbook_xml) = read_zip_text(archive, "xl/workbook.xml")? else {
        return Ok(HashMap::new());
    };
    let Some(workbook_rels_xml) = read_zip_text(archive, "xl/_rels/workbook.xml.rels")? else {
        return Ok(HashMap::new());
    };

    let sheets = parse_workbook_sheet_records(&workbook_xml);
    let relationships = parse_workbook_relationship_targets(&workbook_rels_xml);
    let mut names_by_part = HashMap::new();

    for sheet in sheets {
        let Some(target) = relationships.get(&sheet.relationship_id) else {
            continue;
        };
        names_by_part.insert(normalize_part_path("xl", target), sheet.name);
    }

    Ok(names_by_part)
}

fn read_zip_text(
    archive: &mut ZipArchive<File>,
    part_name: &str,
) -> Result<Option<String>, EngineError> {
    match archive.by_name(part_name) {
        Ok(mut file) => {
            let mut text = String::new();
            file.read_to_string(&mut text)
                .map_err(|err| EngineError::WorkbookLoad(err.to_string()))?;
            Ok(Some(text))
        }
        Err(ZipError::FileNotFound) => Ok(None),
        Err(err) => Err(zip_error(err)),
    }
}

fn parse_workbook_sheet_records(workbook_xml: &str) -> Vec<SheetRecord> {
    let sheet_regex = Regex::new(r#"<sheet\b[^>]*/?>"#).expect("sheet regex is valid");
    let mut records = Vec::new();

    for sheet_match in sheet_regex.find_iter(workbook_xml) {
        let attrs = parse_xml_attributes(sheet_match.as_str());
        let Some(name) = attrs.get("name") else {
            continue;
        };
        let Some(relationship_id) = attrs.get("r:id") else {
            continue;
        };
        records.push(SheetRecord {
            name: xml_unescape(name),
            relationship_id: relationship_id.clone(),
        });
    }

    records
}

fn parse_workbook_relationship_targets(workbook_rels_xml: &str) -> HashMap<String, String> {
    let rel_regex = Regex::new(r#"<Relationship\b[^>]*/?>"#).expect("relationship regex is valid");
    let mut targets = HashMap::new();

    for rel_match in rel_regex.find_iter(workbook_rels_xml) {
        let attrs = parse_xml_attributes(rel_match.as_str());
        let Some(id) = attrs.get("Id") else {
            continue;
        };
        let Some(target) = attrs.get("Target") else {
            continue;
        };
        if target.contains("worksheets/") {
            targets.insert(id.clone(), xml_unescape(target));
        }
    }

    targets
}

fn parse_xml_attributes(tag: &str) -> HashMap<String, String> {
    let attr_regex = Regex::new(r#"([A-Za-z_:\-][A-Za-z0-9_:\-\.]*)\s*=\s*"([^"]*)""#)
        .expect("attribute regex is valid");
    attr_regex
        .captures_iter(tag)
        .filter_map(|captures| {
            let key = captures.get(1)?.as_str().to_string();
            let value = xml_unescape(captures.get(2)?.as_str());
            Some((key, value))
        })
        .collect()
}

fn formula_bool(attrs: &HashMap<String, String>, key: &str) -> bool {
    matches!(
        attrs.get(key).map(String::as_str),
        Some("1" | "true" | "TRUE")
    )
}

fn data_table_formula_summary(attrs: &HashMap<String, String>) -> String {
    let mut pairs = attrs
        .iter()
        .map(|(key, value)| format!("{key}={value}"))
        .collect::<Vec<_>>();
    pairs.sort();
    format!("dataTable {}", pairs.join(" "))
}

fn xml_unescape(value: &str) -> String {
    value
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
}

fn is_worksheet_xml_part(part_name: &str) -> bool {
    part_name.starts_with("xl/worksheets/")
        && part_name.ends_with(".xml")
        && !part_name.contains("/_rels/")
}

fn normalize_part_path(base_dir: &str, target: &str) -> String {
    let mut parts = Vec::new();
    if !target.starts_with('/') {
        parts.extend(base_dir.split('/').filter(|part| !part.is_empty()));
    }

    for part in target.trim_start_matches('/').split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
        }
    }

    parts.join("/")
}

fn sanitized_workbook_path(
    workbook_path: &str,
    workbook_hash: &str,
) -> Result<PathBuf, EngineError> {
    let hash_prefix: String = workbook_hash.chars().take(12).collect();
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|err| EngineError::WorkbookLoad(err.to_string()))?
        .as_nanos();
    let stem = Path::new(workbook_path)
        .file_stem()
        .and_then(|stem| stem.to_str())
        .unwrap_or("workbook")
        .chars()
        .map(|ch| {
            if ch.is_ascii_alphanumeric() || ch == '-' || ch == '_' {
                ch
            } else {
                '-'
            }
        })
        .collect::<String>();
    Ok(std::env::temp_dir().join(format!("warpspeed-{stem}-{hash_prefix}-{nanos}.xlsx")))
}

fn rewrite_zip(
    source_path: &str,
    output_path: &Path,
    mut replacements: HashMap<String, Vec<u8>>,
) -> Result<(), EngineError> {
    let source =
        File::open(source_path).map_err(|err| EngineError::WorkbookLoad(err.to_string()))?;
    let mut archive = ZipArchive::new(source).map_err(zip_error)?;
    let output =
        File::create(output_path).map_err(|err| EngineError::WorkbookLoad(err.to_string()))?;
    let mut writer = ZipWriter::new(output);

    for index in 0..archive.len() {
        let mut file = archive.by_index(index).map_err(zip_error)?;
        let name = file.name().to_string();
        let options = FileOptions::default()
            .compression_method(file.compression())
            .last_modified_time(file.last_modified());

        if name.ends_with('/') {
            writer.add_directory(name, options).map_err(zip_error)?;
            continue;
        }

        writer
            .start_file(name.clone(), options)
            .map_err(zip_error)?;
        if let Some(contents) = replacements.remove(&name) {
            writer
                .write_all(&contents)
                .map_err(|err| EngineError::WorkbookLoad(err.to_string()))?;
        } else {
            io::copy(&mut file, &mut writer)
                .map_err(|err| EngineError::WorkbookLoad(err.to_string()))?;
        }
    }

    writer.finish().map_err(zip_error)?;
    Ok(())
}

fn zip_error(err: ZipError) -> EngineError {
    EngineError::WorkbookLoad(err.to_string())
}

pub(crate) fn remove_sanitized_workbook(path: &Path) {
    let _ = fs::remove_file(path);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strips_data_table_formula_and_preserves_cached_value() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="F450"><f t="dataTable" ref="F450:J454" dt2D="1" dtr="1" r1="E32" r2="O9"/><v>42</v></c><c r="A1"><f>A2+1</f><v>3</v></c></row></sheetData></worksheet>"#;

        let sanitized = strip_data_table_formulas_from_sheet_xml(
            "Data Tables",
            "xl/worksheets/sheet1.xml",
            xml,
        );

        assert!(!sanitized.xml.contains(r#"t="dataTable""#));
        assert!(sanitized.xml.contains("<v>42</v>"));
        assert!(sanitized.xml.contains("<f>A2+1</f>"));
        assert_eq!(sanitized.fallbacks.fallback_formula_cells, 1);
        assert_eq!(sanitized.data_tables.len(), 1);
        assert_eq!(
            sanitized.fallbacks.fallback_reasons[0].location.as_deref(),
            Some("Data Tables!F450:J454")
        );
        assert_eq!(
            sanitized.fallbacks.fallback_details[0].formula.as_deref(),
            Some("dataTable dt2D=1 dtr=1 r1=E32 r2=O9 ref=F450:J454 t=dataTable")
        );
        assert_eq!(sanitized.data_tables[0].range_address, "F450:J454");
        assert_eq!(
            sanitized.data_tables[0]
                .column_input_cell
                .as_ref()
                .map(|cell| cell.address.as_str()),
            Some("E32")
        );
    }

    #[test]
    fn parses_sheet_part_names_from_workbook_relationships() {
        let workbook_xml = r#"<workbook><sheets><sheet name="Assumptions &amp; Cases" sheetId="1" r:id="rId3"/></sheets></workbook>"#;
        let workbook_rels = r#"<Relationships><Relationship Id="rId3" Target="worksheets/sheet12.xml"/></Relationships>"#;
        let sheets = parse_workbook_sheet_records(workbook_xml);
        let rels = parse_workbook_relationship_targets(workbook_rels);

        assert_eq!(sheets[0].name, "Assumptions & Cases");
        assert_eq!(
            normalize_part_path("xl", rels.get(&sheets[0].relationship_id).unwrap()),
            "xl/worksheets/sheet12.xml"
        );
    }
}
