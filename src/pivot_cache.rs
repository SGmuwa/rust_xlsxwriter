// pivot_cache - A module for creating the Excel pivotCacheDefinition.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

use std::cmp::Ordering;
use std::io::Cursor;

use crate::pivot_table::PivotTableSource;
use crate::utility;
use crate::xmlwriter::{
    escape_attribute_escapes, xml_declaration, xml_empty_tag, xml_empty_tag_only, xml_end_tag,
    xml_start_tag,
};
use crate::{ColNum, XlsxError};

// A struct to represent the pivotCacheDefinition.xml file. The cache doesn't
// contain any records and is refreshed by Excel when the file is loaded. It
// can contain the unique values of the fields, see PivotTable::set_field_items().
pub(crate) struct PivotCache {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) source: PivotTableSource,
    pub(crate) field_names: Vec<String>,
    pub(crate) field_values: Vec<Vec<PivotFieldValue>>,
}

impl PivotCache {
    // Create a new PivotCache struct instance.
    pub(crate) fn new(source: PivotTableSource) -> PivotCache {
        let writer = Cursor::new(Vec::with_capacity(2048));

        PivotCache {
            writer,
            source,
            field_names: vec![],
            field_values: vec![],
        }
    }

    // Validate that the header cells in the source data are non-blank and
    // unique (case-insensitively) since Excel uses them as the pivot table
    // field names.
    pub(crate) fn validate_field_names(&self) -> Result<(), XlsxError> {
        let mut seen_names = vec![];

        for (col_offset, name) in self.field_names.iter().enumerate() {
            let cell = utility::row_col_to_cell(
                self.source.first_row,
                self.source.first_col + col_offset as ColNum,
            );

            if name.is_empty() {
                return Err(XlsxError::PivotTableError(format!(
                    "Header cell '{}!{cell}' in the source data is blank",
                    self.source.sheet_name
                )));
            }

            let lowercase_name = name.to_lowercase();
            if seen_names.contains(&lowercase_name) {
                return Err(XlsxError::PivotTableError(format!(
                    "Header '{name}' in cell '{}!{cell}' is duplicated in the source data",
                    self.source.sheet_name
                )));
            }

            seen_names.push(lowercase_name);
        }

        Ok(())
    }

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the pivotCacheDefinition element.
        self.write_pivot_cache_definition();

        // Write the cacheSource element.
        self.write_cache_source();

        // Write the cacheFields element.
        self.write_cache_fields();

        // Close the pivotCacheDefinition tag.
        xml_end_tag(&mut self.writer, "pivotCacheDefinition");
    }

    // Write the <pivotCacheDefinition> element.
    fn write_pivot_cache_definition(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let xmlns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        let attributes = [
            ("xmlns", xmlns.to_string()),
            ("xmlns:r", xmlns_rel.to_string()),
            ("saveData", "0".to_string()),
            ("refreshOnLoad", "1".to_string()),
            ("createdVersion", "3".to_string()),
            ("refreshedVersion", "3".to_string()),
            ("minRefreshableVersion", "3".to_string()),
            ("recordCount", self.source.num_records().to_string()),
        ];

        xml_start_tag(&mut self.writer, "pivotCacheDefinition", &attributes);
    }

    // Write the <cacheSource> element.
    fn write_cache_source(&mut self) {
        let attributes = [("type", "worksheet")];

        xml_start_tag(&mut self.writer, "cacheSource", &attributes);

        // Write the worksheetSource element.
        self.write_worksheet_source();

        xml_end_tag(&mut self.writer, "cacheSource");
    }

    // Write the <worksheetSource> element.
    fn write_worksheet_source(&mut self) {
        let range = utility::cell_range(
            self.source.first_row,
            self.source.first_col,
            self.source.last_row,
            self.source.last_col,
        );

        let attributes = [
            ("ref", range),
            (
                "sheet",
                escape_attribute_escapes(&self.source.sheet_name).into_owned(),
            ),
        ];

        xml_empty_tag(&mut self.writer, "worksheetSource", &attributes);
    }

    // Write the <cacheFields> element.
    fn write_cache_fields(&mut self) {
        let attributes = [("count", self.field_names.len().to_string())];

        xml_start_tag(&mut self.writer, "cacheFields", &attributes);

        for (index, name) in self.field_names.clone().iter().enumerate() {
            let values = self.field_values.get(index).cloned().unwrap_or_default();

            // Write the cacheField element.
            self.write_cache_field(name, &values);
        }

        xml_end_tag(&mut self.writer, "cacheFields");
    }

    // Write the <cacheField> element.
    fn write_cache_field(&mut self, name: &str, values: &[PivotFieldValue]) {
        let attributes = [
            ("name", escape_attribute_escapes(name).into_owned()),
            ("numFmtId", "0".to_string()),
        ];

        xml_start_tag(&mut self.writer, "cacheField", &attributes);

        // Write the sharedItems element.
        self.write_shared_items(values);

        xml_end_tag(&mut self.writer, "cacheField");
    }

    // Write the <sharedItems> element. It is empty unless the unique values of
    // the field have been added, in ascending order. The attributes describe
    // the types of the values in the same way as Excel.
    fn write_shared_items(&mut self, values: &[PivotFieldValue]) {
        if values.is_empty() {
            xml_empty_tag_only(&mut self.writer, "sharedItems");
            return;
        }

        let has_blank = values.contains(&PivotFieldValue::Blank);
        let has_string = values
            .iter()
            .any(|value| matches!(value, PivotFieldValue::String(_)));
        let numbers: Vec<f64> = values
            .iter()
            .filter_map(|value| match value {
                PivotFieldValue::Number(number) => Some(*number),
                _ => None,
            })
            .collect();

        // The attributes are in the schema order used by Excel.
        let mut attributes = vec![];
        let is_numeric = !numbers.is_empty() && !has_string;

        if is_numeric && !has_blank {
            attributes.push(("containsSemiMixedTypes", "0".to_string()));
        }

        if is_numeric {
            attributes.push(("containsString", "0".to_string()));
        }

        if has_blank {
            attributes.push(("containsBlank", "1".to_string()));
        }

        if !numbers.is_empty() && has_string {
            attributes.push(("containsMixedTypes", "1".to_string()));
        }

        if !numbers.is_empty() {
            attributes.push(("containsNumber", "1".to_string()));

            if numbers.iter().all(|number| number.fract() == 0.0) {
                attributes.push(("containsInteger", "1".to_string()));
            }
        }

        if is_numeric {
            let min = numbers.iter().copied().fold(f64::INFINITY, f64::min);
            let max = numbers.iter().copied().fold(f64::NEG_INFINITY, f64::max);
            attributes.push(("minValue", min.to_string()));
            attributes.push(("maxValue", max.to_string()));
        }

        attributes.push(("count", values.len().to_string()));

        xml_start_tag(&mut self.writer, "sharedItems", &attributes);

        for value in values {
            match value {
                PivotFieldValue::Blank => xml_empty_tag_only(&mut self.writer, "m"),
                PivotFieldValue::Boolean(boolean) => {
                    let attributes = [("v", u8::from(*boolean).to_string())];
                    xml_empty_tag(&mut self.writer, "b", &attributes);
                }
                PivotFieldValue::Number(number) => {
                    let attributes = [("v", number.to_string())];
                    xml_empty_tag(&mut self.writer, "n", &attributes);
                }
                PivotFieldValue::String(string) => {
                    let attributes = [("v", escape_attribute_escapes(string).into_owned())];
                    xml_empty_tag(&mut self.writer, "s", &attributes);
                }
            }
        }

        xml_end_tag(&mut self.writer, "sharedItems");
    }
}

// -----------------------------------------------------------------------
// PivotFieldValue
// -----------------------------------------------------------------------

// A unique value of a pivot cache field. The variants map to the child
// elements of <sharedItems>.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum PivotFieldValue {
    Blank,
    Boolean(bool),
    Number(f64),
    String(String),
}

impl PivotFieldValue {
    // Sort the values of a field in ascending order and remove the duplicates.
    // Like Excel, strings are compared ignoring case and the first string of a
    // case-insensitive group is kept. The order of the types is numbers,
    // booleans, strings and then blanks.
    pub(crate) fn sort_and_dedup(values: &mut Vec<PivotFieldValue>) {
        values.sort_by(PivotFieldValue::compare);
        values.dedup_by(|value, other| value.compare(other) == Ordering::Equal);
    }

    fn compare(&self, other: &PivotFieldValue) -> Ordering {
        match (self, other) {
            (PivotFieldValue::Number(value), PivotFieldValue::Number(other)) => {
                value.total_cmp(other)
            }
            (PivotFieldValue::Boolean(value), PivotFieldValue::Boolean(other)) => value.cmp(other),
            (PivotFieldValue::String(value), PivotFieldValue::String(other)) => {
                value.to_lowercase().cmp(&other.to_lowercase())
            }
            _ => self.type_order().cmp(&other.type_order()),
        }
    }

    fn type_order(&self) -> u8 {
        match self {
            PivotFieldValue::Number(_) => 0,
            PivotFieldValue::Boolean(_) => 1,
            PivotFieldValue::String(_) => 2,
            PivotFieldValue::Blank => 3,
        }
    }
}
