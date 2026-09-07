// pivot_cache - A module for creating the Excel pivotCacheDefinition.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

mod tests;

use std::io::Cursor;

use crate::utility;
use crate::xmlwriter::{
    xml_declaration, xml_empty_tag, xml_empty_tag_only, xml_end_tag, xml_start_tag,
};
use crate::{ColNum, RowNum};

// A pivot cache definition holds the metadata for the source data of one or
// more pivot tables: the source range and the names of the source fields.
//
// The associated pivot cache records file, which is a row by row copy of the
// source data, isn't written. Instead the definition is marked with
// `saveData="0"` and `refreshOnLoad="1"` so that the application that opens
// the file recreates the cache, and the pivot table data, from the source
// range.
pub(crate) struct PivotCache {
    pub(crate) writer: Cursor<Vec<u8>>,

    pub(crate) sheet_name: String,
    pub(crate) first_row: RowNum,
    pub(crate) first_col: ColNum,
    pub(crate) last_row: RowNum,
    pub(crate) last_col: ColNum,
    pub(crate) field_names: Vec<String>,
}

impl PivotCache {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new PivotCache struct.
    pub(crate) fn new(
        sheet_name: &str,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> PivotCache {
        let writer = Cursor::new(Vec::with_capacity(2048));

        PivotCache {
            writer,
            sheet_name: sheet_name.to_string(),
            first_row,
            first_col,
            last_row,
            last_col,
            field_names: vec![],
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and generate the XML file.
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
        let xmlns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // The number of source rows, excluding the header row.
        let record_count = self.last_row - self.first_row;

        let attributes = [
            ("xmlns", xmlns.to_string()),
            ("xmlns:r", xmlns_r.to_string()),
            // The cache records are omitted from the file, see the struct docs.
            ("saveData", "0".to_string()),
            ("refreshOnLoad", "1".to_string()),
            ("refreshedBy", "rust_xlsxwriter".to_string()),
            ("refreshedDate", "0".to_string()),
            ("createdVersion", "6".to_string()),
            ("refreshedVersion", "6".to_string()),
            ("minRefreshableVersion", "3".to_string()),
            ("recordCount", record_count.to_string()),
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
        let range =
            utility::cell_range(self.first_row, self.first_col, self.last_row, self.last_col);

        let attributes = [("ref", range), ("sheet", self.sheet_name.clone())];

        xml_empty_tag(&mut self.writer, "worksheetSource", &attributes);
    }

    // Write the <cacheFields> element.
    fn write_cache_fields(&mut self) {
        let attributes = [("count", self.field_names.len().to_string())];

        xml_start_tag(&mut self.writer, "cacheFields", &attributes);

        for name in self.field_names.clone() {
            // Write the cacheField element.
            self.write_cache_field(&name);
        }

        xml_end_tag(&mut self.writer, "cacheFields");
    }

    // Write the <cacheField> element.
    fn write_cache_field(&mut self, name: &str) {
        let attributes = [("name", name), ("numFmtId", "0")];

        xml_start_tag(&mut self.writer, "cacheField", &attributes);

        // The shared items are left empty since they are recreated by the
        // application that opens the file, see the struct docs.
        xml_empty_tag_only(&mut self.writer, "sharedItems");

        xml_end_tag(&mut self.writer, "cacheField");
    }
}
