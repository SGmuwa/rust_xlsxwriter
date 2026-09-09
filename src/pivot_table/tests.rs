// Pivot table unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod pivot_table_tests {

    use crate::pivot_table::PivotTable;
    use crate::test_functions::xml_to_vec;
    use crate::{
        xmlwriter, PivotTableDataField, PivotTableFunction, PivotTableLayout, PivotTableStyle,
        XlsxError,
    };
    use pretty_assertions::assert_eq;

    // The field names in the source data used by the tests below.
    fn field_names() -> Vec<String> {
        ["Region", "Item", "Volume", "Month"]
            .iter()
            .map(ToString::to_string)
            .collect()
    }

    #[test]
    fn test_unknown_field_name() {
        let mut pivot_table = PivotTable::new().add_row_field("Continent");

        let result = pivot_table.set_field_indices(&field_names());

        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));
    }

    #[test]
    fn test_cache_records_not_implemented() {
        let result = PivotTable::new().set_cache_records(true);
        assert!(matches!(result, Err(XlsxError::PivotTableError(_))));

        let result = PivotTable::new().set_cache_records(false);
        assert!(result.is_ok());
    }

    #[test]
    fn test_assemble1() {
        // Row, column, filter and data fields, as in tests/input/pivot_table01.xlsx.
        let mut pivot_table = PivotTable::new()
            .set_name("ItemsByRegion")
            .set_field_values(false)
            .set_extent(false)
            .set_style(PivotTableStyle::Medium9)
            .add_row_field("Region")
            .add_column_field("Item")
            .add_filter_field("Month")
            .add_data_field(PivotTableDataField::new("Volume"));

        pivot_table.index = 1;
        pivot_table.cache_id = 1;
        pivot_table.first_row = 1;
        pivot_table.first_col = 0;
        pivot_table.set_field_indices(&field_names()).unwrap();

        pivot_table.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&pivot_table.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="ItemsByRegion" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">
                    <location ref="A2" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="4">
                        <pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                        <pivotField axis="axisCol" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                        <pivotField dataField="1" showAll="0"/>
                        <pivotField axis="axisPage" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                    </pivotFields>
                    <rowFields count="1"><field x="0"/></rowFields>
                    <rowItems count="1"><i t="grand"><x/></i></rowItems>
                    <colFields count="1"><field x="1"/></colFields>
                    <colItems count="1"><i t="grand"><x/></i></colItems>
                    <pageFields count="1"><pageField fld="3" hier="-1"/></pageFields>
                    <dataFields count="1"><dataField name="Sum of Volume" fld="2"/></dataFields>
                    <pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="0"/>
                </pivotTableDefinition>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble2() {
        // A row field with a non-default function and a user defined name, as in
        // the third pivot table of tests/input/pivot_table02.xlsx.
        let mut pivot_table = PivotTable::new()
            .set_name("AverageVolumeByRegion")
            .set_field_values(false)
            .set_extent(false)
            .set_style(PivotTableStyle::Dark1)
            .add_row_field("Region")
            .add_data_field(
                PivotTableDataField::new("Volume")
                    .set_function(PivotTableFunction::Average)
                    .set_name("Average Volume"),
            );

        pivot_table.index = 3;
        pivot_table.cache_id = 1;
        pivot_table.first_row = 10;
        pivot_table.first_col = 0;
        pivot_table.set_field_indices(&field_names()).unwrap();

        pivot_table.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&pivot_table.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="AverageVolumeByRegion" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">
                    <location ref="A11" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="4">
                        <pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                        <pivotField showAll="0"/>
                        <pivotField dataField="1" showAll="0"/>
                        <pivotField showAll="0"/>
                    </pivotFields>
                    <rowFields count="1"><field x="0"/></rowFields>
                    <rowItems count="1"><i t="grand"><x/></i></rowItems>
                    <dataFields count="1"><dataField name="Average Volume" fld="2" subtotal="average"/></dataFields>
                    <pivotTableStyleInfo name="PivotStyleDark1" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="0"/>
                </pivotTableDefinition>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble_tabular_layout_and_num_format() {
        // Tabular layout and a number format on the data field.
        let mut pivot_table = PivotTable::new()
            .set_layout(PivotTableLayout::Tabular)
            .add_row_field("Region")
            .add_row_field("Item")
            .add_data_field(PivotTableDataField::new("Volume").set_num_format("0.00"));

        pivot_table.index = 1;
        pivot_table.cache_id = 1;
        pivot_table.set_field_indices(&field_names()).unwrap();
        pivot_table.data_fields[0].num_format_index = 164;

        pivot_table.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&pivot_table.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" compact="0" outline="0" outlineData="0" compactData="0" multipleFieldFilters="0">
                    <location ref="A1:C2" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="4">
                        <pivotField axis="axisRow" compact="0" outline="0" subtotalTop="0" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                        <pivotField axis="axisRow" compact="0" outline="0" subtotalTop="0" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                        <pivotField dataField="1" showAll="0"/>
                        <pivotField showAll="0"/>
                    </pivotFields>
                    <rowFields count="2"><field x="0"/><field x="1"/></rowFields>
                    <rowItems count="1"><i t="grand"><x/></i></rowItems>
                    <dataFields count="1"><dataField name="Sum of Volume" fld="2" numFmtId="164"/></dataFields>
                    <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="0"/>
                </pivotTableDefinition>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble_subtotal_caption() {
        // A subtotal caption on the row and column fields. The filter field
        // has no subtotals, so it doesn't get the caption.
        let mut pivot_table = PivotTable::new()
            .set_subtotal_caption("Subtotal")
            .add_row_field("Region")
            .add_column_field("Item")
            .add_filter_field("Month")
            .add_data_field(PivotTableDataField::new("Volume"));

        pivot_table.index = 1;
        pivot_table.cache_id = 1;
        pivot_table.set_field_indices(&field_names()).unwrap();

        pivot_table.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&pivot_table.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">
                    <location ref="A2:C4" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="4">
                        <pivotField axis="axisRow" subtotalCaption="Subtotal" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                        <pivotField axis="axisCol" subtotalCaption="Subtotal" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                        <pivotField dataField="1" showAll="0"/>
                        <pivotField axis="axisPage" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                    </pivotFields>
                    <rowFields count="1"><field x="0"/></rowFields>
                    <rowItems count="1"><i t="grand"><x/></i></rowItems>
                    <colFields count="1"><field x="1"/></colFields>
                    <colItems count="1"><i t="grand"><x/></i></colItems>
                    <pageFields count="1"><pageField fld="3" hier="-1"/></pageFields>
                    <dataFields count="1"><dataField name="Sum of Volume" fld="2"/></dataFields>
                    <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="0"/>
                </pivotTableDefinition>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble_no_subtotals() {
        // The subtotals of one row field turned off, and those of another
        // turned off and back on again.
        let mut pivot_table = PivotTable::new()
            .add_row_field("Region")
            .add_row_field("Item")
            .set_show_field_subtotals("Item", false)
            .set_show_field_subtotals("Region", false)
            .set_show_field_subtotals("Region", true)
            .add_data_field(PivotTableDataField::new("Volume"));

        pivot_table.index = 1;
        pivot_table.cache_id = 1;
        pivot_table.set_field_indices(&field_names()).unwrap();

        pivot_table.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&pivot_table.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">
                    <location ref="A1:B2" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="4">
                        <pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                        <pivotField axis="axisRow" showAll="0" defaultSubtotal="0"><items count="1"><item t="default"/></items></pivotField>
                        <pivotField dataField="1" showAll="0"/>
                        <pivotField showAll="0"/>
                    </pivotFields>
                    <rowFields count="2"><field x="0"/><field x="1"/></rowFields>
                    <rowItems count="1"><i t="grand"><x/></i></rowItems>
                    <dataFields count="1"><dataField name="Sum of Volume" fld="2"/></dataFields>
                    <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="0"/>
                </pivotTableDefinition>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble3() {
        // Grand totals turned off and a default pivot table name.
        let mut pivot_table = PivotTable::new()
            .set_show_row_grand_total(false)
            .set_show_column_grand_total(false)
            .add_row_field("Region")
            .add_data_field(PivotTableDataField::new("Volume"));

        pivot_table.index = 1;
        pivot_table.cache_id = 1;
        pivot_table.set_field_indices(&field_names()).unwrap();

        pivot_table.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&pivot_table.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" rowGrandTotals="0" colGrandTotals="0" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">
                    <location ref="A1:B2" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="4">
                        <pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField>
                        <pivotField showAll="0"/>
                        <pivotField dataField="1" showAll="0"/>
                        <pivotField showAll="0"/>
                    </pivotFields>
                    <rowFields count="1"><field x="0"/></rowFields>
                    <rowItems count="1"><i t="grand"><x/></i></rowItems>
                    <dataFields count="1"><dataField name="Sum of Volume" fld="2"/></dataFields>
                    <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="0"/>
                </pivotTableDefinition>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble_field_values_and_extent() {
        // The default output: the values of the fields used on an axis and an
        // estimate of the extent of the pivot table.
        let mut pivot_table = PivotTable::new()
            .set_layout(PivotTableLayout::Tabular)
            .add_row_field("Region")
            .add_row_field("Item")
            .set_show_field_subtotals("Item", false)
            .add_column_field("Month")
            .add_data_field(PivotTableDataField::new("Volume"));

        pivot_table.index = 1;
        pivot_table.cache_id = 1;
        pivot_table.record_count = 10;
        pivot_table.set_field_indices(&field_names()).unwrap();

        // The number of values of each field, as read from the source data.
        pivot_table.field_item_counts = vec![3, 2, 0, 4];

        pivot_table.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&pivot_table.writer);
        let got = xml_to_vec(got);

        // The extent allows a row per source record, a subtotal row per record
        // for the one row field that has subtotals, a header row per column
        // field and a grand total row.
        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" compact="0" outline="0" outlineData="0" compactData="0" multipleFieldFilters="0">
                    <location ref="A1:G23" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="4">
                        <pivotField axis="axisRow" compact="0" outline="0" subtotalTop="0" showAll="0"><items count="4"><item x="0"/><item x="1"/><item x="2"/><item t="default"/></items></pivotField>
                        <pivotField axis="axisRow" compact="0" outline="0" subtotalTop="0" showAll="0" defaultSubtotal="0"><items count="2"><item x="0"/><item x="1"/></items></pivotField>
                        <pivotField dataField="1" showAll="0"/>
                        <pivotField axis="axisCol" showAll="0"><items count="5"><item x="0"/><item x="1"/><item x="2"/><item x="3"/><item t="default"/></items></pivotField>
                    </pivotFields>
                    <rowFields count="2"><field x="0"/><field x="1"/></rowFields>
                    <rowItems count="1"><i t="grand"><x/></i></rowItems>
                    <colFields count="1"><field x="3"/></colFields>
                    <colItems count="1"><i t="grand"><x/></i></colItems>
                    <dataFields count="1"><dataField name="Sum of Volume" fld="2"/></dataFields>
                    <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="0"/>
                </pivotTableDefinition>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble_extent_rows() {
        // A caller supplied number of rows replaces the estimate.
        let mut pivot_table = PivotTable::new()
            .set_extent_rows(4)
            .add_row_field("Region")
            .add_data_field(PivotTableDataField::new("Volume"));

        pivot_table.index = 1;
        pivot_table.cache_id = 1;
        pivot_table.record_count = 100;
        pivot_table.first_row = 1;
        pivot_table.set_field_indices(&field_names()).unwrap();
        pivot_table.field_item_counts = vec![2, 0, 0, 0];

        pivot_table.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&pivot_table.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">
                    <location ref="A2:B5" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="4">
                        <pivotField axis="axisRow" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField>
                        <pivotField showAll="0"/>
                        <pivotField dataField="1" showAll="0"/>
                        <pivotField showAll="0"/>
                    </pivotFields>
                    <rowFields count="1"><field x="0"/></rowFields>
                    <rowItems count="1"><i t="grand"><x/></i></rowItems>
                    <dataFields count="1"><dataField name="Sum of Volume" fld="2"/></dataFields>
                    <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="0"/>
                </pivotTableDefinition>
            "#,
        );

        assert_eq!(expected, got);
    }
}
