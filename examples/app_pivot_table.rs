// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of creating pivot tables with the `rust_xlsxwriter` library.
//!
//! Note, the cells of a pivot table are only populated once the file has been
//! opened by Excel or `LibreOffice`, which recalculate the summary from the
//! source data.

use rust_xlsxwriter::{
    Format, PivotTable, PivotTableDataField, PivotTableFunction, PivotTableLayout, PivotTableStyle,
    Workbook, Worksheet, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet with the source data for the pivot tables.
    let worksheet = workbook.add_worksheet().set_name("Data")?;
    write_source_data(worksheet)?;

    // The source range is the same for all the pivot tables below so they share
    // a single pivot cache in the output file.
    let source = ("Data", 0, 0, 12, 3);

    // -----------------------------------------------------------------------
    // 1. A simple pivot table with one row field and one data field.
    // -----------------------------------------------------------------------
    let pivot_table = PivotTable::new()
        .set_name("VolumeByRegion")
        .set_data_source(source)
        .add_row_field("Region")
        .add_data_field(PivotTableDataField::new("Volume"));

    let worksheet = workbook.add_worksheet().set_name("Volume by region")?;
    worksheet.set_column_width(0, 16)?;
    worksheet.add_pivot_table(0, 0, &pivot_table)?;

    // -----------------------------------------------------------------------
    // 2. A pivot table with a column field and a filter field.
    // -----------------------------------------------------------------------
    let pivot_table = PivotTable::new()
        .set_name("VolumeByRegionAndItem")
        .set_data_source(source)
        .set_style(PivotTableStyle::Medium9)
        .add_filter_field("Month")
        .add_row_field("Region")
        .add_column_field("Item")
        .add_data_field(PivotTableDataField::new("Volume"));

    let worksheet = workbook.add_worksheet().set_name("Volume by item")?;
    worksheet.set_column_width(0, 16)?;
    worksheet.add_pivot_table(2, 0, &pivot_table)?;

    // -----------------------------------------------------------------------
    // 3. A pivot table in tabular layout with a named data field, a summary
    //    function other than Sum and a number format.
    // -----------------------------------------------------------------------
    let pivot_table = PivotTable::new()
        .set_name("AverageVolume")
        .set_data_source(source)
        .set_style(PivotTableStyle::Dark1)
        .set_layout(PivotTableLayout::Tabular)
        .set_show_column_grand_total(false)
        .add_row_field("Region")
        .add_row_field("Item")
        .add_data_field(
            PivotTableDataField::new("Volume")
                .set_function(PivotTableFunction::Average)
                .set_name("Average volume")
                .set_num_format("#,##0.00"),
        );

    let worksheet = workbook.add_worksheet().set_name("Average volume")?;
    worksheet.set_column_range_width(0, 2, 16)?;
    worksheet.add_pivot_table(0, 0, &pivot_table)?;

    workbook.save("pivot_table.xlsx")?;

    Ok(())
}

// Write the source data used by the pivot tables.
fn write_source_data(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    let data = [
        ("East", "Apple", 9000, "July"),
        ("East", "Apple", 5000, "April"),
        ("South", "Orange", 9000, "September"),
        ("North", "Apple", 2000, "November"),
        ("West", "Apple", 9000, "November"),
        ("South", "Pear", 7000, "October"),
        ("North", "Pear", 9000, "August"),
        ("West", "Orange", 1000, "December"),
        ("West", "Grape", 1000, "November"),
        ("South", "Pear", 10000, "April"),
        ("West", "Grape", 6000, "January"),
        ("South", "Orange", 3000, "May"),
    ];

    worksheet.set_column_range_width(0, 3, 12)?;

    let header_format = Format::new().set_bold();
    worksheet.write_row_with_format(0, 0, ["Region", "Item", "Volume", "Month"], &header_format)?;

    for (row, data) in data.iter().enumerate() {
        let row = 1 + row as u32;
        worksheet.write_string(row, 0, data.0)?;
        worksheet.write_string(row, 1, data.1)?;
        worksheet.write_number(row, 2, data.2)?;
        worksheet.write_string(row, 3, data.3)?;
    }

    Ok(())
}
