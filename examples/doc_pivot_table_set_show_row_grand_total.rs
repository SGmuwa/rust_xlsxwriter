// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! Example of turning off the row grand total of a pivot table.

use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet with the source data for the pivot table.
    let worksheet = workbook.add_worksheet().set_name("Data")?;
    worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    worksheet.write_column(1, 3, [9000, 5000, 7000])?;

    // Create a pivot table without the row grand total.
    let pivot_table = PivotTable::new()
        .set_data_source(("Data", 0, 0, 3, 3))
        .set_show_row_grand_total(false)
        .add_row_field("Region")
        .add_column_field("Item")
        .add_data_field(PivotTableDataField::new("Volume"));

    // Add the pivot table to a new worksheet.
    let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    worksheet.add_pivot_table(0, 0, &pivot_table)?;

    // Save the file to disk.
    workbook.save("pivot_table.xlsx")?;

    Ok(())
}
