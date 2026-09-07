// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! Example of adding a pivot table to a worksheet.

use rust_xlsxwriter::{PivotTable, PivotTableDataField, PivotTableFunction, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet with the source data.
    let worksheet = workbook.add_worksheet().set_name("Data")?;
    worksheet.write_row(0, 0, ["Region", "Item", "Volume"])?;
    worksheet.write_row(1, 0, ["East", "Apple"])?;
    worksheet.write_row(2, 0, ["West", "Apple"])?;
    worksheet.write_row(3, 0, ["East", "Pear"])?;
    worksheet.write_column(1, 2, [9000, 5000, 7000])?;

    // Create a pivot table of the volume per region.
    let pivot_table = PivotTable::new()
        .set_data_source(("Data", 0, 0, 3, 2))
        .add_row_field("Region")
        .add_data_field(PivotTableDataField::new("Volume").set_function(PivotTableFunction::Sum));

    // Add the pivot table to a second worksheet.
    let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    worksheet.add_pivot_table(0, 0, &pivot_table)?;

    workbook.save("pivot_table.xlsx")?;

    Ok(())
}
