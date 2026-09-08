// pivot_table - A module for creating the Excel pivotTable.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

use std::fmt;
use std::io::Cursor;

use crate::xmlwriter::{
    xml_declaration, xml_empty_tag, xml_empty_tag_only, xml_end_tag, xml_start_tag,
};
use crate::{utility, ChartRange, ColNum, IntoChartRange, RowNum, XlsxError};

/// The `PivotTable` struct represents a worksheet pivot table.
///
/// A pivot table is a data summarization tool that aggregates the rows of a
/// range of source data into a report with rows, columns, filters and
/// aggregated values.
///
/// A pivot table is added to a worksheet via the
/// [`Worksheet::add_pivot_table()`](crate::Worksheet::add_pivot_table) method.
/// The source data can be on the same worksheet or on any other worksheet in
/// the workbook.
///
/// ```
/// # // This code is available in examples/doc_pivot_table_intro.rs
/// #
/// # use rust_xlsxwriter::{
/// #     PivotTable, PivotTableDataField, PivotTableFunction, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with the source data.
/// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
/// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
/// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
/// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
/// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
/// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
/// #
///     // Create a pivot table of the volume per region.
///     let pivot_table = PivotTable::new()
///         .set_data_source(("Data", 0, 0, 3, 3))
///         .add_row_field("Region")
///         .add_data_field(PivotTableDataField::new("Volume").set_function(PivotTableFunction::Sum));
///
///     // Add the pivot table to a second worksheet.
///     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
///     worksheet.add_pivot_table(0, 0, &pivot_table)?;
/// #
/// #     workbook.save("pivot_table.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// # Note on refreshing the pivot table data
///
/// `rust_xlsxwriter` writes the definition of the pivot table but not the
/// summarized data itself. The file is flagged so that the application that
/// opens it, such as Excel or `LibreOffice`, recalculates the summary from the
/// source data when the file is loaded. As a result the cells of the pivot
/// table are only populated after the file has been opened once.
///
/// For more information on pivot tables see the Microsoft documentation on
/// [Create a PivotTable to analyze worksheet data].
///
/// [Create a PivotTable to analyze worksheet data]:
///     https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576
///
#[derive(Clone)]
pub struct PivotTable {
    pub(crate) writer: Cursor<Vec<u8>>,

    pub(crate) name: String,
    pub(crate) data_source: ChartRange,
    pub(crate) style: PivotTableStyle,
    pub(crate) layout: PivotTableLayout,
    pub(crate) show_row_grand_totals: bool,
    pub(crate) show_column_grand_totals: bool,
    pub(crate) subtotal_caption: String,

    pub(crate) row_field_names: Vec<String>,
    pub(crate) column_field_names: Vec<String>,
    pub(crate) filter_field_names: Vec<String>,
    pub(crate) data_fields: Vec<PivotTableDataField>,
    pub(crate) no_subtotal_field_names: Vec<String>,

    // The following properties are set in Workbook::prepare_pivot_tables()
    // once the source data can be read from the workbook.
    pub(crate) index: u32,
    pub(crate) cache_id: u32,
    pub(crate) first_row: RowNum,
    pub(crate) first_col: ColNum,
    pub(crate) num_fields: usize,
    pub(crate) row_fields: Vec<usize>,
    pub(crate) column_fields: Vec<usize>,
    pub(crate) filter_fields: Vec<usize>,
    pub(crate) no_subtotal_fields: Vec<usize>,
}

impl Default for PivotTable {
    fn default() -> Self {
        Self::new()
    }
}

impl PivotTable {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new `PivotTable` struct instance.
    ///
    /// A new pivot table needs, at a minimum, a data source (see
    /// [`PivotTable::set_data_source()`]) and one data field (see
    /// [`PivotTable::add_data_field()`]).
    ///
    /// # Examples
    ///
    /// Example of creating a new pivot table and adding it to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_new.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a new pivot table.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_row_field("Region")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn new() -> PivotTable {
        let writer = Cursor::new(Vec::with_capacity(2048));

        PivotTable {
            writer,
            name: String::new(),
            data_source: ChartRange::default(),
            style: PivotTableStyle::Light16,
            layout: PivotTableLayout::Compact,
            show_row_grand_totals: true,
            show_column_grand_totals: true,
            subtotal_caption: String::new(),
            row_field_names: vec![],
            column_field_names: vec![],
            filter_field_names: vec![],
            data_fields: vec![],
            no_subtotal_field_names: vec![],
            index: 0,
            cache_id: 0,
            first_row: 0,
            first_col: 0,
            num_fields: 0,
            row_fields: vec![],
            column_fields: vec![],
            filter_fields: vec![],
            no_subtotal_fields: vec![],
        }
    }

    /// Set the name of the pivot table.
    ///
    /// Excel requires that pivot table names are unique within a workbook and
    /// that they follow the same rules as defined names, see
    /// [`utility::check_name()`](crate::utility::check_name).
    ///
    /// If a name isn't set then a default name like `PivotTable1` is used.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the pivot table.
    ///
    /// # Examples
    ///
    /// Example of setting the name of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table and set its name.
    ///     let pivot_table = PivotTable::new()
    ///         .set_name("VolumeByRegion")
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_row_field("Region")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_name(mut self, name: impl Into<String>) -> PivotTable {
        self.name = name.into();
        self
    }

    /// Set the range of the source data for the pivot table.
    ///
    /// The source data must contain a header row since the pivot table fields
    /// are referred to by the header names. The data doesn't have to be on the
    /// same worksheet as the pivot table.
    ///
    /// # Parameters
    ///
    /// - `range`: The range of the source data. This can be a string like
    ///   `"Data!$A$1:$D$51"` or a 5-tuple of a worksheet name and zero indexed
    ///   row/column values like `("Data", 0, 0, 50, 3)`. See
    ///   [`IntoChartRange`].
    ///
    /// # Examples
    ///
    /// Example of setting the source data range of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_data_source.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table and set the range of its source data. The range
    ///     // could also be given as a string like "Data!$A$1:$D$4".
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_row_field("Region")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_data_source<T>(mut self, range: T) -> PivotTable
    where
        T: IntoChartRange,
    {
        self.data_source = range.new_chart_range();
        self
    }

    /// Set the style of the pivot table.
    ///
    /// The default style is [`PivotTableStyle::Light16`], which is also Excel's
    /// default.
    ///
    /// # Parameters
    ///
    /// - `style`: A [`PivotTableStyle`] enum value.
    ///
    /// # Examples
    ///
    /// Example of setting the style of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_style.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     PivotTable, PivotTableDataField, PivotTableStyle, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table and set its style.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .set_style(PivotTableStyle::Medium10)
    ///         .add_row_field("Region")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_style(mut self, style: PivotTableStyle) -> PivotTable {
        self.style = style;
        self
    }

    /// Set the layout of the row fields of the pivot table.
    ///
    /// Excel has three layouts for the row area of a pivot table: "Compact",
    /// "Outline" and "Tabular", see [`PivotTableLayout`]. The default, like
    /// Excel, is [`PivotTableLayout::Compact`], where all the row fields share
    /// a single column.
    ///
    /// # Parameters
    ///
    /// - `layout`: A [`PivotTableLayout`] enum value.
    ///
    /// # Examples
    ///
    /// Example of setting the row layout of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_layout.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     PivotTable, PivotTableDataField, PivotTableLayout, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table with two row fields in tabular layout so that each
    ///     // row field gets a column of its own.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .set_layout(PivotTableLayout::Tabular)
    ///         .add_row_field("Region")
    ///         .add_row_field("Item")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_layout(mut self, layout: PivotTableLayout) -> PivotTable {
        self.layout = layout;
        self
    }

    /// Turn the row grand totals on or off.
    ///
    /// Row grand totals are on by default.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    /// # Examples
    ///
    /// Example of turning off the row grand total of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_show_row_grand_total.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table without the row grand total.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .set_show_row_grand_total(false)
    ///         .add_row_field("Region")
    ///         .add_column_field("Item")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_show_row_grand_total(mut self, enable: bool) -> PivotTable {
        self.show_row_grand_totals = enable;
        self
    }

    /// Turn the column grand totals on or off.
    ///
    /// Column grand totals are on by default.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    /// # Examples
    ///
    /// Example of turning off the column grand total of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_show_column_grand_total.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table without the column grand total.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .set_show_column_grand_total(false)
    ///         .add_row_field("Region")
    ///         .add_column_field("Item")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_show_column_grand_total(mut self, enable: bool) -> PivotTable {
        self.show_column_grand_totals = enable;
        self
    }

    /// Set the caption used in the subtotal rows and columns.
    ///
    /// Excel labels a subtotal with the name of the item it belongs to
    /// followed by a word from the language of the application, such as `East
    /// Total`. This method replaces that word, which is required for
    /// non-English workbooks and for reports that use a term of their own,
    /// such as `East Subtotal`.
    ///
    /// The caption applies to the fields in the row and column areas. The item
    /// name is always written first, by the application, and its position
    /// cannot be changed.
    ///
    /// # Parameters
    ///
    /// - `caption`: The word used in the subtotal rows and columns.
    ///
    /// # Examples
    ///
    /// Example of setting the subtotal caption of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_subtotal_caption.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, PivotTableLayout, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table whose subtotal rows read "East Subtotal"
    ///     // instead of the default "East Total".
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .set_layout(PivotTableLayout::Tabular)
    ///         .set_subtotal_caption("Subtotal")
    ///         .add_row_field("Region")
    ///         .add_row_field("Item")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_subtotal_caption(mut self, caption: impl Into<String>) -> PivotTable {
        self.subtotal_caption = caption.into();
        self
    }

    /// Turn the subtotals of a single field on or off.
    ///
    /// A field in the row or column area gets a subtotal row or column of its
    /// own, unless it is the innermost field of its area. This method turns
    /// those subtotals off for one field, like the `Subtotals > None` option in
    /// the Excel field settings.
    ///
    /// The usual reason to turn them off is a field whose items are unique,
    /// such as a date or an id: its subtotal repeats the row it belongs to and
    /// doubles the length of the report.
    ///
    /// Subtotals of the other fields, and the grand totals, are not affected.
    /// See [`PivotTable::set_show_row_grand_total()`] for the latter.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of a field in the header row of the source data.
    /// - `enable`: Turn the subtotals of the field on or off. They are on by
    ///   default.
    ///
    /// # Examples
    ///
    /// Example of turning off the subtotals of one field of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_show_field_subtotals.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, PivotTableLayout, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table where the items keep their subtotal rows but
    ///     // the months, one per row, do not.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .set_layout(PivotTableLayout::Tabular)
    ///         .add_row_field("Region")
    ///         .add_row_field("Item")
    ///         .add_row_field("Month")
    ///         .set_show_field_subtotals("Item", true)
    ///         .set_show_field_subtotals("Month", false)
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_show_field_subtotals(mut self, name: impl Into<String>, enable: bool) -> PivotTable {
        let name = name.into();
        self.no_subtotal_field_names.retain(|field| *field != name);

        if !enable {
            self.no_subtotal_field_names.push(name);
        }

        self
    }

    /// Add a field to the row area of the pivot table.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of a field in the header row of the source data.
    ///
    /// # Examples
    ///
    /// Example of adding a field to the row area of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_add_row_field.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table with the regions in the row area.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_row_field("Region")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn add_row_field(mut self, name: impl Into<String>) -> PivotTable {
        self.row_field_names.push(name.into());
        self
    }

    /// Add a field to the column area of the pivot table.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of a field in the header row of the source data.
    ///
    /// # Examples
    ///
    /// Example of adding a field to the column area of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_add_column_field.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table with the items in the column area.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_row_field("Region")
    ///         .add_column_field("Item")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn add_column_field(mut self, name: impl Into<String>) -> PivotTable {
        self.column_field_names.push(name.into());
        self
    }

    /// Add a field to the filter (page) area of the pivot table.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of a field in the header row of the source data.
    ///
    /// # Examples
    ///
    /// Example of adding a field to the filter area of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_add_filter_field.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table with the months in the filter area.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_filter_field("Month")
    ///         .add_row_field("Region")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn add_filter_field(mut self, name: impl Into<String>) -> PivotTable {
        self.filter_field_names.push(name.into());
        self
    }

    /// Add a field to the values area of the pivot table.
    ///
    /// A pivot table requires at least one data field.
    ///
    /// # Parameters
    ///
    /// - `data_field`: A [`PivotTableDataField`] struct reference.
    ///
    /// # Examples
    ///
    /// Example of adding a field to the values area of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_add_data_field.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     PivotTable, PivotTableDataField, PivotTableFunction, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table with the average volume in the values area.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_row_field("Region")
    ///         .add_data_field(
    ///             PivotTableDataField::new("Volume").set_function(PivotTableFunction::Average),
    ///         );
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn add_data_field(mut self, data_field: impl Into<PivotTableDataField>) -> PivotTable {
        self.data_fields.push(data_field.into());
        self
    }

    /// Write the pivot cache records to the file.
    ///
    /// Excel stores a copy of the pivot table source data in the file, in a
    /// "pivot cache records" part, so that the pivot table can be displayed
    /// without recalculating it from the source data.
    ///
    /// `rust_xlsxwriter` doesn't currently write this part. Instead the file is
    /// flagged so that the application that opens it recreates the cache, and
    /// the pivot table values, from the source data.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::PivotTableError`] - This method isn't implemented yet and
    ///   returns an error if `enable` is set to `true`.
    ///
    /// # Examples
    ///
    /// Example of turning the pivot cache records off, which is the default.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_cache_data.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a pivot table. Writing the pivot cache records isn't implemented
    ///     // so they can only be turned off, which is also the default.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .set_cache_data(false)?
    ///         .add_row_field("Region")
    ///         .add_data_field(PivotTableDataField::new("Volume"));
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_cache_data(self, enable: bool) -> Result<PivotTable, XlsxError> {
        if enable {
            return Err(XlsxError::PivotTableError(
                "Writing the pivot cache records is not implemented".to_string(),
            ));
        }

        Ok(self)
    }

    // Resolve the field names used by the pivot table to their indices in the
    // header row of the source data.
    pub(crate) fn set_field_indices(&mut self, field_names: &[String]) -> Result<(), XlsxError> {
        self.num_fields = field_names.len();

        self.row_fields = Self::field_indices(&self.row_field_names, field_names)?;
        self.column_fields = Self::field_indices(&self.column_field_names, field_names)?;
        self.filter_fields = Self::field_indices(&self.filter_field_names, field_names)?;
        self.no_subtotal_fields = Self::field_indices(&self.no_subtotal_field_names, field_names)?;

        for data_field in &mut self.data_fields {
            data_field.field_index = Self::field_index(&data_field.field_name, field_names)?;
        }

        Ok(())
    }

    // Resolve a list of field names to their indices in the source data.
    fn field_indices(names: &[String], field_names: &[String]) -> Result<Vec<usize>, XlsxError> {
        names
            .iter()
            .map(|name| Self::field_index(name, field_names))
            .collect()
    }

    // Resolve a field name to its index in the header row of the source data.
    fn field_index(name: &str, field_names: &[String]) -> Result<usize, XlsxError> {
        match field_names.iter().position(|field| field == name) {
            Some(index) => Ok(index),
            None => Err(XlsxError::PivotTableError(format!(
                "Unknown field name '{name}' in the pivot table source data"
            ))),
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and generate the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the pivotTableDefinition element.
        self.write_pivot_table_definition();

        // Write the location element.
        self.write_location();

        // Write the pivotFields element.
        self.write_pivot_fields();

        if !self.row_fields.is_empty() {
            // Write the rowFields and rowItems elements.
            self.write_axis_fields("rowFields", &self.row_fields.clone());
            self.write_axis_items("rowItems");
        }

        if !self.column_fields.is_empty() {
            // Write the colFields and colItems elements.
            self.write_axis_fields("colFields", &self.column_fields.clone());
            self.write_axis_items("colItems");
        }

        if !self.filter_fields.is_empty() {
            // Write the pageFields element.
            self.write_page_fields();
        }

        if !self.data_fields.is_empty() {
            // Write the dataFields element.
            self.write_data_fields();
        }

        // Write the pivotTableStyleInfo element.
        self.write_pivot_table_style_info();

        // Close the pivotTableDefinition tag.
        xml_end_tag(&mut self.writer, "pivotTableDefinition");
    }

    // Write the <pivotTableDefinition> element.
    fn write_pivot_table_definition(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let name = if self.name.is_empty() {
            format!("PivotTable{}", self.index)
        } else {
            self.name.clone()
        };

        let mut attributes = vec![
            ("xmlns", xmlns.to_string()),
            ("name", name),
            ("cacheId", self.cache_id.to_string()),
            ("applyNumberFormats", "0".to_string()),
            ("applyBorderFormats", "0".to_string()),
            ("applyFontFormats", "0".to_string()),
            ("applyPatternFormats", "0".to_string()),
            ("applyAlignmentFormats", "0".to_string()),
            ("applyWidthHeightFormats", "1".to_string()),
            ("dataCaption", "Values".to_string()),
            ("updatedVersion", "6".to_string()),
            ("minRefreshableVersion", "3".to_string()),
            ("useAutoFormatting", "1".to_string()),
        ];

        if !self.show_row_grand_totals {
            attributes.push(("rowGrandTotals", "0".to_string()));
        }

        if !self.show_column_grand_totals {
            attributes.push(("colGrandTotals", "0".to_string()));
        }

        attributes.extend([
            ("itemPrintTitles", "1".to_string()),
            ("createdVersion", "6".to_string()),
            ("indent", "0".to_string()),
        ]);

        // The compact/outline attributes default to "1" in Excel so only the
        // non-default values are written out.
        if self.layout != PivotTableLayout::Compact {
            attributes.push(("compact", "0".to_string()));
        }

        let outline = u8::from(self.layout != PivotTableLayout::Tabular);
        attributes.extend([
            ("outline", outline.to_string()),
            ("outlineData", outline.to_string()),
        ]);

        if self.layout != PivotTableLayout::Compact {
            attributes.push(("compactData", "0".to_string()));
        }

        attributes.push(("multipleFieldFilters", "0".to_string()));

        xml_start_tag(&mut self.writer, "pivotTableDefinition", &attributes);
    }

    // Write the <location> element.
    fn write_location(&mut self) {
        let cell = utility::row_col_to_cell(self.first_row, self.first_col);

        let attributes = [
            ("ref", cell),
            ("firstHeaderRow", "1".to_string()),
            ("firstDataRow", "2".to_string()),
            ("firstDataCol", "1".to_string()),
        ];

        xml_empty_tag(&mut self.writer, "location", &attributes);
    }

    // Write the <pivotFields> element.
    fn write_pivot_fields(&mut self) {
        let attributes = [("count", self.num_fields.to_string())];

        xml_start_tag(&mut self.writer, "pivotFields", &attributes);

        for index in 0..self.num_fields {
            // Write the pivotField element.
            self.write_pivot_field(index);
        }

        xml_end_tag(&mut self.writer, "pivotFields");
    }

    // Write the <pivotField> element.
    fn write_pivot_field(&mut self, index: usize) {
        let axis = if self.row_fields.contains(&index) {
            Some("axisRow")
        } else if self.column_fields.contains(&index) {
            Some("axisCol")
        } else if self.filter_fields.contains(&index) {
            Some("axisPage")
        } else {
            None
        };

        match axis {
            Some(axis) => {
                let mut attributes = vec![("axis", axis)];

                // The subtotal caption only applies to the areas that have
                // subtotals, which are the row and column areas.
                let subtotal_caption = self.subtotal_caption.clone();
                if !subtotal_caption.is_empty() && axis != "axisPage" {
                    attributes.push(("subtotalCaption", subtotal_caption.as_str()));
                }

                // The layout of a pivot field only applies to the row area.
                if axis == "axisRow" && self.layout != PivotTableLayout::Compact {
                    attributes.push(("compact", "0"));

                    if self.layout == PivotTableLayout::Tabular {
                        attributes.push(("outline", "0"));
                        attributes.push(("subtotalTop", "0"));
                    }
                }

                attributes.push(("showAll", "0"));

                // A field can drop the subtotal that it gets by default.
                if self.no_subtotal_fields.contains(&index) {
                    attributes.push(("defaultSubtotal", "0"));
                }

                xml_start_tag(&mut self.writer, "pivotField", &attributes);

                // Write the placeholder items element. The real items are added
                // by the application that opens the file.
                self.write_default_items();

                xml_end_tag(&mut self.writer, "pivotField");
            }
            None => {
                let is_data_field = self
                    .data_fields
                    .iter()
                    .any(|data_field| data_field.field_index == index);

                if is_data_field {
                    let attributes = [("dataField", "1"), ("showAll", "0")];
                    xml_empty_tag(&mut self.writer, "pivotField", &attributes);
                } else {
                    let attributes = [("showAll", "0")];
                    xml_empty_tag(&mut self.writer, "pivotField", &attributes);
                }
            }
        }
    }

    // Write the <items> element with a single default item.
    fn write_default_items(&mut self) {
        let attributes = [("count", "1")];

        xml_start_tag(&mut self.writer, "items", &attributes);

        let attributes = [("t", "default")];
        xml_empty_tag(&mut self.writer, "item", &attributes);

        xml_end_tag(&mut self.writer, "items");
    }

    // Write the <rowFields>/<colFields> elements.
    fn write_axis_fields(&mut self, tag: &str, fields: &[usize]) {
        let attributes = [("count", fields.len().to_string())];

        xml_start_tag(&mut self.writer, tag, &attributes);

        for index in fields {
            let attributes = [("x", index.to_string())];
            xml_empty_tag(&mut self.writer, "field", &attributes);
        }

        xml_end_tag(&mut self.writer, tag);
    }

    // Write the <rowItems>/<colItems> elements. Only the grand total item is
    // written, the rest are added by the application that opens the file.
    fn write_axis_items(&mut self, tag: &str) {
        let attributes = [("count", "1")];

        xml_start_tag(&mut self.writer, tag, &attributes);

        let attributes = [("t", "grand")];
        xml_start_tag(&mut self.writer, "i", &attributes);
        xml_empty_tag_only(&mut self.writer, "x");
        xml_end_tag(&mut self.writer, "i");

        xml_end_tag(&mut self.writer, tag);
    }

    // Write the <pageFields> element.
    fn write_page_fields(&mut self) {
        let attributes = [("count", self.filter_fields.len().to_string())];

        xml_start_tag(&mut self.writer, "pageFields", &attributes);

        for index in self.filter_fields.clone() {
            let attributes = [("fld", index.to_string()), ("hier", "-1".to_string())];
            xml_empty_tag(&mut self.writer, "pageField", &attributes);
        }

        xml_end_tag(&mut self.writer, "pageFields");
    }

    // Write the <dataFields> element.
    fn write_data_fields(&mut self) {
        let attributes = [("count", self.data_fields.len().to_string())];

        xml_start_tag(&mut self.writer, "dataFields", &attributes);

        for data_field in self.data_fields.clone() {
            // Write the dataField element.
            self.write_data_field(&data_field);
        }

        xml_end_tag(&mut self.writer, "dataFields");
    }

    // Write the <dataField> element.
    fn write_data_field(&mut self, data_field: &PivotTableDataField) {
        let mut attributes = vec![
            ("name", data_field.display_name()),
            ("fld", data_field.field_index.to_string()),
        ];

        // The Sum function is the Excel default and isn't written out.
        if data_field.function != PivotTableFunction::Sum {
            attributes.push(("subtotal", data_field.function.to_string()));
        }

        if data_field.num_format_index > 0 {
            attributes.push(("numFmtId", data_field.num_format_index.to_string()));
        }

        xml_empty_tag(&mut self.writer, "dataField", &attributes);
    }

    // Write the <pivotTableStyleInfo> element.
    fn write_pivot_table_style_info(&mut self) {
        let mut attributes = vec![];

        // Excel omits the name attribute for the "None" style.
        if self.style != PivotTableStyle::None {
            attributes.push(("name", self.style.to_string()));
        }

        attributes.extend([
            ("showRowHeaders", "1".to_string()),
            ("showColHeaders", "1".to_string()),
            ("showRowStripes", "0".to_string()),
            ("showColStripes", "0".to_string()),
            ("showLastColumn", "0".to_string()),
        ]);

        xml_empty_tag(&mut self.writer, "pivotTableStyleInfo", &attributes);
    }
}

// -----------------------------------------------------------------------
// PivotTableDataField
// -----------------------------------------------------------------------

/// The `PivotTableDataField` struct represents a data/values field in a pivot
/// table.
///
/// A data field is a field of the source data that is summarized by the pivot
/// table using a function such as [`PivotTableFunction::Sum`]. It is added to a
/// pivot table via the [`PivotTable::add_data_field()`] method.
///
#[derive(Clone, PartialEq)]
pub struct PivotTableDataField {
    pub(crate) field_name: String,
    pub(crate) name: String,
    pub(crate) function: PivotTableFunction,
    pub(crate) num_format: String,

    // Set in Workbook::prepare_pivot_tables().
    pub(crate) field_index: usize,
    pub(crate) num_format_index: u16,
}

impl PivotTableDataField {
    /// Create a new `PivotTableDataField` struct instance.
    ///
    /// # Parameters
    ///
    /// - `field_name`: The name of a field in the header row of the pivot table
    ///   source data.
    ///
    /// # Examples
    ///
    /// Example of creating a data field for the values area of a pivot table.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_data_field_new.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a data field for the values area of a pivot table.
    ///     let data_field = PivotTableDataField::new("Volume");
    ///
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_row_field("Region")
    ///         .add_data_field(data_field);
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn new(field_name: impl Into<String>) -> PivotTableDataField {
        PivotTableDataField {
            field_name: field_name.into(),
            name: String::new(),
            function: PivotTableFunction::Sum,
            num_format: String::new(),
            field_index: 0,
            num_format_index: 0,
        }
    }

    /// Set the function used to summarize the data field.
    ///
    /// The default function is [`PivotTableFunction::Sum`].
    ///
    /// # Parameters
    ///
    /// - `function`: A [`PivotTableFunction`] enum value.
    ///
    /// # Examples
    ///
    /// Example of setting the summary function of a pivot table data field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_data_field_set_function.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     PivotTable, PivotTableDataField, PivotTableFunction, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a data field that shows the maximum volume.
    ///     let data_field = PivotTableDataField::new("Volume").set_function(PivotTableFunction::Max);
    ///
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_row_field("Region")
    ///         .add_data_field(data_field);
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_function(mut self, function: PivotTableFunction) -> PivotTableDataField {
        self.function = function;
        self
    }

    /// Set a custom name for the data field.
    ///
    /// By default a data field is named after its function and field name, like
    /// "Sum of Volume". This method overrides that with a user defined name.
    ///
    /// # Parameters
    ///
    /// - `name`: The name to display in the pivot table.
    ///
    /// # Examples
    ///
    /// Example of setting the caption of a pivot table data field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_data_field_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a data field with a caption of its own instead of the default
    ///     // "Sum of Volume".
    ///     let data_field = PivotTableDataField::new("Volume").set_name("Total volume");
    ///
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_row_field("Region")
    ///         .add_data_field(data_field);
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_name(mut self, name: impl Into<String>) -> PivotTableDataField {
        self.name = name.into();
        self
    }

    /// Set the number format of the data field.
    ///
    /// The values of a data field are displayed with a general number format by
    /// default. This method sets an Excel number format string like `"0.00"` or
    /// `"[h]:mm:ss"` for the summarized values.
    ///
    /// See also [`Format::set_num_format()`](crate::Format::set_num_format) for
    /// an explanation of Excel number formats.
    ///
    /// # Parameters
    ///
    /// - `num_format`: An Excel number format string.
    ///
    /// # Examples
    ///
    /// Example of setting the number format of a pivot table data field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_data_field_set_num_format.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableDataField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with the source data for the pivot table.
    /// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
    /// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
    /// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
    /// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
    /// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
    /// #
    ///     // Create a data field with a number format.
    ///     let data_field = PivotTableDataField::new("Volume").set_num_format("#,##0.00");
    ///
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Data", 0, 0, 3, 3))
    ///         .add_row_field("Region")
    ///         .add_data_field(data_field);
    /// #
    /// #     // Add the pivot table to a new worksheet.
    /// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
    /// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_num_format(mut self, num_format: impl Into<String>) -> PivotTableDataField {
        self.num_format = num_format.into();
        self
    }

    // Get the user defined name or the Excel style default name.
    pub(crate) fn display_name(&self) -> String {
        if self.name.is_empty() {
            format!("{} of {}", self.function.name(), self.field_name)
        } else {
            self.name.clone()
        }
    }
}

impl From<&PivotTableDataField> for PivotTableDataField {
    fn from(value: &PivotTableDataField) -> PivotTableDataField {
        value.clone()
    }
}

// -----------------------------------------------------------------------
// PivotTableFunction
// -----------------------------------------------------------------------

/// The `PivotTableFunction` enum defines the summary functions used by a
/// [`PivotTableDataField`].
///
/// # Examples
///
/// Example of setting the summary function of a pivot table data field.
///
/// ```
/// # // This code is available in examples/doc_pivot_table_data_field_set_function.rs
/// #
/// # use rust_xlsxwriter::{
/// #     PivotTable, PivotTableDataField, PivotTableFunction, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with the source data for the pivot table.
/// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
/// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
/// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
/// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
/// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
/// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
/// #
///     // Create a data field that shows the maximum volume.
///     let data_field = PivotTableDataField::new("Volume").set_function(PivotTableFunction::Max);
///
///     let pivot_table = PivotTable::new()
///         .set_data_source(("Data", 0, 0, 3, 3))
///         .add_row_field("Region")
///         .add_data_field(data_field);
/// #
/// #     // Add the pivot table to a new worksheet.
/// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
/// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("pivot_table.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum PivotTableFunction {
    /// Sum of the values. This is the default.
    Sum,

    /// Count of the values, like the Excel `COUNTA()` function.
    Count,

    /// Average of the values.
    Average,

    /// Maximum of the values.
    Max,

    /// Minimum of the values.
    Min,

    /// Product of the values.
    Product,

    /// Count of the numeric values, like the Excel `COUNT()` function.
    CountNumbers,

    /// Standard deviation of a sample of the values.
    StdDev,

    /// Standard deviation of the values as a population.
    StdDevPop,

    /// Variance of a sample of the values.
    Var,

    /// Variance of the values as a population.
    VarPop,
}

impl PivotTableFunction {
    // Get the name used by Excel in the default data field caption.
    pub(crate) fn name(self) -> &'static str {
        match self {
            Self::Sum => "Sum",
            Self::Count | Self::CountNumbers => "Count",
            Self::Average => "Average",
            Self::Max => "Max",
            Self::Min => "Min",
            Self::Product => "Product",
            Self::StdDev => "StdDev",
            Self::StdDevPop => "StdDevp",
            Self::Var => "Var",
            Self::VarPop => "Varp",
        }
    }
}

impl fmt::Display for PivotTableFunction {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Sum => write!(f, "sum"),
            Self::Count => write!(f, "count"),
            Self::Average => write!(f, "average"),
            Self::Max => write!(f, "max"),
            Self::Min => write!(f, "min"),
            Self::Product => write!(f, "product"),
            Self::CountNumbers => write!(f, "countNums"),
            Self::StdDev => write!(f, "stdDev"),
            Self::StdDevPop => write!(f, "stdDevp"),
            Self::Var => write!(f, "var"),
            Self::VarPop => write!(f, "varp"),
        }
    }
}

// -----------------------------------------------------------------------
// PivotTableStyle
// -----------------------------------------------------------------------

/// The `PivotTableStyle` enum defines the built-in Excel pivot table styles.
///
/// Excel supports 84 built-in pivot table styles in the "Light", "Medium" and
/// "Dark" categories, plus a "None" style. The default style used by
/// `rust_xlsxwriter`, like Excel, is [`PivotTableStyle::Light16`].
///
/// The style is set via the [`PivotTable::set_style()`] method.
///
/// # Examples
///
/// Example of setting the style of a pivot table.
///
/// ```
/// # // This code is available in examples/doc_pivot_table_set_style.rs
/// #
/// # use rust_xlsxwriter::{
/// #     PivotTable, PivotTableDataField, PivotTableStyle, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with the source data for the pivot table.
/// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
/// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
/// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
/// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
/// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
/// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
/// #
///     // Create a pivot table and set its style.
///     let pivot_table = PivotTable::new()
///         .set_data_source(("Data", 0, 0, 3, 3))
///         .set_style(PivotTableStyle::Medium10)
///         .add_row_field("Region")
///         .add_data_field(PivotTableDataField::new("Volume"));
/// #
/// #     // Add the pivot table to a new worksheet.
/// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
/// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("pivot_table.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum PivotTableStyle {
    /// No pivot table style.
    None,
    /// Pivot Style Light 1.
    Light1,
    /// Pivot Style Light 2.
    Light2,
    /// Pivot Style Light 3.
    Light3,
    /// Pivot Style Light 4.
    Light4,
    /// Pivot Style Light 5.
    Light5,
    /// Pivot Style Light 6.
    Light6,
    /// Pivot Style Light 7.
    Light7,
    /// Pivot Style Light 8.
    Light8,
    /// Pivot Style Light 9.
    Light9,
    /// Pivot Style Light 10.
    Light10,
    /// Pivot Style Light 11.
    Light11,
    /// Pivot Style Light 12.
    Light12,
    /// Pivot Style Light 13.
    Light13,
    /// Pivot Style Light 14.
    Light14,
    /// Pivot Style Light 15.
    Light15,
    /// Pivot Style Light 16.
    Light16,
    /// Pivot Style Light 17.
    Light17,
    /// Pivot Style Light 18.
    Light18,
    /// Pivot Style Light 19.
    Light19,
    /// Pivot Style Light 20.
    Light20,
    /// Pivot Style Light 21.
    Light21,
    /// Pivot Style Light 22.
    Light22,
    /// Pivot Style Light 23.
    Light23,
    /// Pivot Style Light 24.
    Light24,
    /// Pivot Style Light 25.
    Light25,
    /// Pivot Style Light 26.
    Light26,
    /// Pivot Style Light 27.
    Light27,
    /// Pivot Style Light 28.
    Light28,
    /// Pivot Style Medium 1.
    Medium1,
    /// Pivot Style Medium 2.
    Medium2,
    /// Pivot Style Medium 3.
    Medium3,
    /// Pivot Style Medium 4.
    Medium4,
    /// Pivot Style Medium 5.
    Medium5,
    /// Pivot Style Medium 6.
    Medium6,
    /// Pivot Style Medium 7.
    Medium7,
    /// Pivot Style Medium 8.
    Medium8,
    /// Pivot Style Medium 9.
    Medium9,
    /// Pivot Style Medium 10.
    Medium10,
    /// Pivot Style Medium 11.
    Medium11,
    /// Pivot Style Medium 12.
    Medium12,
    /// Pivot Style Medium 13.
    Medium13,
    /// Pivot Style Medium 14.
    Medium14,
    /// Pivot Style Medium 15.
    Medium15,
    /// Pivot Style Medium 16.
    Medium16,
    /// Pivot Style Medium 17.
    Medium17,
    /// Pivot Style Medium 18.
    Medium18,
    /// Pivot Style Medium 19.
    Medium19,
    /// Pivot Style Medium 20.
    Medium20,
    /// Pivot Style Medium 21.
    Medium21,
    /// Pivot Style Medium 22.
    Medium22,
    /// Pivot Style Medium 23.
    Medium23,
    /// Pivot Style Medium 24.
    Medium24,
    /// Pivot Style Medium 25.
    Medium25,
    /// Pivot Style Medium 26.
    Medium26,
    /// Pivot Style Medium 27.
    Medium27,
    /// Pivot Style Medium 28.
    Medium28,
    /// Pivot Style Dark 1.
    Dark1,
    /// Pivot Style Dark 2.
    Dark2,
    /// Pivot Style Dark 3.
    Dark3,
    /// Pivot Style Dark 4.
    Dark4,
    /// Pivot Style Dark 5.
    Dark5,
    /// Pivot Style Dark 6.
    Dark6,
    /// Pivot Style Dark 7.
    Dark7,
    /// Pivot Style Dark 8.
    Dark8,
    /// Pivot Style Dark 9.
    Dark9,
    /// Pivot Style Dark 10.
    Dark10,
    /// Pivot Style Dark 11.
    Dark11,
    /// Pivot Style Dark 12.
    Dark12,
    /// Pivot Style Dark 13.
    Dark13,
    /// Pivot Style Dark 14.
    Dark14,
    /// Pivot Style Dark 15.
    Dark15,
    /// Pivot Style Dark 16.
    Dark16,
    /// Pivot Style Dark 17.
    Dark17,
    /// Pivot Style Dark 18.
    Dark18,
    /// Pivot Style Dark 19.
    Dark19,
    /// Pivot Style Dark 20.
    Dark20,
    /// Pivot Style Dark 21.
    Dark21,
    /// Pivot Style Dark 22.
    Dark22,
    /// Pivot Style Dark 23.
    Dark23,
    /// Pivot Style Dark 24.
    Dark24,
    /// Pivot Style Dark 25.
    Dark25,
    /// Pivot Style Dark 26.
    Dark26,
    /// Pivot Style Dark 27.
    Dark27,
    /// Pivot Style Dark 28.
    Dark28,
}

impl fmt::Display for PivotTableStyle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::None => write!(f, "PivotStyleNone"),
            Self::Light1 => write!(f, "PivotStyleLight1"),
            Self::Light2 => write!(f, "PivotStyleLight2"),
            Self::Light3 => write!(f, "PivotStyleLight3"),
            Self::Light4 => write!(f, "PivotStyleLight4"),
            Self::Light5 => write!(f, "PivotStyleLight5"),
            Self::Light6 => write!(f, "PivotStyleLight6"),
            Self::Light7 => write!(f, "PivotStyleLight7"),
            Self::Light8 => write!(f, "PivotStyleLight8"),
            Self::Light9 => write!(f, "PivotStyleLight9"),
            Self::Light10 => write!(f, "PivotStyleLight10"),
            Self::Light11 => write!(f, "PivotStyleLight11"),
            Self::Light12 => write!(f, "PivotStyleLight12"),
            Self::Light13 => write!(f, "PivotStyleLight13"),
            Self::Light14 => write!(f, "PivotStyleLight14"),
            Self::Light15 => write!(f, "PivotStyleLight15"),
            Self::Light16 => write!(f, "PivotStyleLight16"),
            Self::Light17 => write!(f, "PivotStyleLight17"),
            Self::Light18 => write!(f, "PivotStyleLight18"),
            Self::Light19 => write!(f, "PivotStyleLight19"),
            Self::Light20 => write!(f, "PivotStyleLight20"),
            Self::Light21 => write!(f, "PivotStyleLight21"),
            Self::Light22 => write!(f, "PivotStyleLight22"),
            Self::Light23 => write!(f, "PivotStyleLight23"),
            Self::Light24 => write!(f, "PivotStyleLight24"),
            Self::Light25 => write!(f, "PivotStyleLight25"),
            Self::Light26 => write!(f, "PivotStyleLight26"),
            Self::Light27 => write!(f, "PivotStyleLight27"),
            Self::Light28 => write!(f, "PivotStyleLight28"),
            Self::Medium1 => write!(f, "PivotStyleMedium1"),
            Self::Medium2 => write!(f, "PivotStyleMedium2"),
            Self::Medium3 => write!(f, "PivotStyleMedium3"),
            Self::Medium4 => write!(f, "PivotStyleMedium4"),
            Self::Medium5 => write!(f, "PivotStyleMedium5"),
            Self::Medium6 => write!(f, "PivotStyleMedium6"),
            Self::Medium7 => write!(f, "PivotStyleMedium7"),
            Self::Medium8 => write!(f, "PivotStyleMedium8"),
            Self::Medium9 => write!(f, "PivotStyleMedium9"),
            Self::Medium10 => write!(f, "PivotStyleMedium10"),
            Self::Medium11 => write!(f, "PivotStyleMedium11"),
            Self::Medium12 => write!(f, "PivotStyleMedium12"),
            Self::Medium13 => write!(f, "PivotStyleMedium13"),
            Self::Medium14 => write!(f, "PivotStyleMedium14"),
            Self::Medium15 => write!(f, "PivotStyleMedium15"),
            Self::Medium16 => write!(f, "PivotStyleMedium16"),
            Self::Medium17 => write!(f, "PivotStyleMedium17"),
            Self::Medium18 => write!(f, "PivotStyleMedium18"),
            Self::Medium19 => write!(f, "PivotStyleMedium19"),
            Self::Medium20 => write!(f, "PivotStyleMedium20"),
            Self::Medium21 => write!(f, "PivotStyleMedium21"),
            Self::Medium22 => write!(f, "PivotStyleMedium22"),
            Self::Medium23 => write!(f, "PivotStyleMedium23"),
            Self::Medium24 => write!(f, "PivotStyleMedium24"),
            Self::Medium25 => write!(f, "PivotStyleMedium25"),
            Self::Medium26 => write!(f, "PivotStyleMedium26"),
            Self::Medium27 => write!(f, "PivotStyleMedium27"),
            Self::Medium28 => write!(f, "PivotStyleMedium28"),
            Self::Dark1 => write!(f, "PivotStyleDark1"),
            Self::Dark2 => write!(f, "PivotStyleDark2"),
            Self::Dark3 => write!(f, "PivotStyleDark3"),
            Self::Dark4 => write!(f, "PivotStyleDark4"),
            Self::Dark5 => write!(f, "PivotStyleDark5"),
            Self::Dark6 => write!(f, "PivotStyleDark6"),
            Self::Dark7 => write!(f, "PivotStyleDark7"),
            Self::Dark8 => write!(f, "PivotStyleDark8"),
            Self::Dark9 => write!(f, "PivotStyleDark9"),
            Self::Dark10 => write!(f, "PivotStyleDark10"),
            Self::Dark11 => write!(f, "PivotStyleDark11"),
            Self::Dark12 => write!(f, "PivotStyleDark12"),
            Self::Dark13 => write!(f, "PivotStyleDark13"),
            Self::Dark14 => write!(f, "PivotStyleDark14"),
            Self::Dark15 => write!(f, "PivotStyleDark15"),
            Self::Dark16 => write!(f, "PivotStyleDark16"),
            Self::Dark17 => write!(f, "PivotStyleDark17"),
            Self::Dark18 => write!(f, "PivotStyleDark18"),
            Self::Dark19 => write!(f, "PivotStyleDark19"),
            Self::Dark20 => write!(f, "PivotStyleDark20"),
            Self::Dark21 => write!(f, "PivotStyleDark21"),
            Self::Dark22 => write!(f, "PivotStyleDark22"),
            Self::Dark23 => write!(f, "PivotStyleDark23"),
            Self::Dark24 => write!(f, "PivotStyleDark24"),
            Self::Dark25 => write!(f, "PivotStyleDark25"),
            Self::Dark26 => write!(f, "PivotStyleDark26"),
            Self::Dark27 => write!(f, "PivotStyleDark27"),
            Self::Dark28 => write!(f, "PivotStyleDark28"),
        }
    }
}

// -----------------------------------------------------------------------
// PivotTableLayout
// -----------------------------------------------------------------------

/// The `PivotTableLayout` enum defines the layout of the row area of a pivot
/// table.
///
/// The layout is set via the [`PivotTable::set_layout()`] method.
///
/// # Examples
///
/// Example of setting the row layout of a pivot table.
///
/// ```
/// # // This code is available in examples/doc_pivot_table_set_layout.rs
/// #
/// # use rust_xlsxwriter::{
/// #     PivotTable, PivotTableDataField, PivotTableLayout, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with the source data for the pivot table.
/// #     let worksheet = workbook.add_worksheet().set_name("Data")?;
/// #     worksheet.write_row(0, 0, ["Region", "Item", "Month", "Volume"])?;
/// #     worksheet.write_row(1, 0, ["East", "Apple", "July"])?;
/// #     worksheet.write_row(2, 0, ["West", "Apple", "April"])?;
/// #     worksheet.write_row(3, 0, ["East", "Pear", "July"])?;
/// #     worksheet.write_column(1, 3, [9000, 5000, 7000])?;
/// #
///     // Create a pivot table with two row fields in tabular layout so that each
///     // row field gets a column of its own.
///     let pivot_table = PivotTable::new()
///         .set_data_source(("Data", 0, 0, 3, 3))
///         .set_layout(PivotTableLayout::Tabular)
///         .add_row_field("Region")
///         .add_row_field("Item")
///         .add_data_field(PivotTableDataField::new("Volume"));
/// #
/// #     // Add the pivot table to a new worksheet.
/// #     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
/// #     worksheet.add_pivot_table(0, 0, &pivot_table)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("pivot_table.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum PivotTableLayout {
    /// Compact form. All the row fields share a single column and nested fields
    /// are indented. This is the default, like Excel.
    Compact,

    /// Outline form. Each row field gets its own column but the sub-items are
    /// written on the row below the field name.
    Outline,

    /// Tabular form. Each row field gets its own column and the sub-items are
    /// written on the same row as the field name, like a classic report.
    Tabular,
}
