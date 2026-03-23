# calamineR 0.1.2

* Fixed panic when reading Excel files containing chart sheets (xlsb, xlsx, xlsm, xls)
* Added `sheets_metadata()` to inspect sheet names, types, and visibility
* Added `include_charts` parameter to `excel_sheets()` - defaults to FALSE, returning only readable worksheets
* `read_excel()` now gives a clear error message when attempting to read chart/dialog sheets instead of crashing

# calamineR 0.1.1

* Added `fill_merged_cells` parameter to `read_excel()` - when TRUE, fills merged cells with the value from the top-left cell of the merged region
* Added `merge_regions()` function to retrieve merged cell regions as a data.frame
* Merged cell support for xlsx, xlsm, xlsb, and xls formats
* Package renamed from calaminer to calamineR

# calamineR 0.1.0

* Initial CRAN submission
* Support for xlsx, xlsm, xlsb, xls, and ods formats
* Functions: `read_excel()`, `excel_sheets()`, `sheet_dims()`, `read_sheet_raw()`
* Automatic column type detection: numeric, logical, Date, and character types are inferred from cell data
