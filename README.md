
<!-- README.md is generated from README.Rmd. Please edit that file -->

# calaminer

<!-- badges: start -->
[![R-CMD-check](https://github.com/mmanek/calaminer/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/mmanek/calaminer/actions/workflows/R-CMD-check.yaml)
[![CRAN status](https://www.r-pkg.org/badges/version/calaminer)](https://CRAN.R-project.org/package=calaminer)
<!-- badges: end -->

Fast Excel file reader for R, powered by the Rust [calamine](https://github.com/tafia/calamine) library.

## Supported Formats

- xlsx (Excel 2007+)
- xlsm (Excel with macros)
- xlsb (Excel Binary)
- xls (Excel 97-2003)
- ods (OpenDocument Spreadsheet)

## System Requirements

- R >= 4.0
- Rust toolchain (cargo >= 1.75.0) - install from https://rustup.rs/

## Installation

You can install the development version of calaminer from [GitHub](https://github.com/) with:

``` r
# install.packages("pak")
pak::pak("mmanek/calaminer")
```

## Usage

``` r
library(calaminer)

# Read a sheet as data.frame
df <- read_excel_calamine("data.xlsx")
df <- read_excel_calamine("data.xlsb", sheet = "Sheet2")
df <- read_excel_calamine("data.xlsx", sheet = 2, col_names = FALSE)
df <- read_excel_calamine("data.xlsx", skip = 5)  # Skip first 5 rows

# Get sheet names
sheets <- excel_sheets_calamine("data.xlsx")

# Get sheet dimensions
dims <- sheet_dims_calamine("data.xlsx", 1)
# dims["rows"], dims["cols"]

# Read as raw list of rows (for complex layouts)
rows <- read_sheet_raw("data.xlsx", "Sheet1")
```

## Functions

| Function | Description |
|----------|-------------|
| `read_excel_calamine()` | Read sheet as data.frame |
| `excel_sheets_calamine()` | Get sheet names |
| `sheet_dims_calamine()` | Get sheet dimensions (rows, cols) |
| `read_sheet_raw()` | Read as list of character vectors |

## Performance

Calamine is written in pure Rust and is significantly faster than many alternatives:

- 1.75x faster than excelize (Go)
- 7x faster than ClosedXML (C#)
- 9x faster than openpyxl (Python)

## License

MIT
