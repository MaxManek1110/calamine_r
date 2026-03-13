
<!-- README.md is generated from README.Rmd. Please edit that file -->

# calamineR

Fast Excel file reader for R, powered by the Rust
[calamine](https://github.com/tafia/calamine) library.

## Supported Formats

- xlsx (Excel 2007+)
- xlsm (Excel with macros)
- xlsb (Excel Binary)
- xls (Excel 97-2003)
- ods (OpenDocument Spreadsheet)

## System Requirements

- R \>= 4.0
- Rust toolchain (cargo \>= 1.75.0) - install from <https://rustup.rs/>

## Installation

You can install the development version of calamineR from
[GitHub](https://github.com/) with:

``` r
# install.packages("devtools")
devtools::install_github("MaxManek1110/calamineR")

# from CRAN (if available)
install.packages("calamineR")
```

## Usage

``` r
library(calamineR)

# Create a sample Excel file for demonstration
sample_data <- data.frame(
  name = c("Alice", "Bob", "Charlie"),
  age = c(25, 30, 35),
  score = c(85.5, 92.3, 78.9)
)
demo_file <- tempfile(fileext = ".xlsx")
writexl::write_xlsx(sample_data, demo_file)

# Read a sheet as data.frame
df <- read_excel(demo_file)
print(df)
#>      name age score
#> 1   Alice  25  85.5
#> 2     Bob  30  92.3
#> 3 Charlie  35  78.9

# Read with different options
df <- read_excel(demo_file, col_names = FALSE)
head(df, 3)
#>      V1  V2    V3
#> 1  name age score
#> 2 Alice  25  85.5
#> 3   Bob  30  92.3

df <- read_excel(demo_file, skip = 1)  # Skip header row
head(df, 2)
#>     Alice 25 85.5
#> 1     Bob 30 92.3
#> 2 Charlie 35 78.9

# Get sheet names
sheets <- excel_sheets(demo_file)
print(sheets)
#> [1] "Sheet1"

# Get sheet dimensions
dims <- sheet_dims(demo_file, 1)
cat("Rows:", dims["rows"], "Cols:", dims["cols"], "\n")
#> Rows: 4 Cols: 3

# Read as raw list of rows (for complex layouts)
rows <- read_sheet_raw(demo_file, "Sheet1")
str(rows)
#> List of 4
#>  $ : chr [1:3] "name" "age" "score"
#>  $ : chr [1:3] "Alice" "25" "85.5"
#>  $ : chr [1:3] "Bob" "30" "92.3"
#>  $ : chr [1:3] "Charlie" "35" "78.9"

# Cleanup
unlink(demo_file)
```

## Functions

| Function           | Description                                          |
|--------------------|------------------------------------------------------|
| `read_excel()`     | Read sheet as data.frame                             |
| `excel_sheets()`   | Get sheet names                                      |
| `sheet_dims()`     | Get sheet dimensions (rows, cols)                    |
| `read_sheet_raw()` | Read as list of character vectors                    |
| `merge_regions()`  | Get information about merged cell regions in a sheet |

## Parameters for `read_excel()`

| Parameter | Type | Default | Description |
|----|----|----|----|
| `path` | character | (required) | Path to the Excel file |
| `sheet` | character/integer | `1L` | Sheet name or 1-based index |
| `col_names` | logical | `TRUE` | Use first row as column names |
| `skip` | integer | `0L` | Number of rows to skip |
| `fill_merged_cells` | logical | `FALSE` | Fill merged cells with top-left value |

## Merged Cells Support

When `fill_merged_cells = TRUE`, cells that are part of a merged region
are filled with the value from the top-left cell of that region. This is
useful when reading spreadsheets where headers or data span multiple
cells.

Supported formats: xlsx, xlsm, xlsb, xls (ods not yet supported for
merged cells).

## Performance

Calamine is written in pure Rust and is significantly faster than many
alternatives especially for binary files. However as there is no package
to write xlsb files in R, the benchmark only includes xlsx format where
the difference is not that significant. Performance may vary based on
file size, format, and system resources.

## Benchmark

Comparison against `readxl` for reading a large Excel file with 500,000
rows.

``` r
library(calamineR)

# Create a large test file with 500k rows x 10 columns
set.seed(42)
n_rows <- 500000
test_data <- data.frame(
  id = seq_len(n_rows),
  name = paste0("item_", seq_len(n_rows)),
  value1 = rnorm(n_rows),
  value2 = rnorm(n_rows),
  value3 = runif(n_rows),
  category = sample(LETTERS[1:5], n_rows, replace = TRUE),
  date = as.character(Sys.Date() + sample(1:1000, n_rows, replace = TRUE)),
  flag = sample(c("yes", "no"), n_rows, replace = TRUE),
  amount = round(runif(n_rows, 100, 10000), 2),

  notes = paste0("note_", sample(1:100, n_rows, replace = TRUE))
)

# Write to temporary xlsx file
temp_file <- tempfile(fileext = ".xlsx")
writexl::write_xlsx(test_data, temp_file)
file_size_mb <- round(file.size(temp_file) / 1024^2, 1)
cat("Test file size:", file_size_mb, "MB\n")
#> Test file size: 42 MB
cat("Dimensions:", n_rows, "rows x", ncol(test_data), "columns\n\n")
#> Dimensions: 5e+05 rows x 10 columns

# Benchmark calamineR
cat("calamineR::read_excel()\n")
#> calamineR::read_excel()
time_calamineR <- system.time({
  df_calamineR <- calamineR::read_excel(temp_file)
})
cat("  Time:", round(time_calamineR["elapsed"], 3), "seconds\n")
#>   Time: 7.069 seconds
cat("  Rows:", nrow(df_calamineR), " Cols:", ncol(df_calamineR), "\n\n")
#>   Rows: 500000  Cols: 10

# Benchmark readxl
cat("readxl::read_xlsx()\n")
#> readxl::read_xlsx()
time_readxl <- system.time({
  df_readxl <- readxl::read_xlsx(temp_file)
})
cat("  Time:", round(time_readxl["elapsed"], 3), "seconds\n")
#>   Time: 7.814 seconds
cat("  Rows:", nrow(df_readxl), " Cols:", ncol(df_readxl), "\n\n")
#>   Rows: 500000  Cols: 10

# Summary
cat("--- Summary ---\n")
#> --- Summary ---
cat("calamineR:", round(time_calamineR["elapsed"], 3), "s\n")
#> calamineR: 7.069 s
cat("readxl:   ", round(time_readxl["elapsed"], 3), "s\n")
#> readxl:    7.814 s
cat("Speedup:  ", round(time_readxl["elapsed"] / time_calamineR["elapsed"], 1), "x faster\n")
#> Speedup:   1.1 x faster

# Cleanup
unlink(temp_file)
```

## Development

`calamineR` is built using `extendr` and `Rust`. To build the package,
you need to have `Rust` installed on your machine.

``` r
rextendr::document()
devtools::document()
```
