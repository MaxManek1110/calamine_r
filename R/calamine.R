# Helper function to validate sheet parameter against available sheets
validate_sheet <- function(path, sheet, require_worksheet = TRUE) {
  sheets <- cal_sheet_names(path)
  if (is.numeric(sheet)) {
    if (sheet > length(sheets)) {
      stop(sprintf(
        "Sheet index %d is out of range. File has %d sheet(s): %s",
        sheet, length(sheets), paste(shQuote(sheets), collapse = ", ")
      ), call. = FALSE)
    }
  } else {
    if (!sheet %in% sheets) {
      stop(sprintf(
        "Sheet '%s' not found. Available sheets: %s",
        sheet, paste(shQuote(sheets), collapse = ", ")
      ), call. = FALSE)
    }
  }
  # Check if sheet is a worksheet (not a chart/dialog sheet)
  if (require_worksheet && !cal_is_worksheet(path, sheet)) {
    meta <- cal_sheets_metadata(path)
    sheet_name <- if (is.numeric(sheet)) sheets[sheet] else sheet
    sheet_type <- meta$type[meta$name == sheet_name]
    worksheets <- meta$name[meta$type == "worksheet"]
    stop(sprintf(
      "Sheet '%s' is a %s, not a worksheet. Cannot read data from chart/dialog sheets. Available worksheets: %s",
      sheet_name, sheet_type,
      if (length(worksheets) > 0) paste(shQuote(worksheets), collapse = ", ") else "(none)"
    ), call. = FALSE)
  }
  invisible(TRUE)
}

#' Read Excel File Using Calamine
#'
#' Fast Excel reader powered by the Rust calamine library.
#' Supports xlsx, xlsm, xlsb, xls, and ods formats.
#'
#' @param path Path to the Excel file
#' @param sheet Sheet name (character) or index (integer, 1-based). Default: 1
#' @param col_names Use first row as column names. Default: TRUE
#' @param skip Number of rows to skip before reading. Default: 0
#' @param fill_merged_cells If TRUE, fill merged cells with the value from the
#'   top-left cell of the merged region. Default: FALSE
#'
#' @return A data.frame
#' @export
#'
#' @examples
#' # Using package test file
#' test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
#' if (nzchar(test_file)) {
#'   df <- read_excel(test_file)
#'   head(df)
#'
#'   # Read specific sheet by index
#'   df <- read_excel(test_file, sheet = 1)
#'
#'   # Skip header row
#'   df_no_header <- read_excel(test_file, col_names = FALSE)
#'
#'   # Fill merged cells with their value
#'   df_filled <- read_excel(test_file, fill_merged_cells = TRUE)
#' }
read_excel <- function(path, sheet = 1L, col_names = TRUE, skip = 0L,
                       fill_merged_cells = FALSE) {
  stopifnot(
    "`path` must be a single character string" = is.character(path) && length(path) == 1,
    "File does not exist" = file.exists(path),
    "`sheet` must be length 1" = length(sheet) == 1,
    "`sheet` must be character or numeric" = is.character(sheet) || is.numeric(sheet),
    "`col_names` must be TRUE or FALSE" = is.logical(col_names) && length(col_names) == 1,
    "`skip` must be a non-negative number" = is.numeric(skip) && length(skip) == 1 && skip >= 0,
    "`fill_merged_cells` must be TRUE or FALSE" = is.logical(fill_merged_cells) && length(fill_merged_cells) == 1
  )
  if (is.numeric(sheet)) stopifnot("`sheet` index must be >= 1" = sheet >= 1)
  path <- normalizePath(path)
  ext <- tolower(tools::file_ext(path))
  stopifnot("Unsupported file format. Use: xlsx, xlsm, xlsb, xls, ods" = ext %in% c("xlsx", "xlsm", "xlsb", "xls", "ods"))
  validate_sheet(path, sheet)
  cal_read_sheet_df(path, sheet, col_names, as.integer(skip), fill_merged_cells)
}

#' Get Sheet Names from Excel File
#'
#' @param path Path to the Excel file
#' @param include_charts If FALSE (default), only return worksheet names.
#'   If TRUE, return all sheet names including chart sheets.
#' @return Character vector of sheet names
#' @export
#'
#' @examples
#' test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
#' if (nzchar(test_file)) {
#'   sheets <- excel_sheets(test_file)
#'   print(sheets)
#' }
excel_sheets <- function(path, include_charts = FALSE) {
  stopifnot(
    "`path` must be a single character string" = is.character(path) && length(path) == 1,
    "File does not exist" = file.exists(path),
    "`include_charts` must be TRUE or FALSE" = is.logical(include_charts) && length(include_charts) == 1
  )
  path <- normalizePath(path)
  ext <- tolower(tools::file_ext(path))
  stopifnot("Unsupported file format. Use: xlsx, xlsm, xlsb, xls, ods" = ext %in% c("xlsx", "xlsm", "xlsb", "xls", "ods"))
  if (include_charts) {
    cal_sheet_names(path)
  } else {
    meta <- cal_sheets_metadata(path)
    meta$name[meta$type == "worksheet"]
  }
}

#' Get Sheet Metadata from Excel File
#'
#' Returns detailed information about all sheets including their type
#' (worksheet, chartsheet, etc.) and visibility.
#'
#' @param path Path to the Excel file
#' @return A data.frame with columns:
#'   \item{name}{Sheet name}
#'   \item{type}{Sheet type: "worksheet", "chartsheet", "dialogsheet", "macrosheet", or "vba"}
#'   \item{visible}{Logical indicating if the sheet is visible}
#' @export
#'
#' @examples
#' test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
#' if (nzchar(test_file)) {
#'   meta <- sheets_metadata(test_file)
#'   print(meta)
#' }
sheets_metadata <- function(path) {
  stopifnot(
    "`path` must be a single character string" = is.character(path) && length(path) == 1,
    "File does not exist" = file.exists(path)
  )
  path <- normalizePath(path)
  ext <- tolower(tools::file_ext(path))
  stopifnot("Unsupported file format. Use: xlsx, xlsm, xlsb, xls, ods" = ext %in% c("xlsx", "xlsm", "xlsb", "xls", "ods"))
  cal_sheets_metadata(path)
}

#' Get Sheet Dimensions
#'
#' @param path Path to the Excel file
#' @param sheet Sheet name or index (1-based)
#' @return Named integer vector with "rows" and "cols"
#' @export
#'
#' @examples
#' test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
#' if (nzchar(test_file)) {
#'   dims <- sheet_dims(test_file, 1)
#'   print(dims)  # Named vector: rows, cols
#' }
sheet_dims <- function(path, sheet = 1L) {
  stopifnot(
    "`path` must be a single character string" = is.character(path) && length(path) == 1,
    "File does not exist" = file.exists(path),
    "`sheet` must be length 1" = length(sheet) == 1,
    "`sheet` must be character or numeric" = is.character(sheet) || is.numeric(sheet)
  )
  if (is.numeric(sheet)) stopifnot("`sheet` index must be >= 1" = sheet >= 1)
  path <- normalizePath(path)
  ext <- tolower(tools::file_ext(path))
  stopifnot("Unsupported file format. Use: xlsx, xlsm, xlsb, xls, ods" = ext %in% c("xlsx", "xlsm", "xlsb", "xls", "ods"))
  validate_sheet(path, sheet)
  dims <- cal_sheet_dims(path, sheet)
  names(dims) <- c("rows", "cols")
  dims
}

#' Read Sheet as Raw Rows
#'
#' Returns sheet data as a list of character vectors (one per row).
#' Useful for complex layouts where data.frame structure doesn't fit.
#'
#' @param path Path to the Excel file
#' @param sheet Sheet name or index (1-based)
#' @return List of character vectors
#' @export
#'
#' @examples
#' test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
#' if (nzchar(test_file)) {
#'   rows <- read_sheet_raw(test_file, 1)
#'   # Returns list of character vectors (one per row)
#'   head(rows, 3)
#' }
read_sheet_raw <- function(path, sheet = 1L) {
  stopifnot(
    "`path` must be a single character string" = is.character(path) && length(path) == 1,
    "File does not exist" = file.exists(path),
    "`sheet` must be length 1" = length(sheet) == 1,
    "`sheet` must be character or numeric" = is.character(sheet) || is.numeric(sheet)
  )
  if (is.numeric(sheet)) stopifnot("`sheet` index must be >= 1" = sheet >= 1)
  path <- normalizePath(path)
  ext <- tolower(tools::file_ext(path))
  stopifnot("Unsupported file format. Use: xlsx, xlsm, xlsb, xls, ods" = ext %in% c("xlsx", "xlsm", "xlsb", "xls", "ods"))
  validate_sheet(path, sheet)
  cal_read_sheet(path, sheet)
}

#' Get Merged Cell Regions
#'
#' Returns information about merged cell regions in a sheet.
#' Supported formats: xlsx, xlsm, xlsb, xls (not ods).
#'
#' @param path Path to the Excel file
#' @param sheet Sheet name or index (1-based). Default: 1
#' @return A data.frame with columns:
#'   \item{start_row}{First row of merged region (1-based)}
#'   \item{start_col}{First column of merged region (1-based)}
#'   \item{end_row}{Last row of merged region (1-based)}
#'   \item{end_col}{Last column of merged region (1-based)}
#' @export
#'
#' @examples
#' test_file <- system.file("extdata", "test_merged.xlsx", package = "calamineR")
#' if (nzchar(test_file)) {
#'   regions <- merge_regions(test_file, 1)
#'   print(regions)
#' }
merge_regions <- function(path, sheet = 1L) {
  stopifnot(
    "`path` must be a single character string" = is.character(path) && length(path) == 1,
    "File does not exist" = file.exists(path),
    "`sheet` must be length 1" = length(sheet) == 1,
    "`sheet` must be character or numeric" = is.character(sheet) || is.numeric(sheet)
  )
  if (is.numeric(sheet)) stopifnot("`sheet` index must be >= 1" = sheet >= 1)
  path <- normalizePath(path)
  ext <- tolower(tools::file_ext(path))
  stopifnot("Unsupported file format. Use: xlsx, xlsm, xlsb, xls, ods" = ext %in% c("xlsx", "xlsm", "xlsb", "xls", "ods"))
  validate_sheet(path, sheet)
  cal_merge_regions(path, sheet)
}
