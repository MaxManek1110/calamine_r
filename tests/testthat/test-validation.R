# Test input validation and error messages
# Note: extendr wraps Rust errors in "User function panicked" messages,
# but the actual validation messages are printed to console

test_that("non-existent file throws error", {
  expect_error(
    read_excel("/nonexistent/path/file.xlsx"),
    "File does not exist"
  )
  expect_error(
    excel_sheets("/nonexistent/path/file.xlsx"),
    "File does not exist"
  )
  expect_error(
    sheet_dims("/nonexistent/path/file.xlsx", 1),
    "File does not exist"
  )
})

test_that("unsupported file format throws error", {
  tmp <- tempfile(fileext = ".csv")
  on.exit(unlink(tmp), add = TRUE)
  write.csv(mtcars, tmp, row.names = FALSE)

  expect_error(
    read_excel(tmp),
    "Unsupported file format"
  )
  expect_error(
    excel_sheets(tmp),
    "Unsupported file format"
  )
  expect_error(
    sheet_dims(tmp, 1),
    "Unsupported file format"
  )
})

test_that("sheet as vector throws error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, sheet = 1:2),
    "must be length 1"
  )
  expect_error(
    read_excel(tmp, sheet = c("Sheet1", "Sheet2")),
    "must be length 1"
  )
  expect_error(
    sheet_dims(tmp, sheet = 1:3),
    "must be length 1"
  )
  expect_error(
    read_sheet_raw(tmp, sheet = 1:2),
    "must be length 1"
  )
  expect_error(
    merge_regions(tmp, sheet = 1:2),
    "must be length 1"
  )
})

test_that("empty sheet argument throws error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, sheet = character(0))
  )
  expect_error(
    read_excel(tmp, sheet = integer(0))
  )
})

test_that("sheet name not found throws error with helpful message", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, sheet = "NonExistentSheet"),
    "Sheet 'NonExistentSheet' not found"
  )
  expect_error(
    sheet_dims(tmp, sheet = "NonExistentSheet"),
    "Sheet 'NonExistentSheet' not found"
  )
  expect_error(
    read_sheet_raw(tmp, sheet = "NonExistentSheet"),
    "Sheet 'NonExistentSheet' not found"
  )
  expect_error(
    merge_regions(tmp, sheet = "NonExistentSheet"),
    "Sheet 'NonExistentSheet' not found"
  )
})

test_that("sheet index out of range throws error with helpful message", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, sheet = 99),
    "Sheet index 99 is out of range"
  )
  expect_error(
    sheet_dims(tmp, sheet = 99),
    "Sheet index 99 is out of range"
  )
  expect_error(
    read_sheet_raw(tmp, sheet = 99),
    "Sheet index 99 is out of range"
  )
  expect_error(
    merge_regions(tmp, sheet = 99),
    "Sheet index 99 is out of range"
  )
})

test_that("sheet index less than 1 throws error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, sheet = 0),
    "must be >= 1"
  )
  expect_error(
    read_excel(tmp, sheet = -1),
    "must be >= 1"
  )
  expect_error(
    sheet_dims(tmp, sheet = 0),
    "must be >= 1"
  )
  expect_error(
    read_sheet_raw(tmp, sheet = 0),
    "must be >= 1"
  )
  expect_error(
    merge_regions(tmp, sheet = 0),
    "must be >= 1"
  )
})

test_that("negative skip throws error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, skip = -1),
    "non-negative"
  )
})

test_that("valid inputs still work correctly", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  # Numeric sheet index
  df1 <- read_excel(tmp, sheet = 1)
  expect_s3_class(df1, "data.frame")
  expect_equal(nrow(df1), nrow(mtcars))

  # Sheet name
  sheets <- excel_sheets(tmp)
  df2 <- read_excel(tmp, sheet = sheets[1])
  expect_s3_class(df2, "data.frame")

  # Sheet dims
  dims <- sheet_dims(tmp, sheet = 1)
  expect_length(dims, 2)

  # Skip parameter
  df3 <- read_excel(tmp, skip = 5)
  expect_equal(nrow(df3), nrow(mtcars) - 5)

  # read_sheet_raw with numeric index
  raw1 <- read_sheet_raw(tmp, sheet = 1)
  expect_type(raw1, "list")

  # read_sheet_raw with sheet name
  raw2 <- read_sheet_raw(tmp, sheet = sheets[1])
  expect_type(raw2, "list")

  # merge_regions with numeric index
  regions <- merge_regions(tmp, sheet = 1)
  expect_s3_class(regions, "data.frame")

  # merge_regions with sheet name
  regions2 <- merge_regions(tmp, sheet = sheets[1])
  expect_s3_class(regions2, "data.frame")
})

test_that("merge_regions validates inputs", {
  expect_error(
    merge_regions("/nonexistent/file.xlsx", 1),
    "File does not exist"
  )

  tmp <- tempfile(fileext = ".csv")
  on.exit(unlink(tmp), add = TRUE)
  write.csv(mtcars, tmp, row.names = FALSE)

  expect_error(
    merge_regions(tmp, 1),
    "Unsupported file format"
  )
})

test_that("read_sheet_raw validates inputs", {
  expect_error(
    read_sheet_raw("/nonexistent/file.xlsx", 1),
    "File does not exist"
  )

  tmp <- tempfile(fileext = ".csv")
  on.exit(unlink(tmp), add = TRUE)
  write.csv(mtcars, tmp, row.names = FALSE)

  expect_error(
    read_sheet_raw(tmp, 1),
    "Unsupported file format"
  )
})

test_that("sheet validation shows available sheets in error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  # Error message should include available sheet names

  expect_error(
    read_excel(tmp, sheet = "BadName"),
    "Available sheets:"
  )

  # Error message for index should show count and names

  expect_error(
    read_excel(tmp, sheet = 5),
    "File has 1 sheet"
  )
})
