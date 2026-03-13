# Tests to verify Rust code returns R errors instead of panicking
# Panics in Rust will crash the R session; these tests verify graceful error handling

library(calamineR)

# Helper to create a fake Excel file (actually just garbage bytes)
create_corrupt_file <- function(ext) {

  tmp <- tempfile(fileext = ext)
  writeBin(charToRaw("this is not a valid excel file"), tmp)
  tmp
}

# Helper to create an empty file
create_empty_file <- function(ext) {
  tmp <- tempfile(fileext = ext)
  file.create(tmp)
  tmp
}

# Helper to create file with random bytes
create_random_bytes_file <- function(ext, size = 1000) {
  tmp <- tempfile(fileext = ext)
  writeBin(as.raw(sample(0:255, size, replace = TRUE)), tmp)
  tmp
}

# =============================================================================
# Corrupt/malformed file tests - these should error, not panic
# =============================================================================

test_that("corrupt xlsx file returns R error, not panic", {
  tmp <- create_corrupt_file(".xlsx")
  on.exit(unlink(tmp), add = TRUE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
  expect_error(sheet_dims(tmp, 1))
  expect_error(read_sheet_raw(tmp, 1))
  expect_error(merge_regions(tmp, 1))
})

test_that("corrupt xlsm file returns R error, not panic", {
  tmp <- create_corrupt_file(".xlsm")
  on.exit(unlink(tmp), add = TRUE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

test_that("corrupt xlsb file returns R error, not panic", {
  tmp <- create_corrupt_file(".xlsb")
  on.exit(unlink(tmp), add = TRUE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

test_that("corrupt xls file returns R error, not panic", {
  tmp <- create_corrupt_file(".xls")
  on.exit(unlink(tmp), add = TRUE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

test_that("corrupt ods file returns R error, not panic", {
  tmp <- create_corrupt_file(".ods")
  on.exit(unlink(tmp), add = TRUE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

# =============================================================================
# Empty file tests
# =============================================================================

test_that("empty xlsx file returns R error, not panic", {
  tmp <- create_empty_file(".xlsx")
  on.exit(unlink(tmp), add = TRUE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
  expect_error(sheet_dims(tmp, 1))
  expect_error(read_sheet_raw(tmp, 1))
  expect_error(merge_regions(tmp, 1))
})

test_that("empty xls file returns R error, not panic", {
  tmp <- create_empty_file(".xls")
  on.exit(unlink(tmp), add = TRUE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

test_that("empty xlsb file returns R error, not panic", {
  tmp <- create_empty_file(".xlsb")
  on.exit(unlink(tmp), add = TRUE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

# =============================================================================
# Random bytes file tests
# =============================================================================

test_that("random bytes xlsx file returns R error, not panic", {
  tmp <- create_random_bytes_file(".xlsx")
  on.exit(unlink(tmp), add = TRUE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

test_that("random bytes xls file returns R error, not panic", {
  tmp <- create_random_bytes_file(".xls")
  on.exit(unlink(tmp), add = TRUE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

# =============================================================================
# Edge case sheet indices - should error gracefully
# =============================================================================

test_that("very large sheet index returns R error, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  # Very large indices that could cause overflow
  expect_error(read_excel(tmp, sheet = .Machine$integer.max))
  expect_error(read_excel(tmp, sheet = 1e9))
  expect_error(sheet_dims(tmp, sheet = .Machine$integer.max))
  expect_error(read_sheet_raw(tmp, sheet = 1e9))
  expect_error(merge_regions(tmp, sheet = .Machine$integer.max))
})

test_that("zero sheet index returns R error, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(read_excel(tmp, sheet = 0))
  expect_error(sheet_dims(tmp, sheet = 0))
  expect_error(read_sheet_raw(tmp, sheet = 0))
  expect_error(merge_regions(tmp, sheet = 0))
})

test_that("negative sheet index returns R error, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(read_excel(tmp, sheet = -1))
  expect_error(read_excel(tmp, sheet = -100))
  expect_error(sheet_dims(tmp, sheet = -1))
  expect_error(read_sheet_raw(tmp, sheet = -1))
  expect_error(merge_regions(tmp, sheet = -1))
})

# =============================================================================
# Invalid type inputs - should error gracefully
# =============================================================================

test_that("NULL sheet returns R error, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(read_excel(tmp, sheet = NULL))
  expect_error(sheet_dims(tmp, sheet = NULL))
  expect_error(read_sheet_raw(tmp, sheet = NULL))
  expect_error(merge_regions(tmp, sheet = NULL))
})

test_that("NA sheet returns R error, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(read_excel(tmp, sheet = NA))
  expect_error(read_excel(tmp, sheet = NA_character_))
  expect_error(read_excel(tmp, sheet = NA_integer_))
  expect_error(sheet_dims(tmp, sheet = NA))
  expect_error(read_sheet_raw(tmp, sheet = NA))
  expect_error(merge_regions(tmp, sheet = NA))
})

test_that("empty string sheet name returns R error, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(read_excel(tmp, sheet = ""))
  expect_error(sheet_dims(tmp, sheet = ""))
  expect_error(read_sheet_raw(tmp, sheet = ""))
  expect_error(merge_regions(tmp, sheet = ""))
})

test_that("list as sheet returns R error, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(read_excel(tmp, sheet = list(1)))
  expect_error(sheet_dims(tmp, sheet = list("a")))
})

test_that("logical sheet returns R error, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(read_excel(tmp, sheet = TRUE))
  expect_error(read_excel(tmp, sheet = FALSE))
  expect_error(sheet_dims(tmp, sheet = TRUE))
})

# =============================================================================
# Invalid path inputs - should error gracefully
# =============================================================================
test_that("NULL path returns R error, not panic", {
  expect_error(excel_sheets(NULL))
  expect_error(read_excel(NULL))
  expect_error(sheet_dims(NULL, 1))
})

test_that("NA path returns R error, not panic", {
  expect_error(excel_sheets(NA))
  expect_error(read_excel(NA))
  expect_error(sheet_dims(NA, 1))
})

test_that("empty string path returns R error, not panic", {
  expect_error(excel_sheets(""))
  expect_error(read_excel(""))
  expect_error(sheet_dims("", 1))
})

test_that("vector of paths returns R error, not panic", {
  expect_error(excel_sheets(c("a.xlsx", "b.xlsx")))
  expect_error(read_excel(c("a.xlsx", "b.xlsx")))
})

# =============================================================================
# Skip parameter edge cases
# =============================================================================

test_that("very large skip value returns empty data, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)


  # Skip more rows than exist - should return empty list or df, but not panic
  result <- read_excel(tmp, skip = 1000)
  # Result is either empty list or empty data.frame - both acceptable

  expect_true(length(result) == 0 || (is.data.frame(result) && nrow(result) == 0))
})

test_that("skip equal to row count returns empty data, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)


  # Skip exactly all rows - should return empty but not panic
  result <- read_excel(tmp, skip = nrow(mtcars) + 1, col_names = FALSE)
  # Result is either empty list or empty data.frame - both acceptable
  expect_true(length(result) == 0 || (is.data.frame(result) && nrow(result) == 0))
})

test_that("negative skip returns R error, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(read_excel(tmp, skip = -1))
  expect_error(read_excel(tmp, skip = -100))
})

# =============================================================================
# Unicode and special characters in sheet names
# =============================================================================

test_that("non-existent unicode sheet name returns R error, not panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(read_excel(tmp, sheet = "\U0001F600"))
  expect_error(read_excel(tmp, sheet = "\u4e2d\u6587"))
  expect_error(read_excel(tmp, sheet = "caf\u00e9"))
})

test_that("sheet name with special chars returns R error if not found", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(read_excel(tmp, sheet = "Sheet\t1"))
  expect_error(read_excel(tmp, sheet = "Sheet\n1"))
})

# =============================================================================
# File renamed with wrong extension
# =============================================================================

test_that("text file renamed to xlsx returns R error, not panic", {
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writeLines("hello,world\n1,2\n3,4", tmp)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

test_that("csv file renamed to xlsx returns R error, not panic", {
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  write.csv(mtcars, tmp, row.names = FALSE)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

test_that("binary file renamed to xlsx returns R error, not panic", {
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  saveRDS(mtcars, tmp)

  expect_error(excel_sheets(tmp))
  expect_error(read_excel(tmp))
})

# =============================================================================
# Verify errors contain useful messages (not just "panic")
# =============================================================================

test_that("error messages are informative, not generic panic messages", {
  # Non-existent file
  err <- tryCatch(read_excel("/no/such/file.xlsx"), error = conditionMessage)
  expect_false(grepl("panic", err, ignore.case = TRUE))
  expect_true(grepl("exist|not exist|File does", err, ignore.case = TRUE))

  # Wrong format - create temp csv with xlsx extension
  tmp <- tempfile(fileext = ".csv")
  on.exit(unlink(tmp), add = TRUE)
  write.csv(mtcars, tmp, row.names = FALSE)

  err <- tryCatch(read_excel(tmp), error = conditionMessage)
  expect_false(grepl("panic", err, ignore.case = TRUE))
  expect_true(grepl("format|Unsupported|csv|xlsx", err, ignore.case = TRUE))
})

test_that("sheet not found error is informative", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  err <- tryCatch(read_excel(tmp, sheet = "NoSuchSheet"), error = conditionMessage)
  expect_false(grepl("panic", err, ignore.case = TRUE))
  expect_match(err, "not found", ignore.case = TRUE)
})

test_that("sheet index out of range error is informative", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  err <- tryCatch(read_excel(tmp, sheet = 99), error = conditionMessage)
  expect_false(grepl("panic", err, ignore.case = TRUE))
  expect_match(err, "out of range", ignore.case = TRUE)
})

# =============================================================================
# Concurrent/stress scenarios (basic)
# =============================================================================

test_that("rapid repeated calls with errors don't cause panic", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  # Rapid error calls

  for (i in 1:20) {
    expect_error(read_excel(tmp, sheet = "nonexistent"))
  }

  # Should still work after errors
  result <- read_excel(tmp, sheet = 1)
  expect_s3_class(result, "data.frame")
})

test_that("alternating valid and invalid calls work correctly", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  for (i in 1:10) {
    # Error call
    expect_error(read_excel(tmp, sheet = 999))
    # Valid call
    result <- read_excel(tmp, sheet = 1)
    expect_s3_class(result, "data.frame")
  }
})
