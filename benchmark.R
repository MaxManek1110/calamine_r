# Benchmark: calaminer vs readxlsb for .xlsb files
# Install readxlsb if needed: install.packages("readxlsb")

library(calaminer)

file_path <- "inst/extdata/random_table_40mb.xlsb"

cat("Benchmarking:", file_path, "\n\n")

# Benchmark calaminer
cat("calaminer::read_excel_calamine()\n")
time_calaminer <- system.time({
  df_calaminer <- calaminer::read_excel_calamine(file_path)
})
sheets <- calaminer::excel_sheets_calamine(file_path)
dims <- calaminer::sheet_dims_calamine(file_path, sheet = 1)
range <- paste0("A1:", LETTERS[dims[2][[1]]], dims[1][[1]])
cat("  Time:", time_calaminer["elapsed"], "seconds\n")
cat("  Rows:", nrow(df_calaminer), " Cols:", ncol(df_calaminer), "\n\n")

# Benchmark readxlsb
if (requireNamespace("readxlsb", quietly = TRUE)) {
  cat("readxlsb::read_xlsb()\n")
  time_readxlsb <- system.time({
    df_readxlsb <- readxlsb::read_xlsb(
      file_path,
      sheet = sheets[1],
      range = range
    )
  })
  cat("  Time:", time_readxlsb["elapsed"], "seconds\n")
  cat("  Rows:", nrow(df_readxlsb), " Cols:", ncol(df_readxlsb), "\n\n")

  # Summary
  cat("--- Summary ---\n")
  cat("calaminer:", round(time_calaminer["elapsed"], 3), "s\n")
  cat("readxlsb: ", round(time_readxlsb["elapsed"], 3), "s\n")
  cat(
    "Speedup:  ",
    round(time_readxlsb["elapsed"] / time_calaminer["elapsed"], 1),
    "x faster\n"
  )
} else {
  cat("readxlsb not installed. Run: install.packages('readxlsb')\n")
}


library(data.table)

compare_dt <- function(x, y) {
  x <- as.data.table(x)
  y <- as.data.table(y)

  list(
    rows_equal = nrow(x) == nrow(y),
    cols_equal = setequal(names(x), names(y)),
    value_diff = fsetdiff(x, y),
    value_diff_reverse = fsetdiff(y, x),
    identical = identical(x, y)
  )
}

df1 <- as.data.table(df_calaminer)
df2 <- as.data.table(df_readxlsb)
compare_dt(df1, df2)
