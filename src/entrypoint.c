// We need to forward routine registration from C to Rust
// to avoid the overhead of looking up symbols in runtime.
// See https://extendr.github.io/rextendr/articles/package.html

#include <R.h>
#include <Rinternals.h>
#include <R_ext/Rdynload.h>

// Wrap abort() to prevent R termination - required for CRAN compliance
// Linked via -Wl,--wrap=abort, all abort() calls redirect here
void __wrap_abort(void) {
    Rf_error("fatal error in Rust code");
}

// Rust functions
SEXP wrap__cal_sheet_names(SEXP path);
SEXP wrap__cal_sheets_metadata(SEXP path);
SEXP wrap__cal_is_worksheet(SEXP path, SEXP sheet);
SEXP wrap__cal_read_sheet(SEXP path, SEXP sheet);
SEXP wrap__cal_read_sheet_df(SEXP path, SEXP sheet, SEXP col_names, SEXP skip, SEXP fill_merged);
SEXP wrap__cal_sheet_dims(SEXP path, SEXP sheet);
SEXP wrap__cal_merge_regions(SEXP path, SEXP sheet);

static const R_CallMethodDef CallEntries[] = {
    {"wrap__cal_sheet_names", (DL_FUNC) &wrap__cal_sheet_names, 1},
    {"wrap__cal_sheets_metadata", (DL_FUNC) &wrap__cal_sheets_metadata, 1},
    {"wrap__cal_is_worksheet", (DL_FUNC) &wrap__cal_is_worksheet, 2},
    {"wrap__cal_read_sheet", (DL_FUNC) &wrap__cal_read_sheet, 2},
    {"wrap__cal_read_sheet_df", (DL_FUNC) &wrap__cal_read_sheet_df, 5},
    {"wrap__cal_sheet_dims", (DL_FUNC) &wrap__cal_sheet_dims, 2},
    {"wrap__cal_merge_regions", (DL_FUNC) &wrap__cal_merge_regions, 2},
    {NULL, NULL, 0}
};

void R_init_calamineR(DllInfo *dll) {
    R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
    R_useDynamicSymbols(dll, FALSE);
}
