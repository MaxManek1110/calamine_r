// We need to forward routine registration from C to Rust
// to avoid the overhead of looking up symbols in runtime.
// See https://extendr.github.io/rextendr/articles/package.html

#include <R.h>
#include <Rinternals.h>
#include <R_ext/Rdynload.h>

// Rust functions
SEXP wrap__cal_sheet_names(SEXP path);
SEXP wrap__cal_read_sheet(SEXP path, SEXP sheet);
SEXP wrap__cal_read_sheet_df(SEXP path, SEXP sheet, SEXP col_names, SEXP skip);
SEXP wrap__cal_sheet_dims(SEXP path, SEXP sheet);

static const R_CallMethodDef CallEntries[] = {
    {"wrap__cal_sheet_names", (DL_FUNC) &wrap__cal_sheet_names, 1},
    {"wrap__cal_read_sheet", (DL_FUNC) &wrap__cal_read_sheet, 2},
    {"wrap__cal_read_sheet_df", (DL_FUNC) &wrap__cal_read_sheet_df, 4},
    {"wrap__cal_sheet_dims", (DL_FUNC) &wrap__cal_sheet_dims, 2},
    {NULL, NULL, 0}
};

void R_init_calaminer(DllInfo *dll) {
    R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
    R_useDynamicSymbols(dll, FALSE);
}
