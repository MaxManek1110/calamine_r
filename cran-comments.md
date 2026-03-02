# CRAN Submission Comments

## Test environments

* local Ubuntu 22.04, R 4.2.x, Rust 1.75.0
* GitHub Actions (ubuntu-latest, macOS-latest, windows-latest)
* R-hub

## R CMD check results

0 errors | 0 warnings | 3 notes

### Notes explanation:

1. **Hidden files and directories**: The `src/rust/vendor/` directory contains
   vendored Rust dependencies with `.cargo-checksum.json` files. These are
   required for offline builds as mandated by CRAN policy for packages with
   Rust compiled code. See [Using Rust in CRAN Packages](https://cran.r-project.org/web/packages/using_rust.html).

2. **Installed package size (~7Mb)**: This is typical for packages with Rust
   compiled code. The static library contains optimized, LTO-compiled Rust code.
   This is unavoidable for packages using Rust.

3. **Future file timestamps**: This is a transient check environment issue

## Downstream dependencies

This is a new package with no downstream dependencies.

## Notes for CRAN reviewers

This package wraps the Rust 'calamine' library for high-performance Excel file
reading. The package follows CRAN guidelines for Rust packages:

- All Rust dependencies are vendored for offline builds (no network access needed)
- Rust toolchain (cargo >= 1.75.0) is documented in SystemRequirements
- MSRV (Minimum Supported Rust Version) is checked in configure scripts
- Build is limited to 2 parallel jobs (CARGO_BUILD_JOBS=2)
- Rust and cargo versions are reported before building
- Builds successfully on Linux, macOS, and Windows
