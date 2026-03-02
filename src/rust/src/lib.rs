use extendr_api::prelude::*;
use calamine::{open_workbook_auto, Reader, Data};

// Column type detection
#[derive(Clone, Copy, PartialEq, Debug)]
enum ColType {
    Unknown,
    Numeric,
    Logical,
    DateTime,
    String,
}

fn detect_cell_type(cell: &Data) -> ColType {
    match cell {
        Data::Empty => ColType::Unknown,
        Data::Int(_) | Data::Float(_) => ColType::Numeric,
        Data::Bool(_) => ColType::Logical,
        Data::DateTime(_) | Data::DateTimeIso(_) => ColType::DateTime,
        Data::String(_) | Data::Error(_) | Data::DurationIso(_) => ColType::String,
    }
}

fn merge_types(current: ColType, new: ColType) -> ColType {
    match (current, new) {
        (_, ColType::Unknown) => current,
        (ColType::Unknown, _) => new,
        (a, b) if a == b => a,
        // Numeric and DateTime can coexist (Excel stores dates as numbers)
        (ColType::Numeric, ColType::DateTime) | (ColType::DateTime, ColType::Numeric) => ColType::DateTime,
        _ => ColType::String, // Mixed types fall back to string
    }
}

// Excel serial date to R Date (days since 1970-01-01)
// Excel uses days since 1899-12-30, with a bug treating 1900 as leap year
const EXCEL_EPOCH_OFFSET: f64 = 25569.0; // Days from 1899-12-30 to 1970-01-01

fn excel_to_r_date(serial: f64) -> f64 {
    serial - EXCEL_EPOCH_OFFSET
}

// Helper to get sheet name from Robj (string or numeric index)
fn get_sheet_name(sheet: &Robj, sheet_names: &[String]) -> std::result::Result<String, String> {
    if let Some(name) = sheet.as_str() {
        return Ok(name.to_string());
    }

    // Try integer first, then double (R often passes numbers as doubles)
    let idx = sheet.as_integer()
        .or_else(|| sheet.as_real().map(|f| f as i32))
        .ok_or_else(|| "Sheet must be a string name or numeric index".to_string())?;

    let idx = (idx - 1) as usize; // Convert to 0-based
    if idx >= sheet_names.len() {
        return Err(format!("Sheet index {} out of range (max: {})", idx + 1, sheet_names.len()));
    }

    Ok(sheet_names[idx].clone())
}

/// Get sheet names from an Excel file
/// @param path Path to the Excel file
/// @return Character vector of sheet names
/// @export
#[extendr]
fn cal_sheet_names(path: &str) -> Result<Vec<String>> {
    let workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    Ok(workbook.sheet_names().to_vec())
}

/// Read a sheet from an Excel file as a list of rows
/// @param path Path to the Excel file
/// @param sheet Sheet name or index (1-based)
/// @return List of character vectors (rows)
/// @export
#[extendr]
fn cal_read_sheet(path: &str, sheet: Robj) -> Result<List> {
    let mut workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    let sheet_names = workbook.sheet_names().to_vec();
    let sheet_name = get_sheet_name(&sheet, &sheet_names)
        .map_err(|e| Error::Other(e))?;

    let range = workbook.worksheet_range(&sheet_name)
        .map_err(|e| Error::Other(format!("Failed to read sheet: {}", e)))?;

    // Convert to list of character vectors
    let rows: Vec<Robj> = range.rows()
        .map(|row| {
            let cells: Vec<String> = row.iter()
                .map(|cell| cell_to_string(cell))
                .collect();
            cells.into_robj()
        })
        .collect();

    Ok(List::from_values(rows))
}

/// Read a sheet as a data.frame
/// @param path Path to the Excel file
/// @param sheet Sheet name or index (1-based)
/// @param col_names Use first row as column names
/// @param skip Number of rows to skip before reading
/// @return A data.frame
/// @export
#[extendr]
fn cal_read_sheet_df(path: &str, sheet: Robj, col_names: bool, skip: i32) -> Result<List> {
    let mut workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    let sheet_names = workbook.sheet_names().to_vec();
    let sheet_name = get_sheet_name(&sheet, &sheet_names)
        .map_err(|e| Error::Other(e))?;

    let range = workbook.worksheet_range(&sheet_name)
        .map_err(|e| Error::Other(format!("Failed to read sheet: {}", e)))?;

    // Collect all rows as raw Data cells
    let all_rows: Vec<Vec<Data>> = range.rows()
        .skip(skip as usize)
        .map(|row| row.to_vec())
        .collect();

    if all_rows.is_empty() {
        return Ok(List::new(0));
    }

    let ncols = all_rows.iter().map(|r| r.len()).max().unwrap_or(0);

    // Determine where data starts (after header row if col_names=true)
    let (header_row, data_rows): (Option<&Vec<Data>>, &[Vec<Data>]) = if col_names && !all_rows.is_empty() {
        (Some(&all_rows[0]), &all_rows[1..])
    } else {
        (None, &all_rows[..])
    };

    // Detect column types from data rows
    let mut col_types: Vec<ColType> = vec![ColType::Unknown; ncols];
    for row in data_rows {
        for (col_idx, cell) in row.iter().enumerate() {
            let cell_type = detect_cell_type(cell);
            col_types[col_idx] = merge_types(col_types[col_idx], cell_type);
        }
    }

    // Default unknown types to String
    for t in &mut col_types {
        if *t == ColType::Unknown {
            *t = ColType::String;
        }
    }

    // Extract column names
    let col_names_vec: Vec<String> = if let Some(header) = header_row {
        (0..ncols)
            .map(|i| {
                header.get(i)
                    .map(|c| cell_to_string(c))
                    .filter(|s| !s.is_empty())
                    .unwrap_or_else(|| format!("V{}", i + 1))
            })
            .collect()
    } else {
        (1..=ncols).map(|i| format!("V{}", i)).collect()
    };

    // Build typed columns
    let mut df = List::new(ncols);
    let nrows = data_rows.len();

    for col_idx in 0..ncols {
        let col_robj = match col_types[col_idx] {
            ColType::Numeric => {
                let mut doubles = Doubles::new(nrows);
                for (row_idx, row) in data_rows.iter().enumerate() {
                    let val = row.get(col_idx).unwrap_or(&Data::Empty);
                    doubles.set_elt(row_idx, cell_to_rfloat(val));
                }
                doubles.into_robj()
            }
            ColType::Logical => {
                let mut logicals = Logicals::new(nrows);
                for (row_idx, row) in data_rows.iter().enumerate() {
                    let val = row.get(col_idx).unwrap_or(&Data::Empty);
                    logicals.set_elt(row_idx, cell_to_rbool(val));
                }
                logicals.into_robj()
            }
            ColType::DateTime => {
                // Return as R Date class (numeric with class attribute)
                let mut doubles = Doubles::new(nrows);
                for (row_idx, row) in data_rows.iter().enumerate() {
                    let val = row.get(col_idx).unwrap_or(&Data::Empty);
                    doubles.set_elt(row_idx, cell_to_rdate(val));
                }
                let robj = doubles.into_robj();
                robj.set_class(&["Date"]).ok();
                robj
            }
            ColType::String | ColType::Unknown => {
                let mut strings = Strings::new(nrows);
                for (row_idx, row) in data_rows.iter().enumerate() {
                    let val = row.get(col_idx).unwrap_or(&Data::Empty);
                    strings.set_elt(row_idx, cell_to_rstr(val));
                }
                strings.into_robj()
            }
        };
        df.set_elt(col_idx, col_robj)?;
    }

    // Set names and class
    df.set_names(col_names_vec)?;
    df.set_class(&["data.frame"])?;
    df.set_attrib(
        "row.names",
        (1..=nrows as i32).collect::<Vec<i32>>()
    )?;

    Ok(df)
}

/// Read sheet dimensions (rows, cols)
/// @param path Path to the Excel file
/// @param sheet Sheet name or index (1-based)
/// @return Integer vector c(nrow, ncol)
/// @export
#[extendr]
fn cal_sheet_dims(path: &str, sheet: Robj) -> Result<Vec<i32>> {
    let mut workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    let sheet_names = workbook.sheet_names().to_vec();
    let sheet_name = get_sheet_name(&sheet, &sheet_names)
        .map_err(|e| Error::Other(e))?;

    let range = workbook.worksheet_range(&sheet_name)
        .map_err(|e| Error::Other(format!("Failed to read sheet: {}", e)))?;

    let (rows, cols) = range.get_size();
    Ok(vec![rows as i32, cols as i32])
}

// Helper function to convert calamine Data enum to String
fn cell_to_string(cell: &Data) -> String {
    match cell {
        Data::Empty => String::new(),
        Data::String(s) => s.clone(),
        Data::Int(i) => i.to_string(),
        Data::Float(f) => f.to_string(),
        Data::Bool(b) => if *b { "TRUE".into() } else { "FALSE".into() },
        Data::Error(e) => format!("#ERROR:{:?}", e),
        Data::DateTime(dt) => dt.to_string(),
        Data::DateTimeIso(s) => s.clone(),
        Data::DurationIso(s) => s.clone(),
    }
}

// Convert cell to R numeric (Rfloat)
fn cell_to_rfloat(cell: &Data) -> Rfloat {
    match cell {
        Data::Empty => Rfloat::na(),
        Data::Int(i) => Rfloat::from(*i as f64),
        Data::Float(f) => {
            if f.is_nan() {
                Rfloat::na()
            } else {
                Rfloat::from(*f)
            }
        }
        Data::DateTime(dt) => Rfloat::from(dt.as_f64()),
        Data::Bool(b) => Rfloat::from(if *b { 1.0 } else { 0.0 }),
        Data::String(s) => s.parse::<f64>().map(Rfloat::from).unwrap_or(Rfloat::na()),
        _ => Rfloat::na(),
    }
}

// Convert cell to R logical (Rbool)
fn cell_to_rbool(cell: &Data) -> Rbool {
    match cell {
        Data::Empty => Rbool::na(),
        Data::Bool(b) => Rbool::from(*b),
        Data::Int(i) => Rbool::from(*i != 0),
        Data::Float(f) => {
            if f.is_nan() {
                Rbool::na()
            } else {
                Rbool::from(*f != 0.0)
            }
        }
        Data::String(s) => {
            let lower = s.to_lowercase();
            match lower.as_str() {
                "true" | "yes" | "1" => Rbool::from(true),
                "false" | "no" | "0" => Rbool::from(false),
                _ => Rbool::na(),
            }
        }
        _ => Rbool::na(),
    }
}

// Convert cell to R Date (Rfloat with days since 1970-01-01)
fn cell_to_rdate(cell: &Data) -> Rfloat {
    match cell {
        Data::Empty => Rfloat::na(),
        Data::DateTime(dt) => Rfloat::from(excel_to_r_date(dt.as_f64())),
        Data::Int(i) => Rfloat::from(excel_to_r_date(*i as f64)),
        Data::Float(f) => {
            if f.is_nan() {
                Rfloat::na()
            } else {
                Rfloat::from(excel_to_r_date(*f))
            }
        }
        Data::DateTimeIso(s) => {
            // Try to parse ISO date string
            if let Some(date_part) = s.split('T').next() {
                if let Ok(days) = parse_iso_date_to_r_days(date_part) {
                    return Rfloat::from(days);
                }
            }
            Rfloat::na()
        }
        _ => Rfloat::na(),
    }
}

// Parse ISO date string (YYYY-MM-DD) to R days since 1970-01-01
fn parse_iso_date_to_r_days(s: &str) -> std::result::Result<f64, ()> {
    let parts: Vec<&str> = s.split('-').collect();
    if parts.len() != 3 {
        return Err(());
    }
    let year: i32 = parts[0].parse().map_err(|_| ())?;
    let month: u32 = parts[1].parse().map_err(|_| ())?;
    let day: u32 = parts[2].parse().map_err(|_| ())?;

    // Simple days calculation (not accounting for all edge cases)
    // Days from year 1 to target year
    let days_from_years = (year - 1970) * 365 + count_leap_years(1970, year);
    let days_in_month: [u32; 12] = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let mut days_from_months: u32 = 0;
    for m in 0..(month - 1) as usize {
        days_from_months += days_in_month[m];
    }
    // Add leap day if applicable
    if month > 2 && is_leap_year(year) {
        days_from_months += 1;
    }

    Ok(days_from_years as f64 + days_from_months as f64 + (day - 1) as f64)
}

fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}

fn count_leap_years(from: i32, to: i32) -> i32 {
    let count_up_to = |y: i32| -> i32 {
        let y = y - 1;
        y / 4 - y / 100 + y / 400
    };
    if to >= from {
        count_up_to(to) - count_up_to(from)
    } else {
        -(count_up_to(from) - count_up_to(to))
    }
}

// Convert cell to R string (Rstr)
fn cell_to_rstr(cell: &Data) -> Rstr {
    match cell {
        Data::Empty => Rstr::na(),
        _ => Rstr::from(cell_to_string(cell)),
    }
}

// Macro to generate R exports
extendr_module! {
    mod calamine_r;
    fn cal_sheet_names;
    fn cal_read_sheet;
    fn cal_read_sheet_df;
    fn cal_sheet_dims;
}
