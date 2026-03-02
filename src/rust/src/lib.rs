use extendr_api::prelude::*;
use calamine::{open_workbook_auto, Reader, Data};

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

    let mut rows: Vec<Vec<String>> = range.rows()
        .skip(skip as usize)
        .map(|row| row.iter().map(|c| cell_to_string(c)).collect())
        .collect();

    if rows.is_empty() {
        return Ok(List::new(0));
    }

    let ncols = rows.iter().map(|r| r.len()).max().unwrap_or(0);

    // Pad rows to equal length
    for row in &mut rows {
        while row.len() < ncols {
            row.push(String::new());
        }
    }

    // Extract column names
    let col_names_vec: Vec<String> = if col_names && !rows.is_empty() {
        let names = rows.remove(0);
        names.into_iter()
            .enumerate()
            .map(|(i, n)| if n.is_empty() { format!("V{}", i + 1) } else { n })
            .collect()
    } else {
        (1..=ncols).map(|i| format!("V{}", i)).collect()
    };

    // Build columns
    let mut df = List::new(ncols);
    for col_idx in 0..ncols {
        let col_data: Vec<String> = rows.iter()
            .map(|row| row.get(col_idx).cloned().unwrap_or_default())
            .collect();
        df.set_elt(col_idx, col_data.into_robj())?;
    }

    // Set names and class
    df.set_names(col_names_vec)?;
    df.set_class(&["data.frame"])?;
    df.set_attrib(
        "row.names",
        (1..=rows.len() as i32).collect::<Vec<i32>>()
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

// Macro to generate R exports
extendr_module! {
    mod calamine_r;
    fn cal_sheet_names;
    fn cal_read_sheet;
    fn cal_read_sheet_df;
    fn cal_sheet_dims;
}
