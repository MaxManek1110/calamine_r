use extendr_api::prelude::*;
use calamine::{open_workbook_auto, Reader, Data, SheetType};
use std::io::{Read as IoRead, BufReader, Cursor};
use std::fs::File;
use std::path::Path;

// Supported file extensions
const SUPPORTED_EXTENSIONS: &[&str] = &["xlsx", "xlsm", "xlsb", "xls", "ods"];

/// Validate that the file path exists and has a supported extension
fn validate_path(path: &str) -> std::result::Result<(), String> {
    // Check if file exists
    if !Path::new(path).exists() {
        return Err(format!("File not found: '{}'", path));
    }

    // Check file extension
    let lower_path = path.to_lowercase();
    let has_valid_ext = SUPPORTED_EXTENSIONS.iter().any(|ext| lower_path.ends_with(&format!(".{}", ext)));

    if !has_valid_ext {
        let ext = Path::new(path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("(none)");
        return Err(format!(
            "Unsupported file format: '{}'. Supported formats: {}",
            ext,
            SUPPORTED_EXTENSIONS.join(", ")
        ));
    }

    Ok(())
}

/// Validate that the sheet argument is a single value (not a vector)
fn validate_sheet_arg(sheet: &Robj) -> std::result::Result<(), String> {
    let len = sheet.len();

    if len == 0 {
        return Err("Sheet argument cannot be empty".to_string());
    }

    if len > 1 {
        return Err(format!(
            "Sheet must be a single value, not a vector of length {}. Use sheet = 1 or sheet = \"SheetName\"",
            len
        ));
    }

    // Check type - must be string, integer, or numeric
    let is_valid_type = sheet.is_string()
        || sheet.is_integer()
        || sheet.is_real()
        || sheet.as_str().is_some();

    if !is_valid_type {
        return Err(format!(
            "Sheet must be a character string or numeric index, got type: {:?}",
            sheet.rtype()
        ));
    }

    Ok(())
}

// Merge region structure
#[derive(Debug, Clone)]
struct MergeRegion {
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
}

// Convert column letters to 0-based index (A=0, B=1, ..., Z=25, AA=26, ...)
fn col_to_idx(col: &str) -> u32 {
    let mut result: u32 = 0;
    for c in col.chars() {
        let val = (c.to_ascii_uppercase() as u32) - ('A' as u32) + 1;
        result = result * 26 + val;
    }
    result.saturating_sub(1)
}

// Parse cell reference like "A1" to (row, col) 0-based indices
fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let col_end = cell_ref.chars().position(|c| c.is_ascii_digit())?;
    let col_str = &cell_ref[..col_end];
    let row_str = &cell_ref[col_end..];

    let col = col_to_idx(col_str);
    let row: u32 = row_str.parse::<u32>().ok()?.saturating_sub(1); // Convert to 0-based

    Some((row, col))
}

// Parse cell range like "A1:C3" to MergeRegion
fn parse_cell_range(ref_str: &str) -> Option<MergeRegion> {
    let parts: Vec<&str> = ref_str.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let (start_row, start_col) = parse_cell_ref(parts[0])?;
    let (end_row, end_col) = parse_cell_ref(parts[1])?;

    Some(MergeRegion {
        start_row,
        start_col,
        end_row,
        end_col,
    })
}

// Get xlsx/xlsm sheet path from sheet name by parsing workbook.xml and relationships
fn get_xlsx_sheet_path(archive: &mut zip::ZipArchive<BufReader<File>>, sheet_name: &str) -> Option<String> {
    // Parse workbook.xml to get sheet rId
    let mut workbook_xml = String::new();
    {
        let mut file = archive.by_name("xl/workbook.xml").ok()?;
        file.read_to_string(&mut workbook_xml).ok()?;
    }

    // Find the sheet element with matching name and get its r:id
    // Simple parsing - look for <sheet name="..." sheetId="..." r:id="..."/>
    let mut r_id: Option<String> = None;
    for line in workbook_xml.split('<') {
        if line.starts_with("sheet ") && line.contains(&format!("name=\"{}\"", sheet_name)) {
            // Extract r:id
            if let Some(idx) = line.find("r:id=\"") {
                let start = idx + 6;
                if let Some(end) = line[start..].find('"') {
                    r_id = Some(line[start..start + end].to_string());
                    break;
                }
            }
        }
    }

    let r_id = r_id?;

    // Parse workbook.xml.rels to get the target path
    let mut rels_xml = String::new();
    {
        let mut file = archive.by_name("xl/_rels/workbook.xml.rels").ok()?;
        file.read_to_string(&mut rels_xml).ok()?;
    }

    // Find relationship with matching Id
    for line in rels_xml.split('<') {
        if line.starts_with("Relationship ") && line.contains(&format!("Id=\"{}\"", r_id)) {
            // Extract Target
            if let Some(idx) = line.find("Target=\"") {
                let start = idx + 8;
                if let Some(end) = line[start..].find('"') {
                    let target = &line[start..start + end];
                    // Prepend xl/ if not absolute
                    return Some(if target.starts_with('/') {
                        target[1..].to_string()
                    } else {
                        format!("xl/{}", target)
                    });
                }
            }
        }
    }

    None
}

// Get merge regions from xlsx/xlsm file
fn get_xlsx_merge_regions(path: &str, sheet_name: &str) -> Vec<MergeRegion> {
    let mut regions = Vec::new();

    let file = match File::open(path) {
        Ok(f) => f,
        Err(_) => return regions,
    };

    let mut archive = match zip::ZipArchive::new(BufReader::new(file)) {
        Ok(a) => a,
        Err(_) => return regions,
    };

    // Get the sheet path
    let sheet_path = match get_xlsx_sheet_path(&mut archive, sheet_name) {
        Some(p) => p,
        None => return regions,
    };

    // Read the sheet XML
    let mut sheet_xml = String::new();
    {
        match archive.by_name(&sheet_path) {
            Ok(mut f) => {
                let _ = f.read_to_string(&mut sheet_xml);
            }
            Err(_) => return regions,
        };
    }

    // Parse mergeCells - look for <mergeCell ref="A1:B2"/>
    for part in sheet_xml.split("<mergeCell ") {
        if let Some(idx) = part.find("ref=\"") {
            let start = idx + 5;
            if let Some(end) = part[start..].find('"') {
                let ref_str = &part[start..start + end];
                if let Some(region) = parse_cell_range(ref_str) {
                    regions.push(region);
                }
            }
        }
    }

    regions
}

// Get xlsb sheet path from sheet name by parsing workbook.bin.rels and workbook.bin
fn get_xlsb_sheet_path(archive: &mut zip::ZipArchive<BufReader<File>>, sheet_name: &str, sheet_idx: usize) -> Option<String> {
    // First try to parse workbook.bin to find the sheet rId
    // workbook.bin contains BrtBundleSh records (type 0x009C) with sheet info
    let mut workbook_bin = Vec::new();
    {
        let mut file = archive.by_name("xl/workbook.bin").ok()?;
        file.read_to_end(&mut workbook_bin).ok()?;
    }

    // Parse workbook.bin to find sheet names and their rIds
    let mut sheet_r_ids: Vec<(String, String)> = Vec::new(); // (name, rId)
    let mut cursor = Cursor::new(&workbook_bin);

    loop {
        let rec_type = match read_xlsb_record_type(&mut cursor) {
            Some(t) => t,
            None => break,
        };
        let rec_size = match read_xlsb_record_size(&mut cursor) {
            Some(s) => s,
            None => break,
        };

        let pos = cursor.position() as usize;
        if pos + rec_size > workbook_bin.len() {
            break;
        }

        // BrtBundleSh (0x009C) - sheet info record
        if rec_type == 0x009C && rec_size >= 8 {
            let data = &workbook_bin[pos..pos + rec_size];
            // Structure: hsState (4), iTabID (4), strRelID (XLNullableWideString), strName (XLWideString)
            // Skip hsState (4) and iTabID (4)
            let mut offset = 8;

            // Read strRelID (XLNullableWideString) - 4 byte length, then UTF-16LE chars
            if offset + 4 > rec_size {
                cursor.set_position((pos + rec_size) as u64);
                continue;
            }
            let rel_id_len = u32::from_le_bytes([data[offset], data[offset+1], data[offset+2], data[offset+3]]) as usize;
            offset += 4;

            let rel_id = if rel_id_len > 0 && rel_id_len < 1000 && offset + rel_id_len * 2 <= rec_size {
                let mut s = String::new();
                for i in 0..rel_id_len {
                    let c = u16::from_le_bytes([data[offset + i*2], data[offset + i*2 + 1]]);
                    if let Some(ch) = char::from_u32(c as u32) {
                        s.push(ch);
                    }
                }
                offset += rel_id_len * 2;
                s
            } else {
                cursor.set_position((pos + rec_size) as u64);
                continue;
            };

            // Read strName (XLWideString) - 4 byte length, then UTF-16LE chars
            if offset + 4 > rec_size {
                cursor.set_position((pos + rec_size) as u64);
                continue;
            }
            let name_len = u32::from_le_bytes([data[offset], data[offset+1], data[offset+2], data[offset+3]]) as usize;
            offset += 4;

            let name = if name_len > 0 && name_len < 1000 && offset + name_len * 2 <= rec_size {
                let mut s = String::new();
                for i in 0..name_len {
                    let c = u16::from_le_bytes([data[offset + i*2], data[offset + i*2 + 1]]);
                    if let Some(ch) = char::from_u32(c as u32) {
                        s.push(ch);
                    }
                }
                s
            } else {
                cursor.set_position((pos + rec_size) as u64);
                continue;
            };

            sheet_r_ids.push((name, rel_id));
        }

        cursor.set_position((pos + rec_size) as u64);
    }

    // Find the rId for the target sheet
    let r_id = sheet_r_ids.iter()
        .find(|(name, _)| name == sheet_name)
        .map(|(_, rid)| rid.clone())
        .or_else(|| sheet_r_ids.get(sheet_idx).map(|(_, rid)| rid.clone()))?;

    // Parse workbook.bin.rels to get the target path
    let mut rels_xml = String::new();
    {
        let mut file = archive.by_name("xl/_rels/workbook.bin.rels").ok()?;
        file.read_to_string(&mut rels_xml).ok()?;
    }

    // Find relationship with matching Id
    for line in rels_xml.split('<') {
        if line.starts_with("Relationship ") && line.contains(&format!("Id=\"{}\"", r_id)) {
            if let Some(idx) = line.find("Target=\"") {
                let start = idx + 8;
                if let Some(end) = line[start..].find('"') {
                    let target = &line[start..start + end];
                    return Some(if target.starts_with('/') {
                        target[1..].to_string()
                    } else {
                        format!("xl/{}", target)
                    });
                }
            }
        }
    }

    // Fallback: try common path pattern
    Some(format!("xl/worksheets/sheet{}.bin", sheet_idx + 1))
}

// Get merge regions from xlsb file (binary records)
fn get_xlsb_merge_regions(path: &str, sheet_name: &str, sheet_idx: usize) -> Vec<MergeRegion> {
    let mut regions = Vec::new();

    let file = match File::open(path) {
        Ok(f) => f,
        Err(_) => return regions,
    };

    let mut archive = match zip::ZipArchive::new(BufReader::new(file)) {
        Ok(a) => a,
        Err(_) => return regions,
    };

    // Find the sheet path for xlsb
    let sheet_path = match get_xlsb_sheet_path(&mut archive, sheet_name, sheet_idx) {
        Some(p) => p,
        None => return regions,
    };

    // Read the sheet binary
    let mut sheet_bin = Vec::new();
    {
        match archive.by_name(&sheet_path) {
            Ok(mut f) => {
                let _ = f.read_to_end(&mut sheet_bin);
            }
            Err(_) => return regions,
        };
    }

    // Parse binary records looking for BrtMergeCell (type 0x00B0)
    // XLSB record format: type (variable), size (variable), data
    let mut cursor = Cursor::new(&sheet_bin);

    loop {
        // Read record type (variable length integer)
        let rec_type = match read_xlsb_record_type(&mut cursor) {
            Some(t) => t,
            None => break,
        };

        // Read record size (variable length integer)
        let rec_size = match read_xlsb_record_size(&mut cursor) {
            Some(s) => s,
            None => break,
        };

        let pos = cursor.position() as usize;
        if pos + rec_size > sheet_bin.len() {
            break;
        }

        if rec_type == 0x00B0 {
            // BrtMergeCell: rwFirst (4), rwLast (4), colFirst (4), colLast (4)
            if rec_size >= 16 {
                let data = &sheet_bin[pos..pos + rec_size];
                let rw_first = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
                let rw_last = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
                let col_first = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
                let col_last = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);

                regions.push(MergeRegion {
                    start_row: rw_first,
                    start_col: col_first,
                    end_row: rw_last,
                    end_col: col_last,
                });
            }
        }

        cursor.set_position((pos + rec_size) as u64);
    }

    regions
}

// Read XLSB variable-length record type
fn read_xlsb_record_type(cursor: &mut Cursor<&Vec<u8>>) -> Option<u16> {
    let data = cursor.get_ref();
    let pos = cursor.position() as usize;
    if pos >= data.len() {
        return None;
    }

    let b1 = data[pos];
    if b1 & 0x80 == 0 {
        cursor.set_position((pos + 1) as u64);
        Some(b1 as u16)
    } else {
        if pos + 1 >= data.len() {
            return None;
        }
        let b2 = data[pos + 1];
        cursor.set_position((pos + 2) as u64);
        Some(((b2 as u16) << 7) | ((b1 & 0x7F) as u16))
    }
}

// Read XLSB variable-length record size
fn read_xlsb_record_size(cursor: &mut Cursor<&Vec<u8>>) -> Option<usize> {
    let data = cursor.get_ref();
    let mut pos = cursor.position() as usize;
    if pos >= data.len() {
        return None;
    }

    let mut size: usize = 0;
    let mut shift = 0;

    for i in 0..4 {
        if pos >= data.len() {
            return None;
        }
        let b = data[pos];
        pos += 1;
        size |= ((b & 0x7F) as usize) << shift;
        shift += 7;
        if b & 0x80 == 0 {
            break;
        }
        if i == 3 {
            // 4th byte should not have continuation bit
            break;
        }
    }

    cursor.set_position(pos as u64);
    Some(size)
}

// Get merge regions from xls file (BIFF format in CFB)
fn get_xls_merge_regions(path: &str, sheet_idx: usize) -> Vec<MergeRegion> {
    let mut regions = Vec::new();

    // Read the file
    let mut file_data = Vec::new();
    match File::open(path) {
        Ok(mut f) => {
            if f.read_to_end(&mut file_data).is_err() {
                return regions;
            }
        }
        Err(_) => return regions,
    };

    // Parse CFB (Compound File Binary) to find the Workbook stream
    // This is a simplified parser - looks for the workbook stream directly
    let workbook_data = match extract_cfb_workbook(&file_data) {
        Some(data) => data,
        None => return regions,
    };

    // Parse BIFF records to find MERGECELLS records (type 0x00E5)
    // We need to track which sheet we're in using BOUNDSHEET/BOF records
    let mut pos = 0;
    let mut current_sheet: i32 = -1;
    let mut in_target_sheet = false;

    while pos + 4 <= workbook_data.len() {
        let rec_type = u16::from_le_bytes([workbook_data[pos], workbook_data[pos + 1]]);
        let rec_size = u16::from_le_bytes([workbook_data[pos + 2], workbook_data[pos + 3]]) as usize;

        if pos + 4 + rec_size > workbook_data.len() {
            break;
        }

        let rec_data = &workbook_data[pos + 4..pos + 4 + rec_size];

        match rec_type {
            0x0809 => {
                // BOF - Beginning of File (sheet start)
                if rec_size >= 2 {
                    let bof_type = u16::from_le_bytes([rec_data[0], rec_data[1]]);
                    if bof_type == 0x0010 {
                        // Worksheet
                        current_sheet += 1;
                        in_target_sheet = current_sheet == sheet_idx as i32;
                    }
                }
            }
            0x000A => {
                // EOF - End of File (sheet end)
                if in_target_sheet {
                    // Done with target sheet
                    break;
                }
            }
            0x00E5 => {
                // MERGECELLS
                if in_target_sheet && rec_size >= 2 {
                    let count = u16::from_le_bytes([rec_data[0], rec_data[1]]) as usize;
                    let mut offset = 2;
                    for _ in 0..count {
                        if offset + 8 > rec_size {
                            break;
                        }
                        let rw_first = u16::from_le_bytes([rec_data[offset], rec_data[offset + 1]]) as u32;
                        let rw_last = u16::from_le_bytes([rec_data[offset + 2], rec_data[offset + 3]]) as u32;
                        let col_first = u16::from_le_bytes([rec_data[offset + 4], rec_data[offset + 5]]) as u32;
                        let col_last = u16::from_le_bytes([rec_data[offset + 6], rec_data[offset + 7]]) as u32;

                        regions.push(MergeRegion {
                            start_row: rw_first,
                            start_col: col_first,
                            end_row: rw_last,
                            end_col: col_last,
                        });

                        offset += 8;
                    }
                }
            }
            _ => {}
        }

        pos += 4 + rec_size;
    }

    regions
}

// Extract workbook stream from CFB (Compound File Binary)
fn extract_cfb_workbook(data: &[u8]) -> Option<Vec<u8>> {
    // CFB header check
    if data.len() < 512 {
        return None;
    }

    // Magic number check
    if &data[0..8] != &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
        return None;
    }

    // Get sector size
    let sector_shift = u16::from_le_bytes([data[30], data[31]]);
    let sector_size = 1usize << sector_shift;

    // Get mini sector size
    let mini_sector_shift = u16::from_le_bytes([data[32], data[33]]);
    let _mini_sector_size = 1usize << mini_sector_shift;

    // Get FAT sector count and locations
    let fat_sectors = u32::from_le_bytes([data[44], data[45], data[46], data[47]]) as usize;
    let first_dir_sector = u32::from_le_bytes([data[48], data[49], data[50], data[51]]);
    let _first_mini_fat_sector = u32::from_le_bytes([data[60], data[61], data[62], data[63]]);
    let mini_stream_cutoff = u32::from_le_bytes([data[56], data[57], data[58], data[59]]) as usize;

    // Read FAT
    let mut fat = Vec::new();
    let mut difat: Vec<u32> = Vec::new();

    // First 109 DIFAT entries are in header
    for i in 0..109.min(fat_sectors) {
        let offset = 76 + i * 4;
        if offset + 4 <= data.len() {
            let sector = u32::from_le_bytes([data[offset], data[offset + 1], data[offset + 2], data[offset + 3]]);
            if sector != 0xFFFFFFFE {
                difat.push(sector);
            }
        }
    }

    // Read FAT sectors
    for &sector in &difat {
        let offset = 512 + (sector as usize) * sector_size;
        if offset + sector_size <= data.len() {
            for i in (0..sector_size).step_by(4) {
                let entry = u32::from_le_bytes([
                    data[offset + i],
                    data[offset + i + 1],
                    data[offset + i + 2],
                    data[offset + i + 3],
                ]);
                fat.push(entry);
            }
        }
    }

    // Read directory entries
    let mut dir_entries = Vec::new();
    let mut dir_sector = first_dir_sector;

    while dir_sector != 0xFFFFFFFE && dir_sector != 0xFFFFFFFF {
        let offset = 512 + (dir_sector as usize) * sector_size;
        if offset + sector_size > data.len() {
            break;
        }

        // Each directory entry is 128 bytes
        for i in (0..sector_size).step_by(128) {
            let entry_offset = offset + i;
            if entry_offset + 128 <= data.len() {
                dir_entries.push(&data[entry_offset..entry_offset + 128]);
            }
        }

        // Get next directory sector from FAT
        if (dir_sector as usize) < fat.len() {
            dir_sector = fat[dir_sector as usize];
        } else {
            break;
        }
    }

    // Find Workbook entry (or Book for older formats)
    let mut workbook_entry: Option<&[u8]> = None;
    for entry in &dir_entries {
        let name_len = u16::from_le_bytes([entry[64], entry[65]]) as usize;
        if name_len >= 2 {
            // Convert UTF-16LE name
            let mut name = String::new();
            for i in (0..name_len - 2).step_by(2) {
                let c = u16::from_le_bytes([entry[i], entry[i + 1]]);
                if c != 0 {
                    if let Some(ch) = char::from_u32(c as u32) {
                        name.push(ch);
                    }
                }
            }
            if name == "Workbook" || name == "Book" {
                workbook_entry = Some(*entry);
                break;
            }
        }
    }

    let entry = workbook_entry?;

    // Get stream info
    let start_sector = u32::from_le_bytes([entry[116], entry[117], entry[118], entry[119]]);
    let stream_size = u32::from_le_bytes([entry[120], entry[121], entry[122], entry[123]]) as usize;

    // Read the workbook stream
    let mut result = Vec::with_capacity(stream_size);

    if stream_size < mini_stream_cutoff {
        // Mini stream - need to read from mini FAT
        // For simplicity, skip mini stream support - most workbooks are larger
        return None;
    }

    // Regular sectors
    let mut sector = start_sector;
    while result.len() < stream_size && sector != 0xFFFFFFFE && sector != 0xFFFFFFFF {
        let offset = 512 + (sector as usize) * sector_size;
        if offset + sector_size > data.len() {
            break;
        }

        let bytes_to_read = sector_size.min(stream_size - result.len());
        result.extend_from_slice(&data[offset..offset + bytes_to_read]);

        if (sector as usize) < fat.len() {
            sector = fat[sector as usize];
        } else {
            break;
        }
    }

    Some(result)
}

// Get merge regions based on file format
fn get_merge_regions(path: &str, sheet_name: &str, sheet_idx: usize) -> Vec<MergeRegion> {
    let lower_path = path.to_lowercase();

    if lower_path.ends_with(".xlsx") || lower_path.ends_with(".xlsm") {
        get_xlsx_merge_regions(path, sheet_name)
    } else if lower_path.ends_with(".xlsb") {
        get_xlsb_merge_regions(path, sheet_name, sheet_idx)
    } else if lower_path.ends_with(".xls") {
        get_xls_merge_regions(path, sheet_idx)
    } else {
        // ODS format - not implemented yet, return empty
        Vec::new()
    }
}

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
        // Validate that the sheet name exists
        if !sheet_names.contains(&name.to_string()) {
            let available = if sheet_names.len() <= 5 {
                sheet_names.join(", ")
            } else {
                format!("{}, ... ({} total)", sheet_names[..3].join(", "), sheet_names.len())
            };
            return Err(format!(
                "Sheet '{}' not found. Available sheets: {}",
                name, available
            ));
        }
        return Ok(name.to_string());
    }

    // Try integer first, then double (R often passes numbers as doubles)
    let idx = sheet.as_integer()
        .or_else(|| sheet.as_real().map(|f| f as i32))
        .ok_or_else(|| "Sheet must be a string name or numeric index".to_string())?;

    if idx < 1 {
        return Err(format!("Sheet index must be >= 1, got {}", idx));
    }

    let idx_usize = (idx - 1) as usize; // Convert to 0-based
    if idx_usize >= sheet_names.len() {
        return Err(format!(
            "Sheet index {} out of range. Workbook has {} sheet(s)",
            idx, sheet_names.len()
        ));
    }

    Ok(sheet_names[idx_usize].clone())
}

/// Get sheet names from an Excel file
/// @param path Path to the Excel file
/// @return Character vector of sheet names
/// @export
#[extendr]
fn cal_sheet_names(path: &str) -> Result<Vec<String>> {
    cal_sheet_names_inner(path)
}

fn cal_sheet_names_inner(path: &str) -> Result<Vec<String>> {
    validate_path(path).map_err(|e| Error::Other(e))?;

    let workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    Ok(workbook.sheet_names().to_vec())
}

/// Get sheet metadata (name, type, visibility) from an Excel file
/// @param path Path to the Excel file
/// @return A data.frame with columns: name, type, visible
/// @export
#[extendr]
fn cal_sheets_metadata(path: &str) -> Result<List> {
    cal_sheets_metadata_inner(path)
}

fn cal_sheets_metadata_inner(path: &str) -> Result<List> {
    validate_path(path).map_err(|e| Error::Other(e))?;

    let workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    let sheets = workbook.sheets_metadata();
    let nrows = sheets.len();

    let names: Vec<String> = sheets.iter().map(|s| s.name.clone()).collect();
    let types: Vec<String> = sheets.iter().map(|s| {
        match s.typ {
            SheetType::WorkSheet => "worksheet".to_string(),
            SheetType::ChartSheet => "chartsheet".to_string(),
            SheetType::DialogSheet => "dialogsheet".to_string(),
            SheetType::MacroSheet => "macrosheet".to_string(),
            SheetType::Vba => "vba".to_string(),
        }
    }).collect();
    let visible: Vec<bool> = sheets.iter().map(|s| {
        matches!(s.visible, calamine::SheetVisible::Visible)
    }).collect();

    let mut df = List::new(3);
    df.set_elt(0, names.into_robj())?;
    df.set_elt(1, types.into_robj())?;
    df.set_elt(2, visible.into_robj())?;

    df.set_names(["name", "type", "visible"])?;
    df.set_class(&["data.frame"])?;
    df.set_attrib("row.names", (1..=nrows as i32).collect::<Vec<i32>>())?;

    Ok(df)
}

/// Check if a sheet is a worksheet (not a chart/dialog/macro sheet)
/// @param path Path to the Excel file
/// @param sheet Sheet name or index (1-based)
/// @return Logical: TRUE if worksheet, FALSE otherwise
/// @export
#[extendr]
fn cal_is_worksheet(path: &str, sheet: Robj) -> Result<bool> {
    cal_is_worksheet_inner(path, &sheet)
}

fn cal_is_worksheet_inner(path: &str, sheet: &Robj) -> Result<bool> {
    validate_path(path).map_err(|e| Error::Other(e))?;
    validate_sheet_arg(sheet).map_err(|e| Error::Other(e))?;

    let workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    let sheet_names: Vec<String> = workbook.sheet_names().to_vec();
    let sheet_name = get_sheet_name(sheet, &sheet_names)
        .map_err(|e| Error::Other(e))?;

    let sheets = workbook.sheets_metadata();
    let sheet_meta = sheets.iter().find(|s| s.name == sheet_name);

    match sheet_meta {
        Some(s) => Ok(matches!(s.typ, SheetType::WorkSheet)),
        None => Ok(false),
    }
}

/// Read a sheet from an Excel file as a list of rows
/// @param path Path to the Excel file
/// @param sheet Sheet name or index (1-based)
/// @return List of character vectors (rows)
/// @export
#[extendr]
fn cal_read_sheet(path: &str, sheet: Robj) -> Result<List> {
    cal_read_sheet_inner(path, &sheet)
}

fn cal_read_sheet_inner(path: &str, sheet: &Robj) -> Result<List> {
    validate_path(path).map_err(|e| Error::Other(e))?;
    validate_sheet_arg(sheet).map_err(|e| Error::Other(e))?;

    let mut workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    let sheet_names = workbook.sheet_names().to_vec();
    let sheet_name = get_sheet_name(sheet, &sheet_names)
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
/// @param fill_merged Fill merged cells with the top-left value
/// @return A data.frame
/// @export
#[extendr]
fn cal_read_sheet_df(path: &str, sheet: Robj, col_names: bool, skip: i32, fill_merged: bool) -> Result<List> {
    cal_read_sheet_df_inner(path, &sheet, col_names, skip, fill_merged)
}

fn cal_read_sheet_df_inner(path: &str, sheet: &Robj, col_names: bool, skip: i32, fill_merged: bool) -> Result<List> {
    validate_path(path).map_err(|e| Error::Other(e))?;
    validate_sheet_arg(sheet).map_err(|e| Error::Other(e))?;

    if skip < 0 {
        return Err(Error::Other("skip must be non-negative".to_string()));
    }

    let mut workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    let sheet_names = workbook.sheet_names().to_vec();
    let sheet_name = get_sheet_name(sheet, &sheet_names)
        .map_err(|e| Error::Other(e))?;

    // Get sheet index for xls format parsing
    let sheet_idx = sheet_names.iter().position(|n| n == &sheet_name).unwrap_or(0);

    let range = workbook.worksheet_range(&sheet_name)
        .map_err(|e| Error::Other(format!("Failed to read sheet: {}", e)))?;

    // Collect all rows as raw Data cells
    let mut all_rows: Vec<Vec<Data>> = range.rows()
        .skip(skip as usize)
        .map(|row| row.to_vec())
        .collect();

    // Apply merged cell filling if requested
    if fill_merged && !all_rows.is_empty() {
        let skip_u32 = skip as u32;
        let merge_regions = get_merge_regions(path, &sheet_name, sheet_idx);

        for region in merge_regions {
            // Adjust region for skip - if region starts before skip, adjust accordingly
            if region.end_row < skip_u32 {
                continue; // Region entirely before skipped rows
            }

            let adjusted_start_row = if region.start_row >= skip_u32 {
                (region.start_row - skip_u32) as usize
            } else {
                0 // Region started before skip, use first available row
            };
            let adjusted_end_row = (region.end_row - skip_u32) as usize;

            // Check bounds
            if adjusted_start_row >= all_rows.len() {
                continue;
            }
            let end_row_bounded = adjusted_end_row.min(all_rows.len() - 1);

            // Get the value from the top-left cell of the merge region (in original coordinates)
            let value_row_in_output = if region.start_row >= skip_u32 {
                (region.start_row - skip_u32) as usize
            } else {
                // The top-left cell was skipped; we can't get its value
                continue;
            };

            let start_col = region.start_col as usize;
            let end_col = region.end_col as usize;

            if value_row_in_output >= all_rows.len() {
                continue;
            }

            let source_value = all_rows[value_row_in_output]
                .get(start_col)
                .cloned()
                .unwrap_or(Data::Empty);

            // Only propagate non-empty values
            if matches!(source_value, Data::Empty) {
                continue;
            }

            // Fill all cells in the region with the value
            for row_idx in adjusted_start_row..=end_row_bounded {
                // Ensure row has enough columns
                let row = &mut all_rows[row_idx];
                while row.len() <= end_col {
                    row.push(Data::Empty);
                }

                for col_idx in start_col..=end_col {
                    // Skip the source cell itself
                    if row_idx == value_row_in_output && col_idx == start_col {
                        continue;
                    }
                    row[col_idx] = source_value.clone();
                }
            }
        }
    }

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
                let mut robj = doubles.into_robj();
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
    cal_sheet_dims_inner(path, &sheet)
}

fn cal_sheet_dims_inner(path: &str, sheet: &Robj) -> Result<Vec<i32>> {
    validate_path(path).map_err(|e| Error::Other(e))?;
    validate_sheet_arg(sheet).map_err(|e| Error::Other(e))?;

    let mut workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    let sheet_names = workbook.sheet_names().to_vec();
    let sheet_name = get_sheet_name(sheet, &sheet_names)
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

    // Validate month and day ranges
    if month < 1 || month > 12 || day < 1 || day > 31 {
        return Err(());
    }

    // Simple days calculation (not accounting for all edge cases)
    // Days from year 1 to target year
    let days_from_years = (year - 1970) * 365 + count_leap_years(1970, year);
    let days_in_month: [u32; 12] = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let mut days_from_months: u32 = 0;
    for m in 0..(month - 1) as usize {
        // Safe: month is validated to be 1-12, so m is 0-11
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

/// Get merged cell regions from an Excel file
/// @param path Path to the Excel file
/// @param sheet Sheet name or index (1-based)
/// @return A data.frame with columns: start_row, start_col, end_row, end_col (1-based)
/// @export
#[extendr]
fn cal_merge_regions(path: &str, sheet: Robj) -> Result<List> {
    cal_merge_regions_inner(path, &sheet)
}

fn cal_merge_regions_inner(path: &str, sheet: &Robj) -> Result<List> {
    validate_path(path).map_err(|e| Error::Other(e))?;
    validate_sheet_arg(sheet).map_err(|e| Error::Other(e))?;

    let workbook = open_workbook_auto(path)
        .map_err(|e| Error::Other(format!("Failed to open workbook: {}", e)))?;

    let sheet_names = workbook.sheet_names().to_vec();
    let sheet_name = get_sheet_name(sheet, &sheet_names)
        .map_err(|e| Error::Other(e))?;

    let sheet_idx = sheet_names.iter().position(|n| n == &sheet_name).unwrap_or(0);

    let regions = get_merge_regions(path, &sheet_name, sheet_idx);

    let nrows = regions.len();

    // Build data.frame with 1-based indices for R
    let start_rows: Vec<i32> = regions.iter().map(|r| (r.start_row + 1) as i32).collect();
    let start_cols: Vec<i32> = regions.iter().map(|r| (r.start_col + 1) as i32).collect();
    let end_rows: Vec<i32> = regions.iter().map(|r| (r.end_row + 1) as i32).collect();
    let end_cols: Vec<i32> = regions.iter().map(|r| (r.end_col + 1) as i32).collect();

    let mut df = List::new(4);
    df.set_elt(0, start_rows.into_robj())?;
    df.set_elt(1, start_cols.into_robj())?;
    df.set_elt(2, end_rows.into_robj())?;
    df.set_elt(3, end_cols.into_robj())?;

    df.set_names(["start_row", "start_col", "end_row", "end_col"])?;
    df.set_class(&["data.frame"])?;
    df.set_attrib("row.names", (1..=nrows as i32).collect::<Vec<i32>>())?;

    Ok(df)
}

// Macro to generate R exports
extendr_module! {
    mod calamine_r;
    fn cal_sheet_names;
    fn cal_sheets_metadata;
    fn cal_is_worksheet;
    fn cal_read_sheet;
    fn cal_read_sheet_df;
    fn cal_sheet_dims;
    fn cal_merge_regions;
}
