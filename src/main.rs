use calamine::{open_workbook_auto, Reader};
use rayon::iter::{IntoParallelRefIterator, ParallelIterator};
use std::error::Error;
use std::fs::{self, File};
use std::io::{self, BufWriter, Write};
use std::path::PathBuf;


fn main() {
    let excel_files = fs::read_dir(".").unwrap()
        .map(|entry| entry.unwrap().path())
        .filter(|file_name| !file_name.starts_with("~"))
        .filter(|file_name| {
            if let Some(ext) = file_name.extension() {
                match ext.to_str().unwrap() {
                    "xlsx" | "xls" => return true,
                    _ => return false
                }
            } else {
                return false
            }
        });
    excel_files.collect::<Vec<PathBuf>>()
        .par_iter()
        .for_each(|file_name| { 
            process_file(file_name.as_os_str().to_str().unwrap()).unwrap();
        })
}


fn process_file(file_name: &str) -> Result<(), Box<dyn Error>> {
    let mut workbook = open_workbook_auto(file_name)?;
    let number_lists = workbook.sheet_names().len();
    for list_name in workbook.sheet_names() {
        if let Ok(r) = workbook.worksheet_range(&list_name) {
            let mut writer = if number_lists > 1 {
                make_file_writer(&file_name, Some(&list_name))?
            } else {
                make_file_writer(&file_name, None)?
            };
            for row in r.rows() {
                let line = row
                    .iter()
                    .map(|d| d.to_string())
                    .collect::<Vec<String>>()
                    .join(";") 
                    + "\n";
                writer.write(&line.into_bytes().to_owned()).unwrap();
            }
            writer.flush().unwrap();
        }
    }
    Ok(())
}


fn make_file_writer(file_name_workbook: &str, list_name: Option<&str>) -> Result<BufWriter<File>, io::Error> {
    let file_name_writer_base = file_name_workbook.rsplit_once(".").unwrap().0;
    let file_name_writer = if let Some(list_name) = list_name {
        format!("{file_name_writer_base}#{list_name}.csv")
    } else {
        format!("{file_name_writer_base}.csv")
    };
    let f = fs::File::create(file_name_writer)?;
    Ok(BufWriter::new(f))
}
