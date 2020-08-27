extern crate clap;
extern crate csv;

use calamine::{open_workbook, Reader, Xlsx};
use clap::{App, Arg};
use csv::WriterBuilder;
use std::error::Error;
use std::fs::{self, DirBuilder};
use std::path::{Path, PathBuf};

type MyResult<T> = Result<T, Box<dyn Error>>;

#[derive(Debug)]
pub struct Config {
    files: Vec<String>,
    outdir: String,
    delimiter: u8,
    normalize: bool,
    make_dirs: bool,
}

pub fn get_args() -> MyResult<Config> {
    let matches = App::new("excel2txt")
        .version("0.1.0")
        .author("Ken Youens-Clark <kyclark@gmail.com>")
        .about("Export Excel workbooks into delimited text files")
        .arg(
            Arg::with_name("file")
                .short("f")
                .long("file")
                .value_name("FILE")
                .help("File input")
                .required(true)
                .min_values(1),
        )
        .arg(
            Arg::with_name("outdir")
                .short("o")
                .long("outdir")
                .value_name("DIR")
                .default_value("out")
                .help("Output directory"),
        )
        .arg(
            Arg::with_name("delimiter")
                .short("d")
                .long("delimiter")
                .value_name("DELIM")
                .default_value("\t")
                .help("Delimiter for output files"),
        )
        .arg(
            Arg::with_name("normalize")
                .short("n")
                .long("normalize")
                .help("Normalize headers"),
        )
        .arg(
            Arg::with_name("make_dirs")
                .short("m")
                .long("mkdirs")
                .help("Make output directory for each input file"),
        )
        .get_matches();

    let files = matches.values_of_lossy("file").unwrap();

    let bad: Vec<String> =
        files.iter().cloned().filter(|f| !is_file(f)).collect();

    if !bad.is_empty() {
        let msg = format!(
            "Invalid file{}: {}",
            if bad.len() == 1 { "" } else { "s" },
            bad.join(", ")
        );
        return Err(From::from(msg));
    }

    Ok(Config {
        files: files,
        outdir: matches.value_of("outdir").unwrap().to_string(),
        delimiter: *matches
            .value_of("delimiter")
            .unwrap()
            .as_bytes()
            .first()
            .unwrap(),
        normalize: matches.is_present("normalize"),
        make_dirs: matches.is_present("make_dirs"),
    })
}

// --------------------------------------------------
pub fn run(config: Config) -> MyResult<()> {
    for (i, file) in config.files.into_iter().enumerate() {
        let path = Path::new(&file);
        let basename = path.file_stem().expect("basename");
        let stem = &basename.to_string_lossy().to_string();

        println!("{}: {}", i, basename.to_string_lossy());

        let mut out_dir = PathBuf::from(&config.outdir);
        if config.make_dirs {
            out_dir.push(stem)
        }
        if !out_dir.is_dir() {
            DirBuilder::new().recursive(true).create(&out_dir)?;
        }

        let mut excel: Xlsx<_> = open_workbook(file)?;
        let sheets = excel.sheet_names().to_owned();
        for sheet in sheets {
            let ext = if config.delimiter == 44 { "csv" } else { "txt" }
            let out_file = format!("{}__{}.{}", stem, sheet, ext);
            let out_path = &out_dir.join(out_file);
            let mut wtr = WriterBuilder::new()
                .delimiter(config.delimiter)
                .from_path(out_path)?;
            println!("  Sheet '{}' -> '{}'", sheet, out_path.display());
            if let Some(Ok(r)) = excel.worksheet_range(&sheet) {
                for row in r.rows() {
                    let vals = row
                        .into_iter()
                        .map(|f| format!("{}", f))
                        .collect::<Vec<String>>();
                    wtr.write_record(&vals)?;
                }
            }
            wtr.flush()?;
        }
    }

    Ok(())
}

// --------------------------------------------------
fn is_file(path: &String) -> bool {
    if let Ok(meta) = fs::metadata(path) {
        return meta.is_file();
    } else {
        return false;
    }
}
