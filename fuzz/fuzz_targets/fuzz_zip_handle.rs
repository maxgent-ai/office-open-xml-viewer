#![no_main]

use libfuzzer_sys::fuzz_target;
use ooxml_common::resource::OoxmlFormat;

// Exercises the XLSX free extraction facade's shared validated ZIP reader and
// resource governor. DOCX, XLSX, and PPTX session paths are covered by their
// full-parser targets. Raw ZIP inputs (such as the scheduled XLSX seed) use a
// stable workbook part name and remain intact. Other inputs with a newline use
// the prefix as a co-evolving part name and the suffix as the candidate archive.
fuzz_target!(|data: &[u8]| {
    let (path, zip_bytes) = if data.starts_with(b"PK") {
        ("xl/workbook.xml".into(), data)
    } else {
        match data.iter().position(|&byte| byte == b'\n') {
            Some(split_at) => {
                let (path_bytes, rest) = data.split_at(split_at);
                (String::from_utf8_lossy(path_bytes), &rest[1..])
            }
            None => ("xl/workbook.xml".into(), data),
        }
    };

    // Exercise defaults and each public limit independently. A total limit can
    // bind before the entry limit for this single-entry operation.
    for (max_entry_bytes, max_total_bytes) in [(None, None), (Some(4096), None), (None, Some(2048))]
    {
        let _ = ooxml_common::zip::extract_zip_entry(
            zip_bytes,
            &path,
            OoxmlFormat::Xlsx,
            max_entry_bytes,
            max_total_bytes,
        );
    }
});
