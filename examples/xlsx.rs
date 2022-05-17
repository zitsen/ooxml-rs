use ooxml::document::SpreadsheetDocument;
use std::env;

fn main() {
    let input = env::args()
        .into_iter()
        .skip(1)
        .next()
        .unwrap_or("examples/simple-spreadsheet/data-image-demo.xlsx".to_string());
    println!("input xlsx: {input}");
    let xlsx = SpreadsheetDocument::open(input).unwrap();

    let workbook = xlsx.get_workbook();
    //println!("{:?}", xlsx);

    let _sheet_names = workbook.worksheet_names();

    for (sheet_idx, sheet) in workbook.worksheets().iter().enumerate() {
        println!("worksheet {}", sheet_idx);
        println!("worksheet dimension: {:?}", sheet.dimenstion());
        println!("---------DATA---------");
        for rows in sheet.rows() {
            // get cell values
            let cols: Vec<_> = rows
                .into_iter()
                .map(|cell| cell.value().unwrap_or_default())
                .collect();
            println!("{}", itertools::join(&cols, ","));
        }
        println!("----------------------");
    }
}
