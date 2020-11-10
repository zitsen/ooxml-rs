use ooxml::document::SpreadsheetDocument;

fn main() {
    let xlsx =
        SpreadsheetDocument::open("examples/simple-spreadsheet/data-image-demo.xlsx").unwrap();

    let workbook = xlsx.get_workbook();
    //println!("{:?}", xlsx);

    let _sheet_names = workbook.worksheet_names();

    for (sheet_idx, sheet) in workbook.worksheets().iter().enumerate() {
        println!("worksheet {}", sheet_idx);
        println!("worksheet dimension: {:?}", sheet.dimenstion());
        println!("---------DATA---------");
        for rows in sheet.rows() {
            let cols: Vec<String> = rows
                .into_iter()
                .map(|cell| cell.to_string().unwrap_or_default())
                .collect();
            // use iertools::join or write to csv.
            println!(
                "{}",
                cols.iter().fold(String::new(), |mut l, c| {
                    if l.is_empty() {
                        l.push_str(c);
                    } else {
                        l.push(',');
                        l.push_str(c)
                    }
                    l
                })
            );
        }
        println!("----------------------");
    }
}
