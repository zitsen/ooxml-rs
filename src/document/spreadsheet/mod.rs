//! Excel file format .xlsx document implementation.

use std::{borrow::Cow, cell::RefCell};
use std::{path::Path, rc::Rc};

// use derivative::Derivative;
use derivative::Derivative;

use crate::{
    error::Result,
    packaging::element::*,
    packaging::package::{OpenXmlPackage, Relationships},
};

mod cell;
mod chart;
mod document_type;
mod drawing;
mod media;
mod shared_string;
mod style;
mod workbook;
mod worksheet;

use self::{
    document_type::SpreadsheetDocumentType, style::CellStyleComponent, worksheet::SheetCol,
};

use self::cell::CellValue;
use self::shared_string::SharedStringsPart;
use self::style::StylesPart;
use self::workbook::WorkbookPart;
use self::worksheet::WorksheetPart;

#[derive(Derivative, Clone, Default)]
#[derivative(Debug)]
pub struct SpreadsheetParts {
    initialized: bool,
    #[derivative(Debug = "ignore")]
    pub package: Rc<RefCell<OpenXmlPackage>>,
    pub relationships: Relationships,
    pub workbook: WorkbookPart,
    pub styles: StylesPart,
    pub shared_strings: SharedStringsPart,
    // pub media: Vec<MediaPart>,
    // pub drawings: Vec<DrawingPart>,
    // pub charts: Vec<ChartPart>,
    /// Dict for worksheets, key is uri, value is worksheet part.
    pub worksheets: linked_hash_map::LinkedHashMap<String, WorksheetPart>,
}

impl SpreadsheetParts {
    pub fn from_package(package: Rc<RefCell<OpenXmlPackage>>) -> Self {
        // let parts = package.borrow();
        // let part = parts.get_part("xl/_rels/workbook.xml.rels").unwrap();
        let relationships = {
            let package = package.borrow();
            let part = package.get_part("xl/_rels/workbook.xml.rels").unwrap();
            Relationships::parse_from_xml_reader(part.as_part_bytes())
        };
        let workbook = {
            let package = package.borrow();
            let part = package.get_part("xl/workbook.xml").unwrap();
            WorkbookPart::from_xml_reader(part.as_part_bytes()).expect("workbook part error")
        };
        let shared_strings = {
            let package = package.borrow();
            let part = package.get_part("xl/sharedStrings.xml").unwrap();
            SharedStringsPart::from_xml_reader(part.as_part_bytes()).expect("workbook part error")
        };
        let styles = {
            let package = package.borrow();
            let part = package.get_part("xl/styles.xml").unwrap();
            StylesPart::from_xml_reader(part.as_part_bytes()).expect("workbook part error")
        };
        let mut this = Self {
            package: package,
            relationships,
            workbook,
            shared_strings,
            styles,
            initialized: true,
            ..Default::default()
        };
        this.parse_worksheets();
        this
    }

    pub fn get_shared_string(&self, idx: usize) -> Option<&str> {
        self.shared_strings.get_shared_string(idx)
    }

    pub fn get_cell_style<'a>(&'a self, id: usize) -> Option<CellStyleComponent<'a>> {
        self.styles.get_cell_style_component(id)
    }
    pub fn get_worksheet_part<T: AsRef<str>>(&self, uri: T) -> Option<&WorksheetPart> {
        self.worksheets.get(uri.as_ref())
    }

    fn parse_worksheets(&mut self) {
        // Parse sheet data by relationship target.
        for sheet in &self.workbook.sheets.sheets {
            let relationship = self
                .relationships
                .get_relationship_by_id(&sheet.r_id)
                .expect("the worksheet relationship doest not exist");
            let worksheet_uri = relationship.target();
            // println!("{}", worksheet_uri);
            let package = self.package.borrow();
            let part = package
                .get_part(&format!("xl/{}", worksheet_uri))
                .expect("get worksheet part by uri");
            let sheet =
                WorksheetPart::from_xml_reader(part.as_part_bytes()).expect("parse worksheet");

            self.worksheets.insert(worksheet_uri.into(), sheet);
        }
    }
}
#[derive(Derivative, Default)]
#[derivative(Debug)]
pub struct SpreadsheetDocument {
    /// The OpenXML package itself.
    #[derivative(Debug = "ignore")]
    package: Rc<RefCell<OpenXmlPackage>>,
    /// The spreadsheet `OpenXML Part` collection.
    parts: Rc<RefCell<SpreadsheetParts>>,
    /// The spreadsheet document type, eg. .xlsx, .xlsm, etc.
    document_type: SpreadsheetDocumentType,
    /// Workbook
    workbook: Workbook,
}

#[derive(Derivative)]
#[derivative(Debug, Clone)]
struct Row {
    cells: Vec<CellValue>,
}

#[derive(Derivative)]
#[derivative(Debug, Clone)]
pub struct Worksheet {
    #[derivative(Debug = "ignore")]
    parts: Rc<RefCell<SpreadsheetParts>>,
    name: String,
    sheet_id: usize,
    part: WorksheetPart,
}

impl Worksheet {
    pub fn dimenstion(&self) -> Option<(usize, usize)> {
        self.part.dimension()
    }
    pub fn get_row_size(&self) -> usize {
        self.dimenstion().unwrap_or_default().0
    }
    pub fn get_col_size(&self) -> usize {
        self.dimenstion().unwrap_or_default().1
    }
    pub fn get_shared_string(&self, idx: usize) -> Option<String> {
        let parts = self.parts.as_ref().borrow();
        parts.get_shared_string(idx).map(|s| s.into())
    }
    pub fn get_cell_style(&self, id: usize) {
        let parts = self.parts.as_ref().borrow();
        let cs = parts.get_cell_style(id);
        let cs = cs.unwrap();
        let nf = cs.number_format();
        unimplemented!()
    }
    /// Format a cell's raw value with given cell style id.
    pub fn format_cell_with(&self, raw: &str, style_id: usize) -> Option<String> {
        let parts = self.parts.as_ref().borrow();
        let cs = parts.get_cell_style(style_id);
        let cs = cs.unwrap();
        let nf = cs.number_format();
        let nf = nf.unwrap();
        let code = nf.code.as_str();
        let s = match code {
            "yyyy/m/d;@" => {
                // Excel stores dates as sequential serial numbers so that they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,447 days after January 1, 1900.
                // source: https://support.office.com/en-us/article/datevalue-function-df8b07d4-7761-4a93-bc33-b7471bbff252

                // Here's calamine as_date()
                //let days = raw.parse().expect("date time string") - 25569;
                //let secs = days * 86400;
                //chrono::NaiveDateTime::from_timestamp_opt(secs, 0);

                // use 1899.12.30 instead of 1900.1.1 to fix bug in excel date.
                //let d1900 = chrono::NaiveDate::from_ymd(1900, 1, 1);
                let d1900 = chrono::NaiveDate::from_ymd(1899, 12, 30);
                //println!("{}", d1900);
                let d3 =
                    d1900 + chrono::Duration::days(raw.parse().expect("not a valid date string"));
                format!("{}", d3)
            }
            _ => unimplemented!(),
        };
        Some(s)
    }
    pub fn get_cell_type(&self, idx: usize) {
        unimplemented!()
    }
}

#[derive(Debug)]
pub struct RowsIter<'a> {
    sheet: &'a Worksheet,
    row: usize,
    col: usize,
}

#[derive(Debug)]
pub struct RowIter<'a> {
    sheet: &'a Worksheet,
    row: usize,
    col: usize,
}

impl<'a> RowsIter<'a> {
    fn row_iter(&self) -> RowIter<'a> {
        RowIter {
            sheet: self.sheet,
            row: self.row,
            col: self.col,
        }
    }
}

impl<'a> Iterator for RowsIter<'a> {
    type Item = RowIter<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.row >= self.sheet.get_row_size() {
            return None;
        };
        let row = self.row_iter();
        self.row += 1;
        Some(row)
    }
}

pub struct Cell<'a> {
    sheet: &'a Worksheet,
    row: usize,
    col: usize,
}

impl<'a> Cell<'a> {
    fn inner(&self) -> Option<&SheetCol> {
        let data = self.sheet.part.sheet_data.as_ref().unwrap();
        data.rows
            .as_ref()
            .and_then(|rows| rows.get(self.row))
            .and_then(|row| row.cols.get(self.col))
    }
    pub fn cell_type(&self) {
        unimplemented!()
    }

    pub fn is_merged_cell(&self) -> bool {
        unimplemented!()
    }
    pub fn cell_value(&self) -> Option<CellValue> {
        self.inner().map(|cell| cell.raw_value())
    }

    pub fn to_string(&self) -> Option<String> {
        let inner = self.inner();
        if inner.is_none() {
            return None;
        }
        let inner = inner.unwrap();
        let ctype = inner.cell_type();
        let value = match ctype {
            cell::CellType::Empty => "".to_string(),
            cell::CellType::Raw => inner.raw_value().to_string(),
            cell::CellType::Shared(shared_string_id) => self
                .sheet
                .get_shared_string(shared_string_id)
                .expect("shared string not found"),
            cell::CellType::Styled(style_id) => {
                // let style = self.sheet.get_cell_style(style_id);
                let s = self
                    .sheet
                    .format_cell_with(&inner.v, style_id)
                    .expect("format with cell style");
                //return s;
                //unimplemented!()
                s
            }
        };
        Some(value)
    }
    pub fn cell_style(&self) {}
    pub fn cell_number_format(&self) {}
}
impl<'a> Iterator for RowIter<'a> {
    type Item = Cell<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.col >= self.sheet.get_col_size() {
            return None;
        };
        let cell = Cell {
            sheet: self.sheet,
            row: self.row,
            col: self.col,
        };
        self.col += 1;
        Some(cell)
    }
}

impl Worksheet {
    pub fn rows<'a>(&'a self) -> RowsIter<'a> {
        RowsIter {
            sheet: self,
            row: 0,
            col: 0,
        }
    }
}
#[derive(Derivative, Default)]
#[derivative(Debug, Clone)]
pub struct Workbook {
    #[derivative(Debug = "ignore")]
    parts: Rc<RefCell<SpreadsheetParts>>,
    worksheets: Vec<Worksheet>,
}

impl Workbook {
    pub fn new(parts: impl Into<Rc<RefCell<SpreadsheetParts>>>) -> Self {
        // parse worksheets from spreadsheet parts.
        let parts = parts.into();
        let borrowed_parts = parts.borrow();
        let mut worksheets = Vec::new();

        // Parse sheet data by relationship target.
        for sheet in &borrowed_parts.workbook.sheets.sheets {
            let relationship = borrowed_parts
                .relationships
                .get_relationship_by_id(&sheet.r_id)
                .expect("the worksheet relationship doest not exist");
            let worksheet_uri = relationship.target();

            let part = borrowed_parts.get_worksheet_part(&worksheet_uri).unwrap();
            // println!("{:?}", part);
            let worksheet = Worksheet {
                parts: parts.clone(),
                name: sheet.name.clone(),
                sheet_id: sheet.sheet_id,
                part: part.clone(),
            };
            worksheets.push(worksheet);
        }
        Self {
            parts: parts.clone(),
            worksheets,
        }
    }

    /// Get the worksheet names by it loading order.
    pub fn worksheet_names(&self) -> Vec<String> {
        self.parts
            .borrow()
            .workbook
            .sheet_names()
            .into_iter()
            .map(|v| v.clone())
            .collect()
    }

    /// Immutable worksheets slice
    pub fn worksheets(&self) -> &[Worksheet] {
        self.worksheets.as_slice()
    }
    /// Mutable worksheets slice
    pub fn worksheets_mut(&mut self) -> &[Worksheet] {
        self.worksheets.as_mut_slice()
    }

    /// Add a worksheet.
    pub fn add_worksheet(&mut self, _name: &str) -> &mut Worksheet {
        unimplemented!()
        // let sheet = Worksheet {
        //     parts: self.parts.clone(),
        // };
        // self.worksheets.push(sheet);
        // self.worksheets.last_mut().unwrap()
    }
}

/// Parse or create new spreadsheet document.
///
///
/// ```rust
/// use ooxml::document::SpreadsheetDocument;
/// fn main() {
///     let xlsx = SpreadsheetDocument::open("examples/simple-spreadsheet/data-image-demo.xlsx")
///         .expect("open excel file");
///     let workbook = xlsx.get_workbook();
///     println!("{:?}", workbook);
/// }
/// ```
impl SpreadsheetDocument {
    pub fn create(document_type: SpreadsheetDocumentType) -> Self {
        Self {
            document_type,
            ..Default::default()
        }
    }

    /// Open existing spreadsheet file and parse.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let package = OpenXmlPackage::open(path)?;
        let package = Rc::new(RefCell::new(package));
        let parts = SpreadsheetParts::from_package(package.clone());
        let parts = Rc::new(RefCell::new(parts));
        let workbook = Workbook::new(parts.clone());
        // let content_type = "";
        // let document_type = SpreadsheetDocumentType::from_content_type(content_type);
        let document_type = SpreadsheetDocumentType::Workbook;
        Ok(Self {
            package,
            parts,
            workbook,
            document_type,
        })
    }
    /// Save as new file with `path`.
    pub fn save_as<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.package.borrow().save_as(path)?;
        Ok(())
    }

    pub fn add_workbook(&mut self) -> Workbook {
        Workbook::new(self.parts.clone())
    }

    /// Serialize all parts to package.
    pub fn flush(&self) {
        unimplemented!()
    }
    /// Get workbook
    pub fn get_workbook(&self) -> &Workbook {
        &self.workbook
    }
    /// Get workbook
    pub fn get_workbook_mut(&mut self) -> &mut Workbook {
        unimplemented!()
    }
}

#[test]
fn open() {
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

#[test]
fn chrono() {
    let fmt = "yyyy/m/d";
    let v = 29567;
    let date = chrono::Duration::days(v);
    let date2 = chrono::NaiveDate::from_ymd(1980, 12, 12);
    let date3 = date2 - date;
    println!("{}", date3);
    let d1900 = chrono::NaiveDate::from_ymd(1900, 1, 1);
    println!("{}", d1900);
    let d3 = d1900 + chrono::Duration::days(v);
    println!("{}", d3);
}
