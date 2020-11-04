//! Excel file format .xlsx document implementation.

use std::cell::RefCell;
use std::{path::Path, rc::Rc};

// use derivative::Derivative;
use derivative::Derivative;

use crate::{
    error::OoxmlError,
    packaging::package::{OpenXmlPackage, Relationships},
    packaging::xml::*,
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


use self::document_type::SpreadsheetDocumentType;


use self::shared_string::SharedStringsPart;
use self::style::StylesPart;
use self::workbook::WorkbookPart;
use self::worksheet::WorksheetPart;
use self::cell::CellValue;

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
    cells: Vec<CellValue>
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
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self, OoxmlError> {
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
    pub fn save_as<P: AsRef<Path>>(&self, path: P) -> Result<(), OoxmlError> {
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

    let _sheet_names = workbook.worksheet_names();
    println!("{:?}", xlsx);
}
