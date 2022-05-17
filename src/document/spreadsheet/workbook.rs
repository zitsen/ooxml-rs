use std::borrow::Cow;

use crate::packaging::element::*;
use crate::packaging::namespace::Namespaces;

use quick_xml::events::attributes::Attribute;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "fileVersion")]
pub struct FileVersion {
    pub app_name: Option<String>,
    pub last_edited: Option<usize>,
    pub lowest_edited: Option<usize>,
    pub run_build: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "workbookPr")]
pub struct WorkbookPr {
    date1904: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "workbookView")]
pub struct WorkbookView {
    pub window_width: Option<usize>,
    pub window_height: Option<usize>,
    pub active_tab: Option<usize>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "bookViews")]
pub struct BookViews {
    pub workbook_view: WorkbookView,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheet")]
pub struct Sheet {
    pub name: String,
    pub sheet_id: usize,
    #[serde(rename = "r:id")]
    pub r_id: String,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheets")]
pub struct Sheets {
    #[serde(rename = "sheet")]
    pub sheets: Vec<Sheet>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "calcPr")]
pub struct CalcPr {
    calc_id: Option<String>,
}
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookPart {
    pub file_version: Option<FileVersion>,
    pub book_views: BookViews,
    pub workbook_pr: WorkbookPr,
    pub sheets: Sheets,
    pub calc_pr: Option<CalcPr>,
    #[serde(flatten)]
    namespaces: Namespaces,
}

impl WorkbookPart {
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect()
    }
}
impl OpenXmlElementInfo for WorkbookPart {
    fn tag_name() -> &'static str {
        "workbook"
    }
}

impl OpenXmlDeserializeDefault for WorkbookPart {}

impl OpenXmlSerialize for WorkbookPart {
    fn namespaces(&self) -> Option<Cow<Namespaces>> {
        Some(Cow::Borrowed(&self.namespaces))
    }
    fn attributes(&self) -> Option<Vec<Attribute>> {
        None
    }
    fn write_inner<W: std::io::Write>(&self, writer: W) -> Result<(), crate::error::OoxmlError> {
        let mut xml = quick_xml::Writer::new(writer);
        // use quick_xml::events::*;

        quick_xml::se::to_writer(xml.inner(), &self.file_version)?;
        quick_xml::se::to_writer(xml.inner(), &self.book_views)?;
        quick_xml::se::to_writer(xml.inner(), &self.workbook_pr)?;
        quick_xml::se::to_writer(xml.inner(), &self.sheets)?;
        quick_xml::se::to_writer(xml.inner(), &self.calc_pr)?;
        Ok(())
    }
}

#[test]
fn serde() {
    let workbook =
        WorkbookPart::from_xml_file("examples/simple-spreadsheet/xl/workbook.xml").unwrap();
    println!("{:?}", workbook);
    let str = workbook.to_xml_string().unwrap();
    println!("{}", str);

    let _sheet_names = workbook.sheet_names();
    let workbook2: WorkbookPart = quick_xml::de::from_str(&str).unwrap();
    println!("{:?}", workbook2);
    assert_eq!(workbook, workbook2);
}
