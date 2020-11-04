use super::cell::CellValue;
use crate::packaging::namespace::Namespaces;
use crate::packaging::xml::*;

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheetPr")]
pub struct SheetPr {}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheetView")]
pub struct SheetView {
    tab_selected: Option<usize>,
    workbook_view_id: Option<usize>,
    selection: Option<Selection>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheetViews")]
pub struct SheetViews {
    sheet_view: SheetView,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "calcPr")]
pub struct CalcPr {
    calc_id: String,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "dimension")]
pub struct Dimension {
    r#ref: String,
}
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "calcPr")]
pub struct Selection {
    active_cell: String,
    sqref: String,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheetFormatPr")]
pub struct SheetFormatPr {
    default_col_width: Option<f32>,
    default_row_height: Option<f32>,
    outline_level_row: Option<f32>,
    outline_level_col: Option<f32>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "v")]
pub struct SheetValue {
    v: CellValue,
}
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "c")]
pub struct SheetCol {
    r: String,
    t: Option<String>,
    s: Option<usize>,
    #[serde(rename = "$value")]
    v: String,
}
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "row")]
pub struct SheetRow {
    r: usize,
    spans: String,
    #[serde(rename = "c")]
    cols: Vec<SheetCol>,
}
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheetData")]
pub struct SheetData {
    #[serde(rename = "row")]
    rows: Option<Vec<SheetRow>>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "pageMargins")]
pub struct PageMargins {
    left: Option<f32>,
    right: Option<f32>,
    top: Option<f32>,
    bottom: Option<f32>,
    header: Option<f32>,
    footer: Option<f32>,
}
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "headerFooter")]
pub struct HeaderFooter {}
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "worksheet")]
pub struct WorksheetPart {
    #[serde(flatten)]
    namespaces: Namespaces,
    sheet_pr: SheetPr,
    dimension: Option<Dimension>,
    sheet_views: Option<SheetViews>,
    sheet_format_pr: Option<SheetFormatPr>,
    sheet_data: Option<SheetData>,
    page_margins: PageMargins,
    header_footer: HeaderFooter,
}

impl OpenXmlElementInfo for WorksheetPart {
    fn tag_name() -> &'static str {
        "worksheet"
    }
}

impl OpenXmlFromDeserialize for WorksheetPart {}

impl ToXml for WorksheetPart {
    fn write<W: std::io::Write>(&self, writer: W) -> Result<(), crate::error::OoxmlError> {
        let mut xml = quick_xml::Writer::new(writer);
        use quick_xml::events::*;

        // 1. write decl
        xml.write_event(Event::Decl(BytesDecl::new(
            b"1.0",
            Some(b"UTF-8"),
            Some(b"yes"),
        )))?;
        // quick_xml::se::to_writer(xml.inner(), self).unwrap();

        // 2. start types element
        let mut elem = BytesStart::borrowed_name(Self::tag_name().as_bytes());
        elem.extend_attributes(self.namespaces.to_xml_attributes());
        xml.write_event(Event::Start(elem))?;
        //xml.flush();

        quick_xml::se::to_writer(xml.inner(), &self.sheet_pr)?;
        quick_xml::se::to_writer(xml.inner(), &self.sheet_views)?;
        quick_xml::se::to_writer(xml.inner(), &self.dimension)?;
        quick_xml::se::to_writer(xml.inner(), &self.sheet_format_pr)?;
        quick_xml::se::to_writer(xml.inner(), &self.sheet_data)?;
        quick_xml::se::to_writer(xml.inner(), &self.page_margins)?;
        quick_xml::se::to_writer(xml.inner(), &self.header_footer)?;

        // ends types element.
        let end = BytesEnd::borrowed(Self::tag_name().as_bytes());
        xml.write_event(Event::End(end))?;
        Ok(())
    }
}

#[test]
fn serde() {
    let xml = include_str!("../../../examples/simple-spreadsheet/xl/worksheets/sheet1.xml");
    println!("{}", xml);
    let worksheet = WorksheetPart::from_xml_str(xml).unwrap();
    println!("{:?}", worksheet);
    // let mut buffer = Vec::new();
    // let mut writer = quick_xml::Writer::new(&mut buffer);
    // let mut ser = quick_xml::se::Serializer::with_root(writer, Some("worksheet"));
    // //to_string(&worksheet).unwrap(); //worksheet.to_xml_string().unwrap();
    // worksheet.serialize(&mut ser).unwrap();
    // let str = String::from_utf8_lossy(&buffer);
    // println!("{}", str);
    let str = worksheet.to_xml_string().unwrap();
    println!("{}", str);
    let worksheet2: WorksheetPart = quick_xml::de::from_str(&str).unwrap();
    println!("{:?}", worksheet2);
    assert_eq!(worksheet, worksheet2);
}
