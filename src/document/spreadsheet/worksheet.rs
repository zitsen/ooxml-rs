use super::cell::{CellType, CellValue};
use crate::packaging::element::*;
use crate::packaging::namespace::Namespaces;

use quick_xml::events::attributes::Attribute;
use serde::{Deserialize, Serialize};
use std::borrow::Cow;

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheetPr")]
pub struct SheetPr {}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheetView")]
pub struct SheetView {
    tab_selected: Option<String>,
    workbook_view_id: Option<usize>,
    selection: Option<Selection>,
    show_formulas: Option<bool>,
    show_grid_lines: Option<bool>,
    show_row_col_headers: Option<bool>,
    show_zeros: Option<bool>,
    right_to_left: Option<bool>,
    show_outline_symbols: Option<bool>,
    default_grid_color: Option<String>,
    view: Option<String>,
    top_left_cell: Option<String>,
    color_id: Option<usize>,
    zoom_scale: Option<String>,
    zoom_scale_normal: Option<String>,
    zoom_scale_page_layout_view: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheetViews")]
pub struct SheetViews {
    sheet_view: SheetView,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "calcPr")]
pub struct CalcPr {
    calc_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "dimension")]
pub struct Dimension {
    r#ref: String,
}

impl Dimension {
    pub fn dimension(&self) -> (usize, usize) {
        let range = self.r#ref.as_str();
        let range: Vec<&str> = range.split_terminator(':').collect();

        if range.len() < 2 {
            return (1, 1);
        }
        let start = range[0];
        let end = range[1];
        //let (start, end) = range.split_once(':').expect("split at :");
        fn rangify(range: &str) -> (usize, usize) {
            let re: regex::Regex = regex::Regex::new(r"(?P<col>[A-Z]+)(?P<row>\d+)").unwrap();
            let cap = re.captures(range).unwrap();
            let col = cap.name("col").unwrap().as_str();
            let row = cap.name("row").unwrap().as_str().parse().unwrap_or_default();
            fn col_to_idx(col: &str) -> usize {
                if col.is_empty() {
                    return 0;
                }
                let mut idx = 0;
                for (i, c) in col.chars().rev().enumerate() {
                    let c = c.to_digit(36).unwrap();
                    idx += c * 26u32.pow(i as _);

                }
                return idx as usize;
            }
            (row, col_to_idx(col))
        }
        let start = rangify(start);
        let end = rangify(end);
        (end.0 - start.0 + 1, end.1 - start.1 + 1)
    }
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
    pub r: String,
    pub t: Option<String>,
    pub s: Option<usize>,
    #[serde(rename = "$value")]
    pub v: String,
}

impl SheetCol {
    pub fn raw_value(&self) -> CellValue {
        CellValue::String(self.v.to_string())
    }
    pub fn cell_type(&self) -> CellType {
        match (self.t.as_ref(), self.s.as_ref()) {
            (None, None) => CellType::Raw,
            (Some(t), None) => match t {
                s if s == "s" => CellType::Shared(self.v.parse().expect("sharedString id not valid")),
                n if n == "n" => CellType::Number,
                t => unimplemented!("cell type not supported: {}" , t),
            }
            (None, Some(s)) => CellType::Styled(*s),
            (Some(t), Some(s)) => match t {
                t if t == "s" => CellType::Shared(self.v.parse().expect("sharedString id not valid")),
                t if t == "n" => CellType::StyledNumber(*s),
                t => unimplemented!("cell type not supported: {}" , t),
            }
        }
        // if let Some(t) = self.t.as_ref() {
        //     match t {
        //         s if s == "s" => CellType::Shared(self.v.parse().expect("sharedString id not valid")),
        //         n if n == "n" => CellType::Number,
        //         t => unimplemented!("cell type not supported: {}" , t),
        //     }
        // } else if let Some(s) = self.s {
        //     CellType::Styled(s)
        // } else {
        //     CellType::Raw
        // }
    }
}
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename = "row")]
pub struct SheetRow {
    pub r: usize,
    pub spans: Option<String>,
    #[serde(rename = "c")]
    pub cols: Vec<SheetCol>,
}
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "sheetData")]
pub struct SheetData {
    #[serde(rename = "row")]
    pub rows: Option<Vec<SheetRow>>,
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
    pub sheet_pr: SheetPr,
    pub dimension: Option<Dimension>,
    pub sheet_views: Option<SheetViews>,
    pub sheet_format_pr: Option<SheetFormatPr>,
    pub sheet_data: Option<SheetData>,
    pub page_margins: PageMargins,
    pub header_footer: HeaderFooter,
}

impl WorksheetPart {
    pub fn dimension(&self) -> Option<(usize, usize)> {
        self.dimension.as_ref().map(|dim| dim.dimension())
    }
}

impl OpenXmlElementInfo for WorksheetPart {
    fn tag_name() -> &'static str {
        "worksheet"
    }
}

impl OpenXmlDeserializeDefault for WorksheetPart {}

impl OpenXmlSerialize for WorksheetPart {
    fn namespaces(&self) -> Option<Cow<Namespaces>> {
        Some(Cow::Borrowed(&self.namespaces))
    }
    fn attributes(&self) -> Option<Vec<Attribute>> {
        None
    }
    fn write_inner<W: std::io::Write>(&self, writer: W) -> Result<(), crate::error::OoxmlError> {
        let mut xml = quick_xml::Writer::new(writer);
        use quick_xml::events::*;

        quick_xml::se::to_writer(xml.inner(), &self.sheet_pr)?;
        quick_xml::se::to_writer(xml.inner(), &self.sheet_views)?;
        quick_xml::se::to_writer(xml.inner(), &self.dimension)?;
        quick_xml::se::to_writer(xml.inner(), &self.sheet_format_pr)?;
        quick_xml::se::to_writer(xml.inner(), &self.sheet_data)?;
        quick_xml::se::to_writer(xml.inner(), &self.page_margins)?;
        quick_xml::se::to_writer(xml.inner(), &self.header_footer)?;

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
