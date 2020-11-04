use crate::packaging::namespace::Namespaces;
use crate::packaging::element::*;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "numFmt")]
pub struct NumberFormat {
    #[serde(rename = "numFmtId")]
    id: usize,
    #[serde(rename = "formatCode")]
    code: String,
}

impl OpenXmlElementInfo for NumberFormat {
    fn tag_name() -> &'static str {
        "numFmt"
    }
    fn element_type() -> OpenXmlElementType {
        OpenXmlElementType::Node
    }
}
impl OpenXmlDeserializeDefault for NumberFormat {}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "numFmts")]
pub struct NumberFormats {
    #[serde(rename = "numFmt")]
    num_fmt: Vec<NumberFormat>,
}
impl OpenXmlDeserializeDefault for NumberFormats {}


#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "numFmt")]
pub struct Font {
    #[serde(rename = "sz")]
    size: String,
    /// the color theme id
    color: String,
    name: String,
    charset: Option<String>,
    scheme: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "fonts")]
pub struct Fonts {
    font: Vec<Font>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "fill")]
#[serde(rename_all = "camelCase")]
pub struct Fill {
    pattern_type: Option<String>,
    bg_color: Option<String>,
    fg_color: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "fills")]
pub struct Fills {
    count: usize,
    #[serde(rename = "fill")]
    fills: Vec<Fill>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "borderStyle")]
pub struct BorderStyle {
    style: Option<String>,
    /// color theme id
    color: Option<usize>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "border")]
pub struct Diagonal {}
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "border")]
pub struct Border {
    diagonal: Diagonal,
    left: Option<BorderStyle>,
    right: Option<BorderStyle>,
    top: Option<BorderStyle>,
    bottom: Option<BorderStyle>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "borders")]
pub struct Borders {
    #[serde(rename = "border")]
    borders: Vec<Border>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Xf {
    num_fmt_id: usize,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    apply_number_format: Option<bool>,
    apply_fill: Option<bool>,
    apply_alignment: Option<bool>,
    apply_protection: Option<bool>,
    alignment: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellStyleXfs {
    count: usize,
    xf: Vec<Xf>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellStyle {
    name: String,
    xf_id: usize,
    builtin_id: usize,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellStyles {
    count: usize,
    cell_style: Vec<CellStyle>,
}
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStyles {
    count: usize,
    default_table_style: String,
    default_pilot_style: String,
    styles: Vec<TableStyle>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStyle {}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ExtLst {
    uri: String,
    namespaces: Namespaces,
    slicer_styles: Vec<SlicerStyle>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", rename = "x14:slicerStyles")]
pub struct SlicerStyle {
    default: String,
}

/// App properties
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "styleSheet", rename_all = "camelCase")]
pub struct StylesPart {
    num_fmts: Option<NumberFormats>,
    fonts: Option<Fonts>,
    fills: Option<Fills>,
    cell_style_xfs: Option<CellStyleXfs>,
    // borders: Borders,
    cell_styles: Option<CellStyles>,
    // ext_lst: ExtLst,
    #[serde(flatten)]
    namespaces: Namespaces,
}

impl StylesPart {
    pub fn default_spreadsheet_styles() -> StylesPart {
        let namespaces =
            Namespaces::new("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        let _num_fmts = NumberFormats::default();
        Self {
            namespaces,
            ..Default::default()
        }
    }

    /// Get cell style by id, 0-based.
    pub fn get_cell_style(&self, id: usize) -> Option<&CellStyle> {
        self.cell_styles
            .as_ref()
            .and_then(|cs| cs.cell_style.get(id))
    }
    /// Get cell style xf by id, 0-based.
    pub fn get_xf(&self, id: usize) -> Option<&Xf> {
        self.cell_style_xfs.as_ref().and_then(|xf| xf.xf.get(id))
    }
    /// Get cell style by id, 0-based.
    pub fn get_number_format(&self, id: usize) -> Option<&NumberFormat> {
        self.num_fmts
            .as_ref()
            .and_then(|inner| inner.num_fmt.get(id))
    }
}

impl OpenXmlElementInfo for StylesPart {
    fn tag_name() -> &'static str {
        "styleSheet"
    }
}

impl OpenXmlDeserializeDefault for StylesPart {}

// impl fmt::Display for SharedStringsPart {
//     fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
//         let mut container = Vec::new();
//         let mut cursor = std::io::Cursor::new(&mut container);
//         self.write(&mut cursor).expect("write xml to memory error");
//         let s = String::from_utf8_lossy(&container);
//         write!(f, "{}", s)?;
//         Ok(())
//     }
// }

#[test]
fn test_de() {
    //const raw: &str = ;
    //println!("{}", raw);
    let value = StylesPart::from_xml_file("examples/simple-spreadsheet/xl/styles.xml").unwrap();
    println!("{:?}", value);
    // let display = format!("{}", value);
    // println!("{}", display);
    // assert_eq!(raw, display);
}
