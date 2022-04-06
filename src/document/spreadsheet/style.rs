use crate::packaging::element::*;
use crate::packaging::namespace::Namespaces;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "numFmt")]
pub struct NumberFormat {
    #[serde(rename = "numFmtId")]
    pub id: usize,
    #[serde(rename = "formatCode")]
    pub code: String,
}

impl NumberFormat {
    pub fn new(id: usize, code: impl Into<String>) -> Self {
        Self {
            id,
            code: code.into(),
        }
    }
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
    num_fmt: Option<Vec<NumberFormat>>,
}
impl OpenXmlDeserializeDefault for NumberFormats {}

pub(crate) mod font {
    use serde::{Deserialize, Serialize};

    #[derive(Debug, Clone, Default, Serialize, Deserialize)]
    #[serde(rename = "sz")]
    pub struct FontSize {
        val: f64,
    }
    #[derive(Debug, Clone, Default, Serialize, Deserialize)]
    #[serde(rename = "name")]
    pub struct FontName {
        val: String,
    }
    #[derive(Debug, Clone, Default, Serialize, Deserialize)]
    #[serde(rename = "charset")]
    pub struct FontCharset {
        val: String,
    }
    #[derive(Debug, Clone, Default, Serialize, Deserialize)]
    #[serde(rename = "scheme")]
    pub struct FontScheme {
        val: String,
    }
    #[derive(Debug, Clone, Default, Serialize, Deserialize)]
    #[serde(rename = "color")]
    pub struct FontColor {
        theme: Option<usize>,
        rbg: Option<String>,
    }
    #[derive(Debug, Clone, Default, Serialize, Deserialize)]
    #[serde(rename = "b")]
    pub struct FontBlack;

    #[derive(Debug, Clone, Default, Serialize, Deserialize)]
    #[serde(rename = "numFmt")]
    pub struct Font {
        black: Option<FontBlack>,
        #[serde(rename = "sz")]
        size: Option<FontSize>,
        /// the color theme id
        color: Option<FontColor>,
        name: String,
        charset: Option<String>,
        scheme: Option<String>,
    }

    #[derive(Debug, Clone, Default, Serialize, Deserialize)]
    #[serde(rename = "fonts")]
    pub struct Fonts {
        #[serde(rename = "font")]
        pub(crate) fonts: Vec<Font>,
    }
}
pub use font::*;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "fgColor")]
#[serde(rename_all = "camelCase")]
pub struct PatternFill {
    pattern_type: Option<String>,
    bg_color: Option<BgColor>,
    fg_color: Option<FgColor>,
}
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "fgColor")]
#[serde(rename_all = "camelCase")]
pub struct FgColor {
    theme: Option<usize>,
    tint: Option<f64>,
    indexed: Option<usize>,
}
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "bgColor")]
#[serde(rename_all = "camelCase")]
pub struct BgColor {
    theme: Option<usize>,
    tint: Option<f64>,
    indexed: Option<usize>,
}
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "fill")]
#[serde(rename_all = "camelCase")]
pub struct Fill {
    pattern_fill: Option<PatternFill>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "fills")]
pub struct Fills {
    count: usize,
    #[serde(rename = "fill")]
    pub(crate) fills: Vec<Fill>,
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
pub struct Alignment {
    vertical: Option<String>,
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
    alignment: Option<Alignment>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellStyleXfs {
    count: usize,
    xf: Vec<Xf>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellXfs {
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
pub struct CellStylesPart {
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
    cell_xfs: Option<CellXfs>,
    // borders: Borders,
    cell_styles: Option<CellStylesPart>,
    // ext_lst: ExtLst,
    #[serde(flatten)]
    namespaces: Namespaces,
}

#[derive(Debug)]
pub struct CellFormatComponent<'a> {
    styles: &'a StylesPart,
    xf: &'a Xf,
}

impl<'a> CellFormatComponent<'a> {
    pub fn number_format(&self) -> Option<&NumberFormat> {
        self.styles.get_number_format(self.xf.num_fmt_id)
    }
    pub fn font(&self) -> Option<&Font> {
        self.styles.get_font(self.xf.font_id)
    }
    pub fn fill(&self) -> Option<&Fill> {
        self.styles.get_fill(self.xf.fill_id)
    }
    pub fn apply_number_format(&self) -> bool {
        self.xf.apply_alignment.unwrap_or_default()
    }
}

#[derive(Debug)]
pub struct CellStyleComponent<'a> {
    styles: &'a StylesPart,
    cell_style: &'a CellStyle,
}

impl<'a> CellStyleComponent<'a> {
    pub fn number_format(&self) -> Option<&NumberFormat> {
        self.xf()
            .and_then(|xf| self.styles.get_number_format(xf.num_fmt_id))
    }
    pub fn xf(&self) -> Option<&Xf> {
        self.styles.get_cell_style_xf(self.cell_style.xf_id)
    }
    pub fn font(&self) -> Option<&Font> {
        self.font_id()
            .and_then(|font_id| self.styles.get_font(font_id))
    }
    pub fn fill(&self) -> Option<&Fill> {
        self.xf()
            .map(|xf| xf.fill_id)
            .and_then(|id| self.styles.get_fill(id))
    }
    pub fn apply_number_format(&self) -> bool {
        self.xf()
            .and_then(|xf| xf.apply_alignment)
            .unwrap_or_default()
    }
    pub fn font_id(&self) -> Option<usize> {
        self.xf().map(|xf| xf.font_id)
    }
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
    pub fn get_cell_format_component<'a>(&'a self, id: usize) -> Option<CellFormatComponent<'a>> {
        let xf = self.get_cell_xf(id);
        xf.map(|xf| CellFormatComponent { styles: self, xf })
    }

    pub fn get_cell_style_component<'a>(&'a self, id: usize) -> Option<CellStyleComponent<'a>> {
        let cell_style = self.get_cell_style(id);
        cell_style.map(|cell_style| CellStyleComponent {
            styles: self,
            cell_style,
        })
    }

    // pub fn get_cell_xf(&self, id: usize) -> Option<&Xf> {
    //     self.cell_xfs
    //         .as_ref()
    //         .and_then(|cs| cs.xf.get(id))
    // }

    /// Get cell style by id, 0-based.
    pub fn get_cell_style(&self, id: usize) -> Option<&CellStyle> {
        self.cell_styles
            .as_ref()
            .and_then(|cs| cs.cell_style.get(id))
    }
    pub fn get_cell_style_xf(&self, id: usize) -> Option<&Xf> {
        self.cell_style_xfs.as_ref().and_then(|xf| {
            let xf1 = dbg!(xf.xf.get(id));
            //let xf2 = xf.xf.iter().find()
            xf1
        })
    }
    /// Get cell style xf by id, 0-based.
    pub fn get_cell_xf(&self, id: usize) -> Option<&Xf> {
        self.cell_xfs.as_ref().and_then(|xf| {
            let xf1 = xf.xf.get(id);
            //let xf2 = xf.xf.iter().find()
            xf1
        })
    }
    /// Get cell style by id, 0-based.
    pub fn get_number_format(&self, id: usize) -> Option<&NumberFormat> {
        use static_init::dynamic;

        //
        // 1 0
        // 2 0.00
        // 3 #,##0
        // 4 #,##0.00
        // 5 $#,##0_);($#,##0)
        // 6 $#,##0_);[Red]($#,##0)
        // 7 $#,##0.00_);($#,##0.00)
        // 8 $#,##0.00_);[Red]($#,##0.00)
        // 9 0%
        // 10 0.00%
        // 11 0.00E+00
        // 12 # ?/?
        // 13 # ??/??
        // 14 m/d/yyyy
        // 15 d-mmm-yy
        // 16 d-mmm
        // 17 mmm-yy
        // 18 h:mm AM/PM
        // 19 h:mm:ss AM/PM
        // 20 h:mm
        // 21 h:mm:ss
        // 22 m/d/yyyy h:mm
        // 37 #,##0_);(#,##0)
        // 38 #,##0_);[Red](#,##0)
        // 39 #,##0.00_);(#,##0.00)
        // 40 #,##0.00_);[Red](#,##0.00)
        // 45 mm:ss
        // 46 [h]:mm:ss
        // 47 mm:ss.0
        // 48 ##0.0E+0
        // 49 @
        macro_rules! _num_fmt {
            ($id:expr, $code:expr) => {
                NumberFormat {
                    id: $id,
                    code: $code.to_string(),
                }
            };
        }
        #[dynamic]
        static BUILTIN_NUMBER_FORMATS: Vec<NumberFormat> = vec![
            NumberFormat::new(0, "General"),
            NumberFormat::new(1, "0"),
            NumberFormat::new(2, "0.00"),
            NumberFormat::new(3, "#,##0"),
            NumberFormat::new(4, "#,##0.00"),
            NumberFormat::new(5, "$#,##0_);($#,##0)"),
            NumberFormat::new(6, "$#,##0_);[Red]($#,##0)"),
            NumberFormat::new(7, "$#,##0.00_);($#,##0.00)"),
            NumberFormat::new(8, "$#,##0.00_);[Red]($#,##0.00)"),
            NumberFormat::new(9, "0%"),
            NumberFormat::new(10, "0.00%"),
            NumberFormat::new(11, "0.00E+00"),
            NumberFormat::new(12, "# ?/?"),
            NumberFormat::new(13, "# ??/??"),
            NumberFormat::new(14, "yyyy/m/d"),
            NumberFormat::new(15, "d-mmm-yy"),
            NumberFormat::new(16, "d-mmm"),
            NumberFormat::new(17, "mmm-yy"),
            NumberFormat::new(18, "h:mm AM/PM"),
            NumberFormat::new(19, "h:mm:ss AM/PM"),
            NumberFormat::new(20, "h:mm"),
            NumberFormat::new(21, "h:mm:ss"),
            NumberFormat::new(22, "m/d/yyyy h:mm"),
            NumberFormat::new(37, "#,##0_);(#,##0)"),
            NumberFormat::new(38, "#,##0_);[Red](#,##0)"),
            NumberFormat::new(39, "#,##0.00_);(#,##0.00)"),
            NumberFormat::new(40, "#,##0.00_);[Red](#,##0.00)"),
            NumberFormat::new(45, "mm:ss"),
            NumberFormat::new(46, "[h]:mm:ss"),
            NumberFormat::new(47, "mm:ss.0"),
            NumberFormat::new(48, "##0.0E+0"),
            NumberFormat::new(49, "@"),
        ];

        if id < 50 {
            return BUILTIN_NUMBER_FORMATS.get(id);
        }
        self.num_fmts
            .as_ref()
            //.and_then(|inner| inner.num_fmt.get(id))
            .and_then(|inner| inner.num_fmt.as_ref())
            .and_then(|inner| inner.iter().find(|nf| nf.id == id))
    }

    // pub fn get_number_format_xf(&self, id: usize) ->
    pub fn get_font(&self, id: usize) -> Option<&Font> {
        self.fonts.as_ref().and_then(|fonts| fonts.fonts.get(id))
    }
    pub fn get_fill(&self, id: usize) -> Option<&Fill> {
        self.fills.as_ref().and_then(|fills| fills.fills.get(id))
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
