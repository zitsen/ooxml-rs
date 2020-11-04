# OOXML - Office OpenXML parser in Rust

**The source code is in my private github project https://github.com/zitsen/ooxml-rs . I'm glad to add you as a member of the project when I have your github username.**

This crate is started as a private-purposed project with limited knownledge of Office Open XML. Never use it in production!!!
本项目只作为学习使用，不包含任何开发承诺，不可用于生产环境！！！！

> Office Open XML，为由Microsoft开发的一种以XML为基础并以ZIP格式压缩的电子文件规范，支持文件、表格、备忘录、幻灯片等文件格式。

> Office Open XML (also informally known as OOXML or Microsoft Open XML (MOX)) is a zipped, XML-based file format developed by Microsoft for representing spreadsheets, charts, presentations and word processing documents. The format was initially standardized by Ecma (as ECMA-376), and by the ISO and IEC (as ISO/IEC 29500) in later versions.

OOXML, as it's naming, is trying to be a pure rust implementation of Office Open XML parser - reading and writing ooxml components efficiently in Rust.

## TLDR;

Example code in `examples/xlsx.rs`:

```rust
use ooxml::document::SpreadsheetDocument;

fn main() {
   let xlsx = SpreadsheetDocument::open("examples/simple-spreadsheet/data-image-demo.xlsx").unwrap();
   let workbook = xlsx.get_workbook();
   println!("{:?}", xlsx);
   let sheet_names = workbook.worksheet_names();
   for sheet_name in sheet_names {
      println!("this is sheet {}", sheet_name);
   }

   for sheet in workbook.worksheets() {
      // do stuff.
   }
}
```

Run `cargo run --example xlsx`:

```
SpreadsheetDocument { parts: RefCell { value: SpreadsheetParts { initialized: true, relationships: Relationships { relationships: {"rId5": Relationship { id: "rId5", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", target: "sharedStrings.xml" }, "rId4": Relationship { id: "rId4", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", target: "styles.xml" }, "rId3": Relationship { id: "rId3", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", target: "theme/theme1.xml" }, "rId2": Relationship { id: "rId2", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", target: "worksheets/sheet2.xml" }, "rId1": Relationship { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", target: "worksheets/sheet1.xml" }} }, workbook: WorkbookPart { file_version: FileVersion { app_name: Some("xl"), last_edited: Some(3), lowest_edited: Some(5), run_build: None }, book_views: BookViews { workbook_view: WorkbookView { window_width: Some(23550), window_height: Some(12630), active_tab: Some(1) } }, workbook_pr: WorkbookPr, sheets: Sheets { sheets: [Sheet { name: "image", sheet_id: 2, r_id: "rId1" }, Sheet { name: "new", sheet_id: 3, r_id: "rId2" }] }, calc_pr: CalcPr { calc_id: "144525" }, namespaces: Namespaces({"xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}) }, styles: StylesPart { num_fmts: Some(NumberFormats { num_fmt: [NumberFormat { id: 176, code: "yyyy/m/d;@" }, NumberFormat { id: 42, code: "_ \"￥\"* #,##0_ ;_ \"￥\"* \\-#,##0_ ;_ \"￥\"* \"-\"_ ;_ @_ " }, NumberFormat { id: 43, code: "_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * \"-\"??_ ;_ @_ " }, NumberFormat { id: 44, code: "_ \"￥\"* #,##0.00_ ;_ \"￥\"* \\-#,##0.00_ ;_ \"￥\"* \"-\"??_ ;_ @_ " }, NumberFormat { id: 41, code: "_ * #,##0_ ;_ * \\-#,##0_ ;_ * \"-\"_ ;_ @_ " }] }), fonts: Some(Fonts { font: [Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }, Font { size: "", color: "", name: "", charset: Some(""), scheme: Some("") }] }), fills: Some(Fills { count: 33, fills: [Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }, Fill { pattern_type: None, bg_color: None, fg_color: None }] }), cell_style_xfs: Some(CellStyleXfs { count: 49, xf: [Xf { num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0, apply_number_format: None, apply_fill: None, apply_alignment: None, apply_protection: None, alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 11, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 3, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 22, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 17, fill_id: 19, border_id: 4, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 32, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 26, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 44, font_id: 0, fill_id: 0, border_id: 0, apply_number_format: None, apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 31, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 9, font_id: 0, fill_id: 0, border_id: 0, apply_number_format: None, apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 21, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 20, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 16, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 14, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 13, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 10, fill_id: 6, border_id: 4, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 29, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 16, fill_id: 12, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 27, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 7, fill_id: 5, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 9, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 15, fill_id: 0, border_id: 7, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 18, fill_id: 23, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 13, fill_id: 8, border_id: 5, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 14, fill_id: 6, border_id: 6, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 11, fill_id: 0, border_id: 2, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 12, fill_id: 0, border_id: 0, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 18, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 4, fill_id: 0, border_id: 0, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 42, font_id: 0, fill_id: 0, border_id: 0, apply_number_format: None, apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 7, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 43, font_id: 0, fill_id: 0, border_id: 0, apply_number_format: None, apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 9, fill_id: 0, border_id: 0, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 8, fill_id: 0, border_id: 0, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 30, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 19, fill_id: 0, border_id: 0, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 17, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 0, fill_id: 10, border_id: 8, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 15, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 4, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 25, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 5, fill_id: 0, border_id: 0, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 41, font_id: 0, fill_id: 0, border_id: 0, apply_number_format: None, apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 3, fill_id: 0, border_id: 2, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 24, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 4, fill_id: 0, border_id: 3, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 6, fill_id: 28, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 2, fill_id: 2, border_id: 0, apply_number_format: Some(false), apply_fill: None, apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }, Xf { num_fmt_id: 0, font_id: 1, fill_id: 0, border_id: 1, apply_number_format: Some(false), apply_fill: Some(false), apply_alignment: Some(false), apply_protection: Some(false), alignment: "" }] }), cell_styles: Some(CellStyles { count: 49, cell_style: [CellStyle { name: "常规", xf_id: 0, builtin_id: 0 }, CellStyle { name: "60% - 强调文字颜色 6", xf_id: 1, builtin_id: 52 }, CellStyle { name: "20% - 强调文字颜色 4", xf_id: 2, builtin_id: 42 }, CellStyle { name: "强调文字颜色 4", xf_id: 3, builtin_id: 41 }, CellStyle { name: "输入", xf_id: 4, builtin_id: 20 }, CellStyle { name: "40% - 强调文字颜色 3", xf_id: 5, builtin_id: 39 }, CellStyle { name: "20% - 强调文字颜色 3", xf_id: 6, builtin_id: 38 }, CellStyle { name: "货币", xf_id: 7, builtin_id: 4 }, CellStyle { name: "强调文字颜色 3", xf_id: 8, builtin_id: 37 }, CellStyle { name: "百分比", xf_id: 9, builtin_id: 5 }, CellStyle { name: "60% - 强调文字颜色 2", xf_id: 10, builtin_id: 36 }, CellStyle { name: "60% - 强调文字颜色 5", xf_id: 11, builtin_id: 48 }, CellStyle { name: "强调文字颜色 2", xf_id: 12, builtin_id: 33 }, CellStyle { name: "60% - 强调文字颜色 1", xf_id: 13, builtin_id: 32 }, CellStyle { name: "60% - 强调文字颜色 4", xf_id: 14, builtin_id: 44 }, CellStyle { name: "计算", xf_id: 15, builtin_id: 22 }, CellStyle { name: "强调文字颜色 1", xf_id: 16, builtin_id: 29 }, CellStyle { name: "适中", xf_id: 17, builtin_id: 28 }, CellStyle { name: "20% - 强调文字颜色 5", xf_id: 18, builtin_id: 46 }, CellStyle { name: "好", xf_id: 19, builtin_id: 26 }, CellStyle { name: "20% - 强调文字颜色 1", xf_id: 20, builtin_id: 30 }, CellStyle { name: "汇总", xf_id: 21, builtin_id: 25 }, CellStyle { name: "差", xf_id: 22, builtin_id: 27 }, CellStyle { name: "检查单元格", xf_id: 23, builtin_id: 23 }, CellStyle { name: "输出", xf_id: 24, builtin_id: 21 }, CellStyle { name: "标题 1", xf_id: 25, builtin_id: 16 }, CellStyle { name: "解释性文本", xf_id: 26, builtin_id: 53 }, CellStyle { name: "20% - 强调文字颜色 2", xf_id: 27, builtin_id: 34 }, CellStyle { name: "标题 4", xf_id: 28, builtin_id: 19 }, CellStyle { name: "货币[0]", xf_id: 29, builtin_id: 7 }, CellStyle { name: "40% - 强调文字颜色 4", xf_id: 30, builtin_id: 43 }, CellStyle { name: "千位分隔", xf_id: 31, builtin_id: 3 }, CellStyle { name: "已访问的超链接", xf_id: 32, builtin_id: 9 }, CellStyle { name: "标题", xf_id: 33, builtin_id: 15 }, CellStyle { name: "40% - 强调文字颜色 2", xf_id: 34, builtin_id: 35 }, CellStyle { name: "警告文本", xf_id: 35, builtin_id: 11 }, CellStyle { name: "60% - 强调文字颜色 3", xf_id: 36, builtin_id: 40 }, CellStyle { name: "注释", xf_id: 37, builtin_id: 10 }, CellStyle { name: "20% - 强调文字颜色 6", xf_id: 38, builtin_id: 50 }, CellStyle { name: "强调文字颜色 5", xf_id: 39, builtin_id: 45 }, CellStyle { name: "40% - 强调文字颜色 6", xf_id: 40, builtin_id: 51 }, CellStyle { name: "超链接", xf_id: 41, builtin_id: 8 }, CellStyle { name: "千位分隔[0]", xf_id: 42, builtin_id: 6 }, CellStyle { name: "标题 2", xf_id: 43, builtin_id: 17 }, CellStyle { name: "40% - 强调文字颜色 5", xf_id: 44, builtin_id: 47 }, CellStyle { name: "标题 3", xf_id: 45, builtin_id: 18 }, CellStyle { name: "强调文字颜色 6", xf_id: 46, builtin_id: 49 }, CellStyle { name: "40% - 强调文字颜色 1", xf_id: 47, builtin_id: 31 }, CellStyle { name: "链接单元格", xf_id: 48, builtin_id: 24 }] }), namespaces: Namespaces({"xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}) }, shared_strings: SharedStringsPart { count: 6, unique_count: 6, namespaces: Namespaces({"xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}), strings: [SharedString { t: Value("name") }, SharedString { t: Value("age") }, SharedString { t: Value("birthday") }, SharedString { t: Value("bob") }, SharedString { t: Value("tom") }, SharedString { t: Value("cury") }] }, worksheets: {"worksheets/sheet1.xml": WorksheetPart { namespaces: Namespaces({"xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main", "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006", "xmlns:etc": "http://www.wps.cn/officeDocument/2017/etCustomData"}), sheet_pr: SheetPr, dimension: Some(Dimension { ref: "A1" }), sheet_views: Some(SheetViews { sheet_view: SheetView { tab_selected: None, workbook_view_id: Some(0), selection: Some(Selection { active_cell: "A1", sqref: "A1" }) } }), sheet_format_pr: Some(SheetFormatPr { default_col_width: Some(9.0), default_row_height: Some(13.5), outline_level_row: None, outline_level_col: None }), sheet_data: Some(SheetData { rows: None }), page_margins: PageMargins { left: Some(0.75), right: Some(0.75), top: Some(1.0), bottom: Some(1.0), header: Some(0.5), footer: Some(0.5) }, header_footer: HeaderFooter }, "worksheets/sheet2.xml": WorksheetPart { namespaces: Namespaces({"xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main", "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006", "xmlns:etc": "http://www.wps.cn/officeDocument/2017/etCustomData"}), sheet_pr: SheetPr, dimension: Some(Dimension { ref: "A1:C4" }), sheet_views: Some(SheetViews { sheet_view: SheetView { tab_selected: Some(1), workbook_view_id: Some(0), selection: Some(Selection { active_cell: "C4", sqref: "C4" }) } }), sheet_format_pr: Some(SheetFormatPr { default_col_width: Some(9.0), default_row_height: Some(13.5), outline_level_row: Some(3.0), outline_level_col: Some(2.0) }), sheet_data: Some(SheetData { rows: Some([SheetRow { r: 1, spans: "1:3", cols: [SheetCol { r: "A1", t: Some("s"), s: None, v: "0" }, SheetCol { r: "B1", t: Some("s"), s: None, v: "1" }, SheetCol { r: "C1", t: Some("s"), s: None, v: "2" }] }, SheetRow { r: 2, spans: "1:3", cols: [SheetCol { r: "A2", t: Some("s"), s: None, v: "3" }, SheetCol { r: "B2", t: None, s: None, v: "17" }, SheetCol { r: "C2", t: None, s: Some(1), v: "30662" }] }, SheetRow { r: 3, spans: "1:3", cols: [SheetCol { r: "A3", t: Some("s"), s: None, v: "4" }, SheetCol { r: "B3", t: None, s: None, v: "18" }, SheetCol { r: "C3", t: None, s: Some(1), v: "30297" }] }, SheetRow { r: 4, spans: "1:3", cols: [SheetCol { r: "A4", t: Some("s"), s: None, v: "5" }, SheetCol { r: "B4", t: None, s: None, v: "20" }, SheetCol { r: "C4", t: None, s: Some(1), v: "29567" }] }]) }), page_margins: PageMargins { left: Some(0.75), right: Some(0.75), top: Some(1.0), bottom: Some(1.0), header: Some(0.5), footer: Some(0.5) }, header_footer: HeaderFooter }} } }, document_type: Workbook, workbook: Workbook { worksheets: [Worksheet { name: "image", sheet_id: 2, part: WorksheetPart { namespaces: Namespaces({"xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main", "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006", "xmlns:etc": "http://www.wps.cn/officeDocument/2017/etCustomData"}), sheet_pr: SheetPr, dimension: Some(Dimension { ref: "A1" }), sheet_views: Some(SheetViews { sheet_view: SheetView { tab_selected: None, workbook_view_id: Some(0), selection: Some(Selection { active_cell: "A1", sqref: "A1" }) } }), sheet_format_pr: Some(SheetFormatPr { default_col_width: Some(9.0), default_row_height: Some(13.5), outline_level_row: None, outline_level_col: None }), sheet_data: Some(SheetData { rows: None }), page_margins: PageMargins { left: Some(0.75), right: Some(0.75), top: Some(1.0), bottom: Some(1.0), header: Some(0.5), footer: Some(0.5) }, header_footer: HeaderFooter } }, Worksheet { name: "new", sheet_id: 3, part: WorksheetPart { namespaces: Namespaces({"xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main", "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006", "xmlns:etc": "http://www.wps.cn/officeDocument/2017/etCustomData"}), sheet_pr: SheetPr, dimension: Some(Dimension { ref: "A1:C4" }), sheet_views: Some(SheetViews { sheet_view: SheetView { tab_selected: Some(1), workbook_view_id: Some(0), selection: Some(Selection { active_cell: "C4", sqref: "C4" }) } }), sheet_format_pr: Some(SheetFormatPr { default_col_width: Some(9.0), default_row_height: Some(13.5), outline_level_row: Some(3.0), outline_level_col: Some(2.0) }), sheet_data: Some(SheetData { rows: Some([SheetRow { r: 1, spans: "1:3", cols: [SheetCol { r: "A1", t: Some("s"), s: None, v: "0" }, SheetCol { r: "B1", t: Some("s"), s: None, v: "1" }, SheetCol { r: "C1", t: Some("s"), s: None, v: "2" }] }, SheetRow { r: 2, spans: "1:3", cols: [SheetCol { r: "A2", t: Some("s"), s: None, v: "3" }, SheetCol { r: "B2", t: None, s: None, v: "17" }, SheetCol { r: "C2", t: None, s: Some(1), v: "30662" }] }, SheetRow { r: 3, spans: "1:3", cols: [SheetCol { r: "A3", t: Some("s"), s: None, v: "4" }, SheetCol { r: "B3", t: None, s: None, v: "18" }, SheetCol { r: "C3", t: None, s: Some(1), v: "30297" }] }, SheetRow { r: 4, spans: "1:3", cols: [SheetCol { r: "A4", t: Some("s"), s: None, v: "5" }, SheetCol { r: "B4", t: None, s: None, v: "20" }, SheetCol { r: "C4", t: None, s: Some(1), v: "29567" }] }]) }), page_margins: PageMargins { left: Some(0.75), right: Some(0.75), top: Some(1.0), bottom: Some(1.0), header: Some(0.5), footer: Some(0.5) }, header_footer: HeaderFooter } }] } }
this is sheet image
this is sheet new
```

## Library Design

The main idea come from the [DotNet OpenXML SDK].

1. Implement [OpenXML Package Convention] for any OOXML format(docx/xlsx/pptx...), including:
   - package read and write
   - content type parsing
   - relationship common types
2. Implement shared OpenXML parts
   - content type
   - core properties
   - app properties
   - file properties
   - embedded package
   - image
   - theme
   - style
3. Implement [Excel/SpreadsheetML specifications](http://officeopenxml.com/anatomyofOOXML-xlsx.php)
   - Calculation Chain
   - Chartsheet
   - Comments
   - Connections
   - Custom Property
   - Customer XML Mappings
   - Dialogsheet
   - Drawings
   - External Workbook References
   - Metadata
   - Pivot Table
   - Pivot Table Cache Definition
   - Pivot Table Cache Records
   - Query Table
   - Shared String Table
   - Shared Workbook Revision Log
   - Shared Workbook User Data
   - Single Cell Table Definition
   - Table Definition
   - Volatile Dependencies
   - Workbook
   - Worksheet
4. Other OpenXML formats(docx, pptx)

Codebase tree structure will be like below.

```text
src
├── document
│   ├── mod.rs
│   ├── presentation
│   │   └── mod.rs
│   ├── spreadsheet
│   │   ├── cell.rs
│   │   ├── chart.rs
│   │   ├── document_type.rs
│   │   ├── drawing.rs
│   │   ├── media.rs
│   │   ├── mod.rs
│   │   ├── shared_string.rs
│   │   ├── style.rs
│   │   ├── workbook.rs
│   │   └── worksheet.rs
│   └── wordprocessing
│       └── mod.rs
├── drawing
│   └── mod.rs
├── error.rs
├── lib.rs
├── math
│   └── mod.rs
└── packaging
    ├── app_property.rs
    ├── content_type.rs
    ├── custom_property.rs
    ├── element.rs
    ├── mod.rs
    ├── namespace.rs
    ├── package.rs
    ├── part
    │   ├── container.rs
    │   ├── mod.rs
    │   └── pair.rs
    ├── property.rs
    ├── relationship
    │   ├── mod.rs
    │   └── reference.rs
    ├── variant.rs
    ├── xml.rs
    └── zip.rs
```

## Definitions For the Crate

**The main design principle is `typed everything`.**

- **`Package`**: A `Package` is a zipped OpenXML document, which could be wordprocessing/spreadsheet/presentation document.
- **`Element`**: An `Element` is an OpenXML element reperasenting data details in each xml.
- **`Part`**: A `Part` is a collection of `Element`s or pure data that should be serializing to an file in the package.
- **`Component`**: A `Component` is the bridge of behaviors and the internal OpenXML stuff, including `Package`, `Element`, and `Part`.
- **`Property`**: A `Property` represents attributes for an element.
- **`Document`**: A `Document` is the entry `Component` for an real document, eg. `SpreadSheetDocument` etc.
- **`RelationShip`**: A `RelationShip` is a link relationship for the element and other resources from a `Part`.

The data flows open or create an document will be like below.

```plantuml
Document -> Package : open/parse from
Package -> Parts : parse to parts
Parts -> Components: build components tree
Components -> Elements: elements one-to-one map
Elements -> Components: elements changes
Components -> Parts: components write back
Parts -> Package: serialize to package
Package <- Document: flush, save or others

Document -> Components: create new document. add or remove components
Components <-> Elements: operations
Components -> Parts: component add/remove
Parts -> Package: serialize to package
Document -> Package: flush, save or others
```

## Initialize Implementing Features

- [x] OPC parsing, include read and write
- [x] Shared components
  - [x] content type
  - [x] core properties
  - [x] app properties
  - [ ] file properties(not in schedule)
  - [ ] embedded package(not int schedule)
  - [ ] image
  - [ ] theme
  - [ ] style
- [ ] SpreadsheetML
  - [ ] Workbook
  - [ ] Worksheet

TODOS:
- create marker traits for OpenXML element, make it more generialize.
- use `minidom` in an xml part, tracking the changes and write back to dom tree.
- lazy parse some of the openxml part for first start speedup.
- implement helper macros for component generation.
  
## Tokei - 2020-11-04-11:35:51

```text
===============================================================================
 Language            Files        Lines         Code     Comments       Blanks
===============================================================================
 Markdown                1          272            0          230           42
 Plain Text              1            1            0            1            0
 TOML                    1           23           21            1            1
 XML                    52          164          164            0            0
-------------------------------------------------------------------------------
 Rust                   34         2721         2189          194          338
 |- Markdown            14          106            7           90            9
 (Total)                           2827         2196          284          347
===============================================================================
 Total                  89         3287         2381          516          390
===============================================================================
```

## Concepts

### Office Open XML, or OpenXML

Office Open XML (also informally known as OOXML or Microsoft Open XML (MOX)) is a zipped, XML-based file format developed by Microsoft for representing spreadsheets, charts, presentations and word processing documents. The format was initially standardized by Ecma (as ECMA-376), and by the ISO and IEC (as ISO/IEC 29500) in later versions.

Microsoft Office 2010 provides read support for ECMA-376, read/write support for ISO/IEC 29500 Transitional, and read support for ISO/IEC 29500 Strict. Microsoft Office 2013 and Microsoft Office 2016 additionally support both reading and writing of ISO/IEC 29500 Strict.While Office 2013 and onward have full read/write support for ISO/IEC 29500 Strict, Microsoft has not yet implemented the strict non-transitional, or original standard, as the default file format yet due to remaining interoperability concerns.

### OpenXML Package Convention

The Open Packaging Conventions (OPC) is a container-file technology initially created by Microsoft to store a combination of XML and non-XML files that together form a single entity such as an Open XML Paper Specification (OpenXPS) document. OPC-based file formats combine the advantages of leaving the independent file entities embedded in the document intact and resulting in much smaller files compared to normal use of XML.

### Standard ECMA-376

[Standard ECMA-376] - The Office Open XML File Formats standard.

1st edition (December 2006), 2nd edition (December 2008), 3rd edition (June 2011), 4th edition (December 2012) and 5th edition (Part 3, December 2015; and Parts 1 & 4, December 2016).

Edition downloads:

- [ECMA-376 5th edition Part 1]
- [ECMA-376 5th edition Part 3]
- [ECMA-376 5th edition Part 4]
  
- [ECMA-376 4th edition Part 1]
- [ECMA-376 4th edition Part 2]
- [ECMA-376 4th edition Part 3]
- [ECMA-376 4th edition Part 4]
  
Currently is 4th edition, technically aligned with ISO/IEC 29500. 5th edition is ongoing. There is a [Office Open XML Overview] introduction pdf file.

### SpreadsheetML

A SpreadsheetML or .xlsx file is a zip file (a package) containing a number of "parts" (typically UTF-8 or UTF-16 encoded) or XML files. The package may also contain other media files such as images. The structure is organized according to the Open Packaging Conventions as outlined in Part 2 of the OOXML standard ECMA-376.

You can look at the file structure and the files that comprise a SpreadsheetML file by simply unzipping the .xlsx file.

```text
├── [Content_Types].xml
├── docProps
│   ├── app.xml
│   ├── core.xml
│   └── custom.xml
├── _rels
└── xl
    ├── charts
    │   ├── chart1.xml
    │   ├── colors1.xml
    │   ├── _rels
    │   │   └── chart1.xml.rels
    │   └── style1.xml
    ├── drawings
    │   ├── drawing1.xml
    │   ├── drawing2.xml
    │   └── _rels
    │       ├── drawing1.xml.rels
    │       └── drawing2.xml.rels
    ├── media
    │   └── image1.png
    ├── _rels
    │   └── workbook.xml.rels
    ├── sharedStrings.xml
    ├── styles.xml
    ├── theme
    │   └── theme1.xml
    ├── workbook.xml
    └── worksheets
        ├── _rels
        │   ├── sheet1.xml.rels
        │   └── sheet2.xml.rels
        ├── sheet1.xml
        └── sheet2.xml
```

The number and types of parts will vary based on what is in the spreadsheet, but there will always be a `[Content_Types].xml`, one or more relationship parts, a workbook part , and at least one worksheet. The core data of the spreadsheet is contained within the worksheet part(s), discussed in more detail at [xslx Content Overview](http://officeopenxml.com/SScontentOverview.php).

## Resources

1. Wikipedia Office OpenXML: [English](https://en.wikipedia.org/wiki/Office_Open_XML), [中文](https://zh.wikipedia.org/wiki/Office_Open_XML).
2. Microsoft [DotNet OpenXML SDK] documents and [source code](https://github.com/OfficeDev/Open-XML-SDK/).
3. Wikipedia [OpenXML Package Convention] - [开放打包约定].
4. What is OOXML: http://officeopenxml.com/
5. SpreadsheetML: http://officeopenxml.com/anatomyofOOXML-xlsx.php
6. Rust [quick-xml](https://crates.io/crates/quick-xml) [documents](https://docs.rs/quick-xml/0.20.0).
7. Rust [docx-rs](https://crates.io/crates/docx-rs) [documents](https://docs.rs/docx-rs) and [source code on github](https://github.com/bokuweb/docx-rs).
8. Go Excel file parser [excelize](https://github.com/360EntSecGroup-Skylar/excelize).
9. [Standard ECMA-376].

[Office Open XML]: http://officeopenxml.com/
[DotNet OpenXML SDK]: https://docs.microsoft.com/en-us/dotnet/api/overview/openxml/?view=openxml-2.8.1
[OpenXML Package Convention]: https://en.wikipedia.org/wiki/Open_Packaging_Conventions
[开放打包约定]: https://zh.wikipedia.org/wiki/%E5%BC%80%E6%94%BE%E6%89%93%E5%8C%85%E7%BA%A6%E5%AE%9A
[Standard ECMA-376]: https://www.ecma-international.org/publications/standards/Ecma-376.htm
[ECMA-376 5th edition Part 1]: https://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fifth%20Edition,%20Part%201%20-%20Fundamentals%20And%20Markup%20Language%20Reference.zip
[ECMA-376 5th edition Part 3]: https://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fifth%20Edition,%20Part%203%20-%20Markup%20Compatibility%20and%20Extensibility.zip
[ECMA-376 5th edition Part 4]: https://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fifth%20Edition,%20Part%204%20-%20Transitional%20Migration%20Features.zip

[ECMA-376 4th edition Part 1]: https://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fourth%20Edition,%20Part%201%20-%20Fundamentals%20And%20Markup%20Language%20Reference.zip
[ECMA-376 4th edition Part 2]: https://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fourth%20Edition,%20Part%202%20-%20Open%20Packaging%20Conventions.zip
[ECMA-376 4th edition Part 3]: https://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fourth%20Edition,%20Part%203%20-%20Markup%20Compatibility%20and%20Extensibility.zip
[ECMA-376 4th edition Part 4]: https://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fourth%20Edition,%20Part%204%20-%20Transitional%20Migration%20Features.zip
[Office Open XML Overview]: https://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf