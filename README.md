# OOXML - Office OpenXML parser in Rust

**This crate is started as a private-purposed project with limited knownledge of Office Open XML, use it with caution!**

> Office Open XML，为由Microsoft开发的一种以XML为基础并以ZIP格式压缩的电子文件规范，支持文件、表格、备忘录、幻灯片等文件格式。

> Office Open XML (also informally known as OOXML or Microsoft Open XML (MOX)) is a zipped, XML-based file format developed by Microsoft for representing spreadsheets, charts, presentations and word processing documents. The format was initially standardized by Ecma (as ECMA-376), and by the ISO and IEC (as ISO/IEC 29500) in later versions.

OOXML, as it's naming, is trying to be a pure rust implementation of Office Open XML parser - reading and writing ooxml components efficiently in Rust. But at now, only xlsx parsing is supported.

## TLDR;

Example code in `examples/xlsx.rs`:

```rust
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
            // get cell values
            let cols: Vec<_> = rows
                .into_iter()
                .map(|cell| cell.value().unwrap_or_default())
                .collect();
            println!("{}", itertools::join(&cols, ","));
        }
    }
}

```

Run `cargo run --example xlsx`:

```
worksheet 0
worksheet dimension: Some((1, 1))
---------DATA---------

----------------------
worksheet 1
worksheet dimension: Some((4, 4))
---------DATA---------
name,age,birthday,last edited
bob,17,1983/12/12,2020/10/11 19:59
tom,18,1982/12/12,2020/10/11 19:59
cury,20,1980-12-12,2020-10-11 19:59
----------------------
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