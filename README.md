# OOXML - Office Open XML reader and writer

**This crate is started as a private-purposed project with limited knownledge of Office Open XML. Never use it in production!!!**
**本项目只作为学习使用，不包含任何开发承诺，不可用于生产环境！！！！**

> Office Open XML，为由Microsoft开发的一种以XML为基础并以ZIP格式压缩的电子文件规范，支持文件、表格、备忘录、幻灯片等文件格式。

> Office Open XML (also informally known as OOXML or Microsoft Open XML (MOX)) is a zipped, XML-based file format developed by Microsoft for representing spreadsheets, charts, presentations and word processing documents. The format was initially standardized by Ecma (as ECMA-376), and by the ISO and IEC (as ISO/IEC 29500) in later versions.

OOXML, as it's naming, is trying to be a pure rust implementation of Office Open XML parser - reading and writing ooxml components efficiently in Rust.

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
   
## Initialize Implementing Features

First, implement a excel parser with these features:

- [x] Open package detection/validation
- [x] Package parsing with Open Package Convention
- [ ] Package relationship
- [ ] SpreadsheetML
  - [ ] Workbook
  - [ ] Worksheet

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

The number and types of parts will vary based on what is in the spreadsheet, but there will always be a `[Content_Types].xml`, one or more relationship parts, a workbook part , and at least one worksheet. The core data of the spreadsheet is contained within the worksheet part(s), discussed in more detail at [xsxl Content Overview](http://officeopenxml.com/SScontentOverview.php).

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