use crate::packaging::content_type::ContentType;
use crate::error::OoxmlError;

use serde::{Deserialize, Serialize};

use std::{fmt, fs::File, path::Path};
use std::io::prelude::*;

pub type DateTime = String;
pub const CORE_PROPERTIES_URI: &str = "docProps/core.xml";
pub const CORE_PROPERTIES_NAMESPACE: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
pub const DC_NAMESPACE: &str = "http://purl.org/dc/elements/1.1/";
pub const DCTERMS_NAMESPACE: &str = "http://purl.org/dc/terms/";
pub const DCMITYPE_NAMESPACE: &str = "http://purl.org/dc/dcmitype/";
pub const XSI_NAMESPACE: &str = "http://www.w3.org/2001/XMLSchema-instance";

pub const CORE_PROPERTIES_TAG: &str = "cp:coreProperties";
pub const CORE_PROPERTIES_NAMESPACE_ATTRIBUTE: &str = "xmlns:dc";
pub const DC_NAMESPACE_ATTRIBUTE: &str = "xmlns:dc";
pub const DCTERMS_NAMESPACE_ATTRIBUTE: &str = "xmlns:dcterms";
pub const DCMITYPE_NAMESPACE_ATTRIBUTE: &str = "xmlns:dcmitype";
pub const XSI_NAMESPACE_ATTRIBUTE: &str = "xmlns:xsi";

pub const PROPERTY_CATEGORY_TAG: &str = "dc:creator";
pub const PROPERTY_CONTENT_STATUS_TAG: &str = "dc:contentStatus";
pub const PROPERTY_CONTENT_TYPE_TAG: &str = "dc:contentType";
pub const PROPERTY_CREATED_TAG: &str = "dcterms:created";
pub const PROPERTY_CREATOR_TAG: &str = "dc:creator";
pub const PROPERTY_DESCRIPTION_TAG: &str = "dc:description";
pub const PROPERTY_IDENTIFIER_TAG: &str = "dc:identifier";
pub const PROPERTY_KEYWORDS_TAG: &str = "dc:keywords";
pub const PROPERTY_LANGUAGE_TAG: &str = "dc:language";
pub const PROPERTY_MODIFIED_TAG: &str = "dcterms:modified";
pub const PROPERTY_LAST_MODIFIED_BY_TAG: &str = "cp:lastModifiedBy";
pub const PROPERTY_LAST_PRINTED_TAG: &str = "cp:lastPrinted";
pub const PROPERTY_REVISION_TAG: &str = "cp:revision";
pub const PROPERTY_SUBJECT_TAG: &str = "cp:subject";
pub const PROPERTY_TITLE_TAG: &str = "cp:title";
pub const PROPERTY_VERSION_TAG: &str = "cp:version";
/// Package properties, all the terms came from OpenXML SDK.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all(deserialize = "camelCase"))]
pub struct Properties {
    category: Option<String>,
    content_status: Option<String>,
    content_type: Option<ContentType>,
    created: Option<DateTime>,
    creator: Option<String>,
    description: Option<String>,
    identifier: Option<String>,
    keywords: Option<String>,
    language: Option<String>,
    modified: Option<String>,
    last_modified_by: Option<String>,
    last_printed: Option<DateTime>,
    revision: Option<String>,
    subject: Option<String>,
    title: Option<String>,
    version: Option<String>,
}

impl fmt::Display for Properties {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut container = Vec::new();
        let mut cursor = std::io::Cursor::new(&mut container);
        self.write(&mut cursor).expect("write xml to memory error");
        let s = String::from_utf8_lossy(&container);
        write!(f, "{}", s)?;
        Ok(())
    }
}

impl Properties {
        /// Parse content types data from an xml reader.
        pub fn parse_from_xml_reader<R: BufRead>(reader: R) -> Self {
            quick_xml::de::from_reader(reader).unwrap()
        }
    
        /// Parse content types data from an xml str.
        pub fn parse_from_xml_str(reader: &str) -> Self {
            quick_xml::de::from_str(reader).unwrap()
        }
    
        /// Save to file path.
        pub fn save_as<P: AsRef<Path>>(&self, path: P) -> Result<(), OoxmlError> {
            let file = File::create(path)?;
            self.write(file)
        }
    
        /// Write to an writer
        pub fn write<W: std::io::Write>(&self, writer: W) -> Result<(), OoxmlError> {
            let mut xml = quick_xml::Writer::new( writer);
            use quick_xml::events::attributes::Attribute;
            use quick_xml::events::*;
    
            // 1. write decl
            xml.write_event(Event::Decl(BytesDecl::new(
                b"1.0",
                Some(b"UTF-8"),
                Some(b"yes"),
            )))?;
    
            // 2. start types element
            let mut elem = BytesStart::borrowed_name(CORE_PROPERTIES_TAG.as_bytes());
            
            macro_rules! nsattr {
                ($ns:ident) => {
                    paste::paste! {
                        Attribute {
                            key: [<$ns _NAMESPACE_ATTRIBUTE>].as_bytes(),
                            value: [<$ns _NAMESPACE>].as_bytes().into(),
                        }
                    }
                }
            }

            elem.extend_attributes(vec![
                nsattr!(CORE_PROPERTIES),
                nsattr!(DC),
                nsattr!(DCTERMS),
                nsattr!(DCMITYPE),
                nsattr!(XSI),
            ]);
            xml.write_event(Event::Start(elem))?;

            macro_rules! field {
                ($field:ident) => {
                    paste::paste! {
                        if let Some(field) = &self.$field {
                            let start = BytesStart::borrowed_name([<PROPERTY_ $field:upper _TAG>].as_bytes());
                            let text = BytesText::from_plain_str(&field);
                            let end = BytesEnd::borrowed([<PROPERTY_ $field:upper _TAG>].as_bytes());
                            xml.write_event(Event::Start(start))?;
                            xml.write_event(Event::Text(text))?;
                            xml.write_event(Event::End(end))?;

                        }
                    }
                };
                ($field:ident, $xsi:expr) => {
                    paste::paste! {
                        if let Some(field) = &self.$field {
                            let mut start = BytesStart::borrowed_name([<PROPERTY_ $field:upper _TAG>].as_bytes());
                            start.extend_attributes(vec![
                                Attribute {
                                    key: b"xsi:type",
                                    value: "dcterms:W3CDTF".as_bytes().into()
                                }
                            ]);
                            let text = BytesText::from_plain_str(&field);
                            let end = BytesEnd::borrowed([<PROPERTY_ $field:upper _TAG>].as_bytes());
                            xml.write_event(Event::Start(start))?;
                            xml.write_event(Event::Text(text))?;
                            xml.write_event(Event::End(end))?;

                        }
                    }
                }
            }
            // FIXME(@zitsen): add more field to xml.
            field!(created, "term");
            field!(creator);
            field!(last_modified_by);
            field!(modified, "term");
            field!(revision);

            // bellow may not right
            field!(category);
            field!(content_status);
            field!(content_type);
            field!(description);
            field!(identifier);
            field!(keywords);
            field!(language);
            field!(last_printed);
            field!(subject);
            field!(title);
            field!(version);
    
            // ends types element.
            let end = BytesEnd::borrowed(CORE_PROPERTIES_TAG.as_bytes());
            xml.write_event(Event::End(end))?;
            Ok(())
        }
}
#[test]
fn test_de() {
    let raw = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <dcterms:created xsi:type="dcterms:W3CDTF">1970-01-01T00:00:00Z</dcterms:created>
      <dc:creator>unknown</dc:creator>
      <cp:lastModifiedBy>unknown</cp:lastModifiedBy>
      <dcterms:modified xsi:type="dcterms:W3CDTF">1970-01-01T00:00:00Z</dcterms:modified>
      <cp:revision>1</cp:revision>
    </cp:coreProperties>"#;
    println!("{}", raw);
    let value: Properties = quick_xml::de::from_str(raw).unwrap();
    println!("{:?}", value);
    let display = format!("{}", value);
    println!("{}", display);
    // assert_eq!(raw, display);
}
