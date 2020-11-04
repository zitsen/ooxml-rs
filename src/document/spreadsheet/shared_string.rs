use crate::error::OoxmlError;

use serde::{Deserialize, Serialize};

use std::io::prelude::*;
use std::{fmt, fs::File, path::Path};

// pub const SHARED_STRINGS_URI: &str = "xl/sharedStrings.xml";

pub const SHARED_STRINGS_TAG: &str = "sst";
// pub const SHARED_STRINGS_NAMESPACE_ATTRIBUTE: &str = "xmlns";
// pub const SHARED_STRINGS_NAMESPACE: &str =
// "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

pub const SHARED_STRING_TAG: &str = "si";

use crate::packaging::namespace::Namespaces;
use crate::packaging::xml::*;

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all(deserialize = "camelCase"), rename = "t")]
pub struct Value(String);
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all(deserialize = "camelCase"), rename = "si")]
pub struct SharedString {
    t: Value,
}

// #[test]
// fn custom_property_de() {
//     const xml: &str = r#"<t>name</t>"#;
//     let p: SharedString = quick_xml::de::from_str(xml).unwrap();
//     println!("{:?}", p);
//     let s = quick_xml::se::to_string(&p).unwrap();
//     println!("{:?}", s);
//     assert_eq!(xml, s);
// }

// #[test]
// fn custom_properties_de() {
//     const xml: &str = r#"<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="KSOProductBuildVer"><vt:lpwstr>2052-11.1.0.9662</vt:lpwstr></property>"#;
//     let p: Vec<SharedString> = quick_xml::de::from_str(xml).unwrap();
//     println!("{:?}", p);
//     let s = quick_xml::se::to_string(&p).unwrap();
//     println!("{:?}", s);
//     assert_eq!(xml, s);
// }
/// Custom properties
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all(deserialize = "camelCase"), rename = "sst")]
pub struct SharedStringsPart {
    count: usize,
    unique_count: usize,
    #[serde(flatten)]
    namespaces: Namespaces,
    #[serde(rename = "si")]
    strings: Vec<SharedString>,
}

impl OpenXmlElementInfo for SharedStringsPart {
    fn tag_name() -> &'static str {
        "sst"
    }

    fn element_type() -> OpenXmlElementType {
        OpenXmlElementType::Root
    }
}

impl OpenXmlFromDeserialize for SharedStringsPart {}

impl fmt::Display for SharedStringsPart {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut container = Vec::new();
        let mut cursor = std::io::Cursor::new(&mut container);
        self.write(&mut cursor).expect("write xml to memory error");
        let s = String::from_utf8_lossy(&container);
        write!(f, "{}", s)?;
        Ok(())
    }
}

impl SharedStringsPart {
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
        let mut xml = quick_xml::Writer::new(writer);
        use quick_xml::events::attributes::Attribute;
        use quick_xml::events::*;

        // 1. write decl
        xml.write_event(Event::Decl(BytesDecl::new(
            b"1.0",
            Some(b"UTF-8"),
            Some(b"yes"),
        )))?;
        //quick_xml::se::to_writer(xml.inner(), self).unwrap();

        // 2. start types element
        let mut elem = BytesStart::borrowed_name(SHARED_STRINGS_TAG.as_bytes());

        //elem.extend_attributes(vec![nsattr!(SHARED_STRINGS), nsattr!(VT)]);
        elem.extend_attributes(self.namespaces.to_xml_attributes());
        elem.extend_attributes(vec![
            Attribute {
                key: b"count",
                value: format!("{}", self.count).as_bytes().into(),
            },
            Attribute {
                key: b"uniqueCount",
                value: format!("{}", self.unique_count).as_bytes().into(),
            },
        ]);
        xml.write_event(Event::Start(elem))?;
        for si in &self.strings {
            let elem = BytesStart::borrowed_name(SHARED_STRING_TAG.as_bytes());
            xml.write_event(Event::Start(elem))?;
            quick_xml::se::to_writer(xml.inner(), &si.t)?;
            let end = BytesEnd::borrowed(SHARED_STRING_TAG.as_bytes());
            xml.write_event(Event::End(end))?;
        }

        // // ends types element.
        let end = BytesEnd::borrowed(SHARED_STRINGS_TAG.as_bytes());
        xml.write_event(Event::End(end))?;
        Ok(())
    }
}
#[test]
fn test_de() {
    const raw: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="14" uniqueCount="5"><si><t>name</t></si><si><t>age</t></si><si><t>张三</t></si><si><t>李四</t></si><si><t>王五</t></si></sst>"#;
    println!("{}", raw);
    let value: SharedStringsPart = quick_xml::de::from_str(raw).unwrap();
    println!("{:?}", value);
    let display = format!("{}", value);
    println!("{}", display);
    assert_eq!(raw, display);
}
