use crate::error::OoxmlError;

use crate::packaging::variant::Variant;

use serde::{Deserialize, Serialize};

use std::io::prelude::*;
use std::{fmt, fs::File, path::Path};

pub const CUSTOM_PROPERTIES_URI: &str = "docProps/custom.xml";

pub const CUSTOM_PROPERTIES_TAG: &str = "Properties";
pub const CUSTOM_PROPERTIES_NAMESPACE_ATTRIBUTE: &str = "xmlns";
pub const CUSTOM_PROPERTIES_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

pub const CUSTOM_PROPERTY_TAG: &str = "property";

pub const VT_NAMESPACE_ATTRIBUTE: &str = "xmlns:vt";
pub const VT_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

use crate::packaging::namespace::Namespaces;
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all(deserialize = "camelCase"), rename = "property")]
pub struct CustomProperty {
    pub fmtid: String,
    pub pid: String,
    pub name: String,
    #[serde(rename = "$value")]
    value: Variant,
}

#[test]
fn custom_property_de() {
    const xml: &str = r#"<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="KSOProductBuildVer"><vt:lpwstr>2052-11.1.0.9662</vt:lpwstr></property>"#;
    let p: CustomProperty = quick_xml::de::from_str(xml).unwrap();
    println!("{:?}", p);
    let s = quick_xml::se::to_string(&p).unwrap();
    println!("{:?}", s);
    assert_eq!(xml, s);
}

#[test]
fn custom_properties_de() {
    const xml: &str = r#"<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="KSOProductBuildVer"><vt:lpwstr>2052-11.1.0.9662</vt:lpwstr></property>"#;
    let p: Vec<CustomProperty> = quick_xml::de::from_str(xml).unwrap();
    println!("{:?}", p);
    let s = quick_xml::se::to_string(&p).unwrap();
    println!("{:?}", s);
    assert_eq!(xml, s);
}
/// Custom properties
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "Properties")]
pub struct CustomProperties {
    #[serde(flatten)]
    namespaces: Namespaces,
    #[serde(rename = "property")]
    properties: Vec<CustomProperty>,
}

impl fmt::Display for CustomProperties {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut container = Vec::new();
        let mut cursor = std::io::Cursor::new(&mut container);
        self.write(&mut cursor).expect("write xml to memory error");
        let s = String::from_utf8_lossy(&container);
        write!(f, "{}", s)?;
        Ok(())
    }
}

impl CustomProperties {
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

        use quick_xml::events::*;

        // 1. write decl
        xml.write_event(Event::Decl(BytesDecl::new(
            b"1.0",
            Some(b"UTF-8"),
            Some(b"yes"),
        )))?;
        //quick_xml::se::to_writer(xml.inner(), self).unwrap();

        // 2. start types element
        let mut elem = BytesStart::borrowed_name(CUSTOM_PROPERTIES_TAG.as_bytes());

        // macro_rules! nsattr {
        //     ($ns:ident) => {
        //         paste::paste! {
        //             Attribute {
        //                 key: [<$ns _NAMESPACE_ATTRIBUTE>].as_bytes(),
        //                 value: [<$ns _NAMESPACE>].as_bytes().into(),
        //             }
        //         }
        //     };
        // }

        //elem.extend_attributes(vec![nsattr!(CUSTOM_PROPERTIES), nsattr!(VT)]);
        elem.extend_attributes(self.namespaces.to_xml_attributes());
        xml.write_event(Event::Start(elem))?;
        quick_xml::se::to_writer(xml.inner(), &self.properties)?;

        // // ends types element.
        let end = BytesEnd::borrowed(CUSTOM_PROPERTIES_TAG.as_bytes());
        xml.write_event(Event::End(end))?;
        Ok(())
    }
}
#[test]
fn test_de() {
    const raw: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="KSOProductBuildVer"><vt:lpwstr>2052-11.1.0.9662</vt:lpwstr></property></Properties>"#;
    println!("{}", raw);
    let value: CustomProperties = quick_xml::de::from_str(raw).unwrap();
    println!("{:?}", value);
    let display = format!("{}", value);
    println!("{}", display);
    assert_eq!(raw, display);
}
