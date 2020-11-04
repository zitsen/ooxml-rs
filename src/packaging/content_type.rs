use std::fmt;
use std::fs::File;
use std::io::prelude::*;
use std::path::Path;

use linked_hash_map::LinkedHashMap;
use serde::de::{Deserialize, Deserializer, MapAccess, Visitor};

use crate::error::OoxmlError;

pub type ContentType = String;

#[derive(Debug, serde::Deserialize, serde::Serialize, PartialEq)]
#[serde(rename_all = "PascalCase")]
struct Default {
    extension: String,
    content_type: ContentType,
}
#[derive(Debug, serde::Deserialize, serde::Serialize, PartialEq)]
#[serde(rename_all = "PascalCase")]
struct Override {
    part_name: String,
    content_type: ContentType,
}
// #[derive(Debug, serde::Deserialize, serde::Serialize, PartialEq)]
// #[serde(rename_all = "PascalCase", untagged)]
// enum ContentTypeItem {
//     Default(Default),
//     Override(Override),
// }

// impl ContentTypeItem {
//     pub fn new(content_type: &str) -> Self {
//         let mime: mime::Mime = content_type.parse().expect("unrecognized content type");
//         Self::new_default(mime.subtype().as_str(), content_type)
//     }
//     pub fn new_default(extension: &str, content_type: &str) -> Self {
//         ContentTypeItem::Default(Default {
//             extension: extension.into(),
//             content_type: content_type.into(),
//         })
//     }
//     pub fn new_override(part_name: &str, content_type: &str) -> Self {
//         ContentTypeItem::Override(Override {
//             part_name: part_name.into(),
//             content_type: content_type.into(),
//         })
//     }
// }
// #[test]
// fn test_content_type_serde() {
//     let default = ContentTypeItem::new("image/png");

//     let ser = quick_xml::se::to_string(&default).unwrap();
//     assert_eq!(ser, r#"<Default Extension="png" ContentType="image/png"/>"#);
//     let de: ContentTypeItem = quick_xml::de::from_str(&ser).unwrap();
//     assert_eq!(default, de);
//     println!("{:?}", de);
// }

#[derive(Debug, PartialEq, Default, Clone)]
pub struct ContentTypes {
    defaults: LinkedHashMap<String, ContentType>,
    overrides: LinkedHashMap<String, ContentType>,
}

struct ContentTypesVisitor;

impl<'de> Visitor<'de> for ContentTypesVisitor {
    // The type that our Visitor is going to produce.
    type Value = ContentTypes;

    // Format a message stating what data this Visitor expects to receive.
    fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
        formatter.write_str("a very special map")
    }

    // Deserialize MyMap from an abstract "map" provided by the
    // Deserializer. The MapAccess input is a callback provided by
    // the Deserializer to let us see each entry in the map.
    fn visit_map<M>(self, mut access: M) -> Result<Self::Value, M::Error>
    where
        M: MapAccess<'de>,
    {
        let mut types: ContentTypes = ContentTypes::default();
        while let Some(key) = access.next_key()? {
            let key: String = key;
            match &key {
                s if s == XmlnsAttributeName => {
                    let _xmlns: String = access.next_value()?;
                }
                s if s == TypesTagName => {
                    //let v: Vec<String> = access.next_value()?;
                    unreachable!();
                }
                s if s == DefaultTagName => {
                    let v: Default = access.next_value()?;
                    types.add_default_element(v.extension, v.content_type);
                }
                s if s == OverrideTagName => {
                    let v: Override = access.next_value()?;
                    types.add_override_element(v.part_name, v.content_type);
                }
                _ => {
                    unreachable!("content type unsupport!");
                }
            }
        }
        Ok(types)
    }
}
impl<'de> Deserialize<'de> for ContentTypes {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        // Instantiate our Visitor and ask the Deserializer to drive
        // it over the input data, resulting in an instance of MyMap.
        deserializer.deserialize_map(ContentTypesVisitor)
    }
}

pub const CONTENT_TYPES_FILE: &'static str = "[Content_Types].xml";
pub const TypesNamespaceUri: &'static str =
    "http://schemas.openxmlformats.org/package/2006/content-types";
pub const TypesTagName: &'static str = "Types";
pub const DefaultTagName: &'static str = "Default";
pub const OverrideTagName: &'static str = "Override";
pub const PartNameAttributeName: &'static str = "PartName";
pub const ExtensionAttributeName: &'static str = "Extension";
pub const ContentTypeAttributeName: &'static str = "ContentType";
pub const XmlnsAttributeName: &'static str = "xmlns";

impl fmt::Display for ContentTypes {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut container = Vec::new();
        let mut cursor = std::io::Cursor::new(&mut container);
        self.write(&mut cursor).expect("write xml to memory error");
        let s = String::from_utf8_lossy(&container);
        write!(f, "{}", s)?;
        // /write!("{}", String::from_utf8(cursor.into_inner().into));
        Ok(())
    }
}
impl ContentTypes {
    /// Parse content types data from an xml reader.
    pub fn parse_from_xml_reader<R: BufRead>(reader: R) -> Self {
        quick_xml::de::from_reader(reader).unwrap()
    }

    /// Parse content types data from an xml str.
    pub fn parse_from_xml_str(reader: &str) -> Self {
        quick_xml::de::from_str(reader).unwrap()
    }

    /// follow OpenXML SDK function definitions, but not implemented.
    pub fn add_content_type() {}
    pub fn get_content_type() {}
    pub fn delete_content_type(&mut self, _content_type: &ContentType) {

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

        // 2. start types element
        let mut elem = BytesStart::borrowed_name(TypesTagName.as_bytes());
        let ns = Attribute {
            key: XmlnsAttributeName.as_bytes(),
            value: TypesNamespaceUri.as_bytes().into(),
        };
        elem.extend_attributes(vec![ns]);
        xml.write_event(Event::Start(elem))?;

        // 3. write default entries
        for (key, value) in &self.defaults {
            xml.write_event(Event::Empty(
                BytesStart::borrowed_name(DefaultTagName.as_bytes()).with_attributes(vec![
                    Attribute {
                        key: ExtensionAttributeName.as_bytes(),
                        value: key.as_bytes().into(),
                    },
                    Attribute {
                        key: ContentTypeAttributeName.as_bytes(),
                        value: value.as_bytes().into(),
                    },
                ]),
            ))?;
        }

        // 4. write override entries
        for (key, value) in &self.overrides {
            xml.write_event(Event::Empty(
                BytesStart::borrowed_name(OverrideTagName.as_bytes()).with_attributes(vec![
                    Attribute {
                        key: PartNameAttributeName.as_bytes(),
                        value: key.as_bytes().into(),
                    },
                    Attribute {
                        key: ContentTypeAttributeName.as_bytes(),
                        value: value.as_bytes().into(),
                    },
                ]),
            ))?;
        }
        // ends types element.
        let end = BytesEnd::borrowed(TypesTagName.as_bytes());
        xml.write_event(Event::End(end))?;
        Ok(())
    }

    pub fn add_default_element(&mut self, extension: String, content_type: ContentType) {
        self.defaults.insert(extension, content_type);
    }

    pub fn add_override_element(&mut self, part_name: String, content_type: ContentType) {
        self.overrides.insert(part_name, content_type);
    }

    pub fn is_empty(&self) -> bool {
        self.defaults.is_empty() && self.overrides.is_empty()
    }
}

#[test]
fn test_de() {
    let raw = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/><Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;
    println!("{}", raw);
    let content_types: ContentTypes = quick_xml::de::from_str(raw).unwrap();
    println!("{:?}", content_types);
    let display = format!("{}", content_types);
    assert_eq!(raw, display);
}
