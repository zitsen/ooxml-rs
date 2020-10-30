mod reference;

pub use reference::ReferenceRelationship;
pub struct RelationshipId(String);
pub struct ExternalRelationship(ReferenceRelationship);
pub struct DataPartReferenceRelationship(ReferenceRelationship);

pub struct HyperlinkRelationship(ReferenceRelationship);

pub struct DataPartReferenceRelationships {}
pub struct ExternalRelationships {}

pub struct HyperlinkRelationships {}

use std::fmt;
use std::fs::File;
use std::io::prelude::*;
use std::path::Path;

use linked_hash_map::LinkedHashMap;
use serde::de::{Deserialize, Deserializer, MapAccess, Visitor};

use crate::error::OoxmlError;


pub const RELATIONSHIPS_FILE: &'static str = "_rels/.rels";

const XMLNS_ATTRIBUTE_NAME: &'static str = "xmlns";
const RELATIONSHIP_NAMESPACE_URI: &'static str = "http://schemas.openxmlformats.org/package/2006/relationships";
// const XMLNS_R_ATTRIBUTE_NAME: &'static str = "xmlns:r";
const RELATIONSHIP_TAG_NAME: &'static str = "Relationship";
const RELATIONSHIPS_TAG_NAME: &'static str = "Relationships";

#[derive(Debug, PartialEq, Default, Clone, serde::Deserialize, serde::Serialize)]
#[serde(rename_all = "PascalCase")]
pub struct Relationship {
    id: String,
    r#type: String,
    target: String,
}
#[derive(Debug, PartialEq, Default, Clone)]
pub struct Relationships {
    relationships: LinkedHashMap<String, Relationship>,
}

struct RelationshipsVisitor;

impl<'de> Visitor<'de> for RelationshipsVisitor {
    // The type that our Visitor is going to produce.
    type Value = Relationships;

    // Format a message stating what data this Visitor expects to receive.
    fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
        formatter.write_str("relationships deserializing error")
    }

    // Deserialize MyMap from an abstract "map" provided by the
    // Deserializer. The MapAccess input is a callback provided by
    // the Deserializer to let us see each entry in the map.
    fn visit_map<M>(self, mut access: M) -> Result<Self::Value, M::Error>
    where
        M: MapAccess<'de>,
    {
        let mut types: Relationships = Relationships::default();
        while let Some(key) = access.next_key()? {
            let key: String = key;
            match &key {
                s if s == XMLNS_ATTRIBUTE_NAME => {
                    let _xmlns: String = access.next_value()?;
                }
                s if s == RELATIONSHIPS_TAG_NAME => {
                    unreachable!();
                }
                s if s == RELATIONSHIP_TAG_NAME => {
                    let v: Relationship = access.next_value()?;
                    types.add_relationship(v);
                    //types.add_default_element(v.extension, v.content_type);
                }
                _ => {
                    println!("{}", key);
                    unreachable!("content type unsupport!");
                }
            }
        }
        Ok(types)
    }
}
impl<'de> Deserialize<'de> for Relationships {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        // Instantiate our Visitor and ask the Deserializer to drive
        // it over the input data, resulting in an instance of MyMap.
        deserializer.deserialize_map(RelationshipsVisitor)
    }
}

impl fmt::Display for Relationships {
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
impl Relationships {
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
        let mut file = File::create(path)?;
        self.write(file)
    }

    /// Write to an writer
    pub fn write<W: std::io::Write>(&self, mut writer: W) -> Result<(), OoxmlError> {
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
        let mut elem = BytesStart::borrowed_name(RELATIONSHIPS_TAG_NAME.as_bytes());
        let ns = Attribute {
            key: XMLNS_ATTRIBUTE_NAME.as_bytes(),
            value: RELATIONSHIP_NAMESPACE_URI.as_bytes().into(),
        };
        elem.extend_attributes(vec![ns]);
        xml.write_event(Event::Start(elem))?;

        // 3. write default entries
        for (_key, value) in &self.relationships {
            quick_xml::se::to_writer(xml.inner(), value).expect("[inernal] serde to xml error");
        }

        // ends types element.
        let end = BytesEnd::borrowed(RELATIONSHIPS_TAG_NAME.as_bytes());
        xml.write_event(Event::End(end))?;
        Ok(())
    }

    pub fn add_relationship(&mut self, relationship: Relationship) {
        self.relationships.insert(relationship.id.clone(), relationship);
    }

    pub fn is_empty(&self) -> bool {
        self.relationships.is_empty()
    }

    pub fn contains(&self, id: &str) -> bool {
        self.relationships.contains_key(id)
    }
}

#[test]
fn test_de() {
    let raw = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>"#;
    println!("{}", raw);
    let relationships: Relationships = quick_xml::de::from_str(raw).unwrap();
    println!("{:?}", relationships);
    let display = format!("{}", relationships);
    assert_eq!(raw, display);
}
