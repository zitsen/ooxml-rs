use linked_hash_map::LinkedHashMap;

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Namespaces(LinkedHashMap<String, String>);

impl Namespaces {
    pub fn new<S: Into<String>>(uri: S) -> Self {
        let mut ns = Namespaces::default();
        ns.set_default_namespace(uri);
        ns
    }
    pub fn add_namespace<S1: Into<String>, S2: Into<String>>(
        &mut self,
        decl: S1,
        uri: S2,
    ) {
        self.0.insert(decl.into(), uri.into());
    }
    pub fn remove_namespace(&mut self, decl: &str) {
        self.0.remove(decl);
    }
    pub fn set_default_namespace<S: Into<String>>(&mut self, uri: S) {
        self.add_namespace("xmlns", uri);
    }
}

use quick_xml::events::attributes::Attribute;

impl Namespaces {
    pub fn to_xml_attributes(&self) -> Vec<Attribute> {
        self.0.iter().map(|(k, v)| {
            Attribute {
                key: k.as_bytes(),
                value: v.as_bytes().into(),
            }
        }).collect()
    }
}

use std::fmt;
use serde::de::{Visitor, MapAccess, Deserialize, Deserializer};
struct ContentTypesVisitor;

impl<'de> Visitor<'de> for ContentTypesVisitor {
    // The type that our Visitor is going to produce.
    type Value = Namespaces;

    // Format a message stating what data this Visitor expects to receive.
    fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
        formatter.write_str("unexpected namespace attribute")
    }

    fn visit_map<M>(self, mut access: M) -> Result<Self::Value, M::Error>
    where
        M: MapAccess<'de>,
    {
        let mut ns = Namespaces::default();
        while let Some(key) = access.next_key()? {
            let key: String = key;
            match key {
                s if s.starts_with("xmlns") => {
                    let xmlns: String = access.next_value()?;
                    ns.add_namespace(s, xmlns);
                }
                s => {
                    log::debug!("unrecognized namespace: {}!", s);
                    //unreachable!(format!("unrecognized namespace: {}!", s));
                }
            }
        }
        Ok(ns)
    }
}

impl<'de> Deserialize<'de> for Namespaces {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        // Instantiate our Visitor and ask the Deserializer to drive
        // it over the input data, resulting in an instance of MyMap.
        deserializer.deserialize_map(ContentTypesVisitor)
    }
}

use serde::ser::{Serialize, Serializer, SerializeMap};

impl Serialize for Namespaces {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        let mut map = serializer.serialize_map(Some(self.0.len()))?;
        for (k, v) in &self.0 {
            map.serialize_entry(k, v)?;
        }
        map.end()
    }
}

#[test]
fn de() {
    #[derive(Debug, serde::Serialize, serde::Deserialize)]
    #[serde(rename = "Relationships")]
    struct S {
        #[serde(flatten)]
        namespaces: Namespaces,
        #[serde(rename = "$value")]
        a: String,
    }

    let xml = r#"<Relationships xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">Text</Relationships>"#;
    let s: S = quick_xml::de::from_str(xml).unwrap();
    println!("{:?}", s);
    let xml = quick_xml::se::to_string(&s).unwrap();
    println!("{}", xml);
}