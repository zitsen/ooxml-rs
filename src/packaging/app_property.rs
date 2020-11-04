use crate::error::OoxmlError;
use crate::packaging::content_type::ContentType;
use crate::packaging::variant::*;
use crate::packaging::xml::{FromXml, ToXml};

use serde::{Deserialize, Serialize};
use linked_hash_map::LinkedHashMap;

use std::io::prelude::*;
use std::{borrow::Cow, fmt, fs::File, path::Path};

use super::xml::OpenXmlFromDeserialize;

pub const APP_PROPERTIES_URI: &str = "docProps/app.xml";

pub const APP_PROPERTIES_TAG: &str = "Properties";
pub const APP_PROPERTIES_NAMESPACE_ATTRIBUTE: &str = "xmlns";
pub const APP_PROPERTIES_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/app-properties";

pub const APP_PROPERTY_TAG: &str = "property";

pub const VT_NAMESPACE_ATTRIBUTE: &str = "xmlns:vt";
pub const VT_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Application(String);

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TitlesOfParts {
    #[serde(rename(deserialize = "$value", serialize = "vt:vector"))]
    value: Variant
}

#[test]
fn serde_titles_of_parts() {
    let v = Variant::VtVector { size: 2, base_type: "lpstr".into(), variants: vec![
        Variant::VtLpstr("Sheet1".into()),
        Variant::VtLpstr("Sheet2".into()),
    ]};
    let xml = quick_xml::se::to_string(&v).unwrap();
    assert_eq!(xml, r#"<vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector>"#);

    let v = TitlesOfParts { value: v };
    let xml = quick_xml::se::to_string(&v).unwrap();
    assert_eq!(xml, r#"<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts>"#);

    #[derive(Debug, Serialize, Deserialize, PartialEq)]
    #[serde(rename = "Properties", rename_all = "PascalCase")]
    struct A {
        titles_of_parts: TitlesOfParts,
    }
    let v = A { titles_of_parts: v };
    let xml = quick_xml::se::to_string(&v).unwrap();
    assert_eq!(xml, r#"<Properties><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#);
    let v2: A = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, v2);
}
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HeadingPairs {
    #[serde(rename = "$value")]
    variant: Variant
}
/// App properties
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "Properties", rename_all = "PascalCase")]
pub struct AppProperties {
    #[serde(rename = "xmlns")]
    nlms: String,
    #[serde(rename = "xmlns:vt")]
    nlms_vt: String,
    application: Option<Application>,
    heading_pairs: Option<HeadingPairs>,
    titles_of_parts: Option<TitlesOfParts>,
    lines: Option<String>,
    links_up_to_date: Option<String>,
    local_name: Option<String>,
    company: Option<String>,
    template: Option<String>,
    manager: Option<String>,
    pages: Option<String>,
}

impl OpenXmlFromDeserialize for AppProperties {}

#[test]
fn serde() {
    let raw = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>WPS 表格</Application><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#;
    let v = AppProperties::from_xml_str(raw).unwrap();
    println!("{:?}", v);
    let xml = quick_xml::se::to_string(&v).unwrap();
    const decl: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
    assert_eq!(raw, format!("{}{}", decl, xml));
}
