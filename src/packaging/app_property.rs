use std::{borrow::Cow, io::Write};

use super::element::*;
use super::namespace::Namespaces;
use super::variant::*;

use quick_xml::events::attributes::Attribute;
use serde::{Deserialize, Serialize};

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
    value: Variant,
}

#[test]
fn serde_titles_of_parts() {
    let v = Variant::VtVector {
        size: 2,
        base_type: "lpstr".into(),
        variants: vec![
            Variant::VtLpstr("Sheet1".into()),
            Variant::VtLpstr("Sheet2".into()),
        ],
    };
    let xml = quick_xml::se::to_string(&v).unwrap();
    assert_eq!(
        xml,
        r#"<vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector>"#
    );

    let v = TitlesOfParts { value: v };
    let xml = quick_xml::se::to_string(&v).unwrap();
    assert_eq!(
        xml,
        r#"<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts>"#
    );

    #[derive(Debug, Serialize, Deserialize, PartialEq)]
    #[serde(rename = "Properties", rename_all = "PascalCase")]
    struct A {
        titles_of_parts: TitlesOfParts,
    }
    let v = A { titles_of_parts: v };
    let xml = quick_xml::se::to_string(&v).unwrap();
    assert_eq!(
        xml,
        r#"<Properties><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#
    );
    let v2: A = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, v2);
}
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HeadingPairs {
    #[serde(rename = "$value")]
    variant: Variant,
}
/// App properties
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "Properties", rename_all = "PascalCase")]
pub struct AppProperties {
    #[serde(flatten, skip_serializing)]
    pub namespaces: Namespaces,
    pub application: Option<Application>,
    pub heading_pairs: Option<HeadingPairs>,
    pub titles_of_parts: Option<TitlesOfParts>,
    pub lines: Option<String>,
    pub links_up_to_date: Option<String>,
    pub local_name: Option<String>,
    pub company: Option<String>,
    pub template: Option<String>,
    pub manager: Option<String>,
    pub pages: Option<String>,
}

impl OpenXmlElementInfo for AppProperties {
    fn tag_name() -> &'static str {
        APP_PROPERTIES_TAG
    }

    fn element_type() -> OpenXmlElementType {
        OpenXmlElementType::Root
    }
}

impl OpenXmlSerialize for AppProperties {
    fn namespaces(&self) -> Option<Cow<Namespaces>> {
        Some(Cow::Borrowed(&self.namespaces))
    }
    fn attributes(&self) -> Option<Vec<Attribute>> {
        None
    }
    fn write_inner<W: Write>(&self, writer: W) -> crate::error::Result<()> {
        let mut writer = quick_xml::Writer::new(writer);

        macro_rules! se_field {
            ($field:ident) => {
                if let Some($field) = &self.$field {
                    quick_xml::se::to_writer(writer.inner(), $field)?;
                }
            };
        }
        se_field!(application);
        se_field!(heading_pairs);
        se_field!(titles_of_parts);

        macro_rules! string_field {
            ($field:ident, $tag:literal) => {
                paste::paste! {
                    if let Some($field) = &self.$field {
                        use quick_xml::events::*;
                        let start = BytesStart::borrowed_name($tag);
                        writer.write_event(Event::Start(start))?;
                        writer.write_event(Event::Text(BytesText::from_plain_str($field)))?;
                        let end = BytesEnd::borrowed($tag);
                        writer.write_event(Event::End(end))?;
                    }
                }
            };
        }
        // string_filed!(links_up_to_date, b"LinksUpToDate");
        string_field!(company, b"Company");
        string_field!(template, b"Template");
        string_field!(manager, b"Manager");
        string_field!(pages, b"Pages");

        Ok(())
    }
}
impl OpenXmlDeserializeDefault for AppProperties {}

#[test]
fn serde() {
    let raw = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>WPS 表格</Application><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#;
    let v: AppProperties = AppProperties::from_xml_str(raw).unwrap();
    println!("{:?}", v);
    // let xml = quick_xml::se::to_string(&v).unwrap();
    let xml = v.to_xml_string().unwrap();
    assert_eq!(raw, xml);
}
