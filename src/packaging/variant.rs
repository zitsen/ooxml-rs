//! The main entry of variant is `VariantTypes`.

use serde::{Deserialize, Serialize};

/// OpenXML variant types
///
/// See also: http://www.datypic.com/sc/ooxml/ns-docPropsVTypes.html
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "vt:variant")]
pub enum Variant {
    #[serde(rename = "vt:vector")]
    VtVector {
        size: usize,
        #[serde(rename = "baseType")]
        base_type: String,
        #[serde(rename = "$value")]
        variants: Vec<Variant>,
    },
    #[serde(rename = "vt:variant")]
    VtVariant {
        #[serde(rename = "$value")]
        value: Box<Variant>,
    },
    // //VTArray(Vec<impl Any>),
    // //VTBlob(&[u8]),
    // VTEmpty,
    #[serde(rename = "vt:null")]
    VtNull,
    // VTByte(u8),
    // VTShort(i16),
    // VTInt32(i32),
    // //#[serde(rename = "vt:int64")]
    // VTInt64(i64),
    #[serde(rename = "vt:i1")]
    VtI1(i8),
    #[serde(rename = "vt:i2")]
    VtI2(i8),
    #[serde(rename = "vt:i4")]
    VtI4(i8),
    #[serde(rename = "vt:i8")]
    VtI8(i8),
    #[serde(rename = "vt:lpstr")]
    VtLpstr(String),
    #[serde(rename = "vt:lpwstr")]
    VtLpwstr(String),
}

impl Default for Variant {
    fn default() -> Self {
        Variant::VtNull
    }
}

#[test]
fn serde_vt_variant() {
    let v = Variant::VtLpwstr("text".into());
    let xml = quick_xml::se::to_string(&v).unwrap();
    println!("{}", xml);
    assert_eq!(format!("{}", xml), "<vt:lpwstr>text</vt:lpwstr>");
    let vd = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, vd);

    let v = Variant::VtVariant { value: Box::new(v) };
    let xml = quick_xml::se::to_string(&v).unwrap();
    println!("{}", xml);
    assert_eq!(xml, "<vt:variant><vt:lpwstr>text</vt:lpwstr></vt:variant>");
    let vd = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, vd);
}
#[test]
fn serde_vt_vector() {
    let v = Variant::VtLpwstr("text".into());
    let xml = quick_xml::se::to_string(&v).unwrap();
    println!("{}", xml);
    assert_eq!(format!("{}", xml), "<vt:lpwstr>text</vt:lpwstr>");
    let vd = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, vd);

    let v = Variant::VtVector {
        size: 1,
        base_type: "lpwstr".into(),
        variants: vec![v],
    };
    let xml = quick_xml::se::to_string(&v).unwrap();
    println!("{}", xml);
    assert_eq!(
        xml,
        r#"<vt:vector size="1" baseType="lpwstr"><vt:lpwstr>text</vt:lpwstr></vt:vector>"#
    );
    let vd = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, vd);
}

#[test]
fn variant_serde() {
    // assert_eq!(format!("{}", xml), "<vt:lpwstr>text</vt:lpwstr>");
    let v = Variant::VtLpwstr("text".into());
    let xml = quick_xml::se::to_string(&v).unwrap();
    println!("{}", xml);
    assert_eq!(format!("{}", xml), "<vt:lpwstr>text</vt:lpwstr>");
    let vd = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, vd);

    let v = Variant::VtVariant { value: Box::new(v) };
    let xml = quick_xml::se::to_string(&v).unwrap();
    println!("{}", xml);
    assert_eq!(xml, "<vt:variant><vt:lpwstr>text</vt:lpwstr></vt:variant>");
    let vd = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, vd);

    let v = Variant::VtLpwstr("text".into());
    let xml = quick_xml::se::to_string(&v).unwrap();
    println!("{}", xml);
    assert_eq!(xml, "<vt:lpwstr>text</vt:lpwstr>");
    let vd = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, vd);

    let v = Variant::VtVariant {
        value: Box::new(Variant::VtI4(2)),
    };
    //let v = Variant::VTVector(vec![v.clone(), v.clone()]);
    let xml = quick_xml::se::to_string(&v).unwrap();
    println!("{}", xml);
    assert_eq!(xml, "<vt:variant><vt:i4>2</vt:i4></vt:variant>");
    let vd = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, vd);

    let v = Variant::VtVector {
        size: 2,
        base_type: "variant".into(),
        variants: vec![v],
    };
    let xml = quick_xml::se::to_string(&v).unwrap();
    println!("{}", xml);
    assert_eq!(
        xml,
        r#"<vt:vector size="2" baseType="variant"><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector>"#
    );
    let vd = quick_xml::de::from_str(&xml).unwrap();
    assert_eq!(v, vd);
}
