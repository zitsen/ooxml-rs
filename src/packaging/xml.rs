use std::marker::Sized;
pub trait FromXml: Sized {
    fn from_xml(s: &str) -> Result<Self, crate::error::OoxmlError>;
}

pub trait ToXml {
    fn to_xml(&self) -> String;
}
