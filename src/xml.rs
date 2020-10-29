use std::io::prelude::*;
use thiserror::Error;

#[derive(Error, Debug)]
pub enum XmlError {
    #[error("XML parsing error")]
    ReadError(#[from] quick_xml::Error),
}
pub trait FromXML {
    fn from_xml<R: Read>(reader: R) -> Result<Self, XmlError>
    where
        Self: std::marker::Sized;
}
