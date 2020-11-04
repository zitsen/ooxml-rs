mod container;
mod pair;

use crate::error::OoxmlError;
use crate::packaging::content_type::ContentType;

use std::io::Cursor;
use std::io::prelude::*;
use std::path::PathBuf;


#[derive(Debug, Clone, Default)]
pub struct OpenXmlPart {
    uri: PathBuf,
    content_type: Option<ContentType>,
    raw: Cursor<Vec<u8>>,
}

impl OpenXmlPart {
    pub fn from_reader<S: Into<PathBuf>, R: Read>(uri: S, mut reader: R) -> Result<Self, OoxmlError> {
        let mut raw = Cursor::new(Vec::new());
        std::io::copy(&mut reader, &mut raw)?;
        let part = Self {
            raw,
            uri: uri.into(),
            ..Default::default()
        };
        Ok(part)
    }

    pub fn new_with_content_type<S: Into<PathBuf>>(uri: S, content_type: impl Into<ContentType>) -> Self {
        Self {
            uri: uri.into(),
            content_type: Some(content_type.into()),
            ..Default::default()
        }
    }

    pub fn new<S: Into<PathBuf>, C: Into<ContentType>, R: Read>(uri: S, _content_type: C, mut reader: R) -> Result<Self, OoxmlError> {
        let mut raw = Cursor::new(Vec::new());
        std::io::copy(&mut reader, &mut raw)?;
        let part = Self {
            raw,
            uri: uri.into(),
            ..Default::default()
        };
        Ok(part)
    }

    pub fn as_part_bytes(&self) -> &[u8] {
        self.raw.get_ref()
    }

    pub fn content_type<>(&self) -> &Option<ContentType> {
        &self.content_type
    }
}
