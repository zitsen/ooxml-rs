mod container;
mod pair;

use crate::error::OoxmlError;

use std::io::Cursor;
use std::io::prelude::*;
use std::path::PathBuf;

use url::Url;
#[derive(Debug, Clone, Default)]
pub struct OpenXmlPart {
    uri: PathBuf,
    content_type: Option<String>,
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

    pub fn as_part_bytes(&self) -> &[u8] {
        self.raw.get_ref()
    }
}
