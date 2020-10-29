use std::io::prelude::*;
use std::path::Path;
use std::fs::File;
use std::cell::{Ref, RefCell};

use zip::ZipArchive;
use std::collections::BTreeMap;
use serde::{Serialize, Deserialize};

use crate::OoxmlError;

pub struct Relationships {
    
}

pub struct OpenXMLPart<'oxml, R: Read + Seek> {
    package: Ref<'oxml, OpenXMLPackage<R>>,
    content_type: String,
}

pub struct OpenXMLPackage<R: Read + Seek> {
    zip: RefCell<ZipArchive<R>>,
}

impl OpenXMLPackage<File> {
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self, OoxmlError> {
        let mut file = std::fs::File::open(path)?;
        Self::new(file)
    }
}

impl<R: Read + Seek> OpenXMLPackage<R> {
    pub fn new(reader: R) -> Result<Self, OoxmlError> {
        let mut zip = ZipArchive::new(reader)?;

        macro_rules! read {
            ($xml:tt, $name:expr) => {{
                let mut file = zip.by_name($name)?;
                let mut buffer = String::new();
                file.read_to_string(&mut buffer)?;
                buffer
            }};
        }

        Ok(Self { zip: RefCell::new(zip) })
    }

    pub fn parts(&self) {
        unimplemented!()
    }
}