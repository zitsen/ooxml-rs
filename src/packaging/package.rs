use crate::error::OoxmlError;
use crate::packaging::content_type::ContentTypes;
use crate::packaging::part::OpenXmlPart;

use std::collections::BTreeMap;
use std::collections::LinkedList;
use std::fs::File;
use std::io::prelude::*;
use std::path::Path;

use zip::ZipArchive;

use linked_hash_map::LinkedHashMap;
use serde::{Deserialize, Serialize};
use url::Url;

/// A common OpenXML package manager, compatible with any [OpenXML Package Convertion]()
#[derive(Debug, Clone)]
pub struct OpenXmlPackage {
    content_types: ContentTypes,
    parts: LinkedHashMap<String, OpenXmlPart>,
}

impl OpenXmlPackage {
    /// Open a OpenXML file path, parse everything into the memory.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self, OoxmlError> {
        let mut file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }
    /// Parse OpenXML package from reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self, OoxmlError> {
        let mut zip = ZipArchive::new(reader)?;
        let mut parts = LinkedHashMap::new();
        let mut content_types_id = None;
        for i in 0..zip.len() {
            let mut file = zip.by_index(i)?;
            // skip directory entries seems ok.
            if file.is_dir() {
                continue;
            }
            let filename = file.name();
            if filename == crate::packaging::content_type::CONTENT_TYPES_FILE {
                content_types_id = Some(i);
                continue;
            }
            let filename = filename.to_string();
            let uri = std::path::PathBuf::from(&filename);
            let part = OpenXmlPart::from_reader(uri, &mut file)?;
            parts.insert(filename, part);
        }
        if content_types_id.is_none() {
            return Err(OoxmlError::PackageContentTypeError);
        }

        let content_types = {
            let mut part =
                zip.by_index(content_types_id.expect("content types file id is not none"))?;
            let mut xml = String::new();
            part.read_to_string(&mut xml)?;
            ContentTypes::parse_from_xml_str(&xml)
        };

        let package = OpenXmlPackage {
            content_types,
            parts,
        };
        Ok(package)
    }

    /// Save as file, write zip package for office document.
    pub fn save_as<P: AsRef<Path>>(&self, path: P) -> Result<(), OoxmlError> {
        let mut file = File::create(path)?;
        self.write(&mut file)?;
        Ok(())
    }

    pub fn write<W: Write + Seek>(&self, writer: W) -> Result<(), OoxmlError> {
        let mut zip = zip::ZipWriter::new(writer);
        let options =
            zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file(crate::packaging::content_type::CONTENT_TYPES_FILE, options)?;
        self.content_types.write(&mut zip)?;
        for (path, part) in self.parts.iter() {
            zip.start_file(path, options)?;
            zip.write(part.as_part_bytes())?;
        }
        zip.flush();
        Ok(())
    }

    pub fn get_parts() {}
    pub fn create_part() {}

    pub fn flush() {}

    pub fn create_relationship() {}

    pub fn delete_relationship() {}

    pub fn get_relationships() {}

    pub fn get_relationships_by_type(relationship_type: String) {}

    pub fn relationship_exist(id: String) -> bool {
        unimplemented!()
    }

    pub fn create_part_core() {}

    pub fn delete_part_core() {}
}

#[test]
fn open() {
    let package = OpenXmlPackage::open("examples/excel-demo/demo.xlsx").unwrap();
    // write back
    package.save_as("tests/write-back.xlsx").unwrap();
    std::fs::remove_file("tests/write-back.xlsx").unwrap();
}
