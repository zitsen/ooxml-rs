use crate::error::OoxmlError;
use crate::packaging::content_type::{ContentType, ContentTypes};
use crate::packaging::custom_property::CustomProperties;
use crate::packaging::part::OpenXmlPart;
use crate::packaging::property::Properties;
pub use crate::packaging::relationship::Relationships;
use crate::packaging::app_property::AppProperties;
use crate::packaging::xml::*;



use std::fs::File;
use std::io::prelude::*;
use std::path::Path;

use zip::ZipArchive;

use linked_hash_map::LinkedHashMap;



use crate::packaging::content_type::CONTENT_TYPES_FILE;
use crate::packaging::custom_property::CUSTOM_PROPERTIES_URI;
use crate::packaging::app_property::APP_PROPERTIES_URI;
use crate::packaging::property::CORE_PROPERTIES_URI;
use crate::packaging::relationship::RELATIONSHIPS_FILE;

/// A common OpenXML package manager, compatible with any [OpenXML Package Convertion]()
///
/// FIXME(@zitsen): A dirty dict representing data change should be added.
#[derive(Debug, Clone, Default)]
pub struct OpenXmlPackage {
    content_types: ContentTypes,
    relationships: Relationships,
    app_properties: AppProperties,
    properties: Properties,
    custom_properties: Option<CustomProperties>,
    parts: LinkedHashMap<String, OpenXmlPart>,
}

impl OpenXmlPackage {
    /// Open a OpenXML file path, parse everything into the memory.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self, OoxmlError> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }
    /// Parse OpenXML package from reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self, OoxmlError> {
        let mut zip = ZipArchive::new(reader)?;
        let mut package = OpenXmlPackage::default();
        let mut content_types_id = None;
        for i in 0..zip.len() {
            let mut file = zip.by_index(i)?;
            // skip directory entries seems ok.
            if file.is_dir() {
                continue;
            }
            let filename = file.name().to_string();
            if filename == CONTENT_TYPES_FILE {
                content_types_id = Some(i);
                let mut xml = String::new();
                file.read_to_string(&mut xml)?;
                package.content_types = ContentTypes::parse_from_xml_str(&xml);
                continue;
            } else if filename == RELATIONSHIPS_FILE {
                let mut xml = String::new();
                file.read_to_string(&mut xml)?;
                package.relationships = Relationships::parse_from_xml_str(&xml);
                continue;
            } else if filename == CORE_PROPERTIES_URI {
                let mut xml = String::new();
                file.read_to_string(&mut xml)?;
                package.properties = Properties::parse_from_xml_str(&xml);
                continue;
            } else if filename == CUSTOM_PROPERTIES_URI {
                let mut xml = String::new();
                file.read_to_string(&mut xml)?;
                package.custom_properties = Some(CustomProperties::parse_from_xml_str(&xml));
            } else if filename == APP_PROPERTIES_URI {
                let mut xml = String::new();
                file.read_to_string(&mut xml)?;
                package.app_properties = FromXml::from_xml_str(&xml)?;
            } else {
                let uri = std::path::PathBuf::from(&filename);
                let part = OpenXmlPart::from_reader(uri, &mut file)?;
                package.parts.insert(filename, part);
            }
        }
        if content_types_id.is_none() {
            return Err(OoxmlError::PackageContentTypeError);
        }
        assert!(package.has_content_types());
        assert!(package.has_relationships());
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
        zip.start_file(CONTENT_TYPES_FILE, options)?;
        self.content_types.write(&mut zip)?;
        zip.start_file(RELATIONSHIPS_FILE, options)?;
        self.relationships.write(&mut zip)?;
        zip.start_file(CORE_PROPERTIES_URI, options)?;
        self.properties.write(&mut zip)?;
        if let Some(custom_properties) = &self.custom_properties {
            zip.start_file(CORE_PROPERTIES_URI, options)?;
            custom_properties.write(&mut zip)?;
        }
        for (path, part) in self.parts.iter() {
            zip.start_file(path, options)?;
            zip.write(part.as_part_bytes())?;
        }
        zip.flush()?;
        Ok(())
    }

    pub fn has_content_types(&self) -> bool {
        !self.content_types.is_empty()
    }

    pub fn has_relationships(&self) -> bool {
        !self.relationships.is_empty()
    }

    /// A part is dirty if data has been changed.
    pub fn is_dirty(&self, _uri: &str) -> bool {
        unimplemented!()
    }

    /// Get OpenXML `Part` by uri.
    pub fn get_part(&self, uri: &str) -> Option<&OpenXmlPart> {
        self.parts.get(uri)
    }

    pub fn create_part() {}

    pub fn flush() {}

    pub fn create_relationship() {}

    pub fn delete_relationship() {}

    pub fn get_relationships() {}

    pub fn get_relationships_by_type(_relationship_type: String) {}

    pub fn relationship_exist(&self, id: &str) -> bool {
        self.relationships.contains(id)
    }

    pub fn create_part_core(&mut self, uri: &str, content_type: &ContentType) {
        let part = OpenXmlPart::new_with_content_type(uri, content_type);
        self.parts.insert(uri.into(), part);
        // if content type is not exist, add to content types.
    }
    pub fn create_part_core_with_data(
        &mut self,
        uri: &str,
        content_type: &ContentType,
        data: &[u8],
    ) -> Result<(), OoxmlError> {
        let part = OpenXmlPart::new(uri, content_type, data)?;
        self.parts.insert(uri.into(), part);
        // if content type is not exist, add to content types.
        Ok(())
    }

    /// Delete the part corresponding to the uri specified.
    ///
    /// Delete the content type for this part if it was specified as an override.
    pub fn delete_part_core(&mut self, uri: &str) {
        let part = self.parts.remove(uri);
        if let Some(part) = part {
            if let Some(content_type) = part.content_type() {
                self.content_types.delete_content_type(content_type);
            }
        }
    }
}

#[test]
fn open_and_save() {
    let package = OpenXmlPackage::open("examples/excel-demo/demo.xlsx").unwrap();
    // write back
    package.save_as("tests/write-back.xlsx").unwrap();
    let package = OpenXmlPackage::open("examples/docx-demo/rust-docx-rs.docx").unwrap();
    // write back
    package.save_as("tests/write-back.docx").unwrap();
    // std::fs::remove_file("tests/write-back.xlsx").unwrap();
}
