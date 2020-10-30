use std::{fs::File, io::prelude::*, path::Path};
use std::marker::Sized;
use std::io::BufReader;

use crate::error::OoxmlError;

use quick_xml::events::attributes::Attribute;
use quick_xml::events::*;

pub trait FromXml: Sized {
    /// Parse from an xml stream reader
    fn from_xml_reader<R: BufRead>(reader: R) -> Result<Self, OoxmlError>;

    /// Parse from an xml raw string.
    fn from_xml_str(s: &str) -> Result<Self, OoxmlError> {
        Self::from_xml_reader(s.as_bytes())
    }

    /// Open a OpenXML file path, parse everything into the memory.
    fn from_xml_file<P: AsRef<Path>>(path: P) -> Result<Self, OoxmlError> {
        let file = std::fs::File::open(path)?;
        let reader = BufReader::new(file);
        Self::from_xml_reader(reader)
    }
}

impl<T: serde::de::DeserializeOwned + Sized> FromXml for T {
    fn from_xml_reader<R: BufRead>(reader: R) -> Result<Self, OoxmlError> {
        Ok(quick_xml::de::from_reader(reader)?)
    }
}
pub trait ToXml: Sized {
    /// Implement the write method, the trait will do the rest.
    fn write<W: Write>(&self, writer: W) -> Result<(), OoxmlError>;

    /// Write standalone xml to the writer.
    ///
    /// Will add `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` to writer.
    fn write_standalone<W: Write>(&self, writer: W) -> Result<(), OoxmlError> {
        let mut xml = quick_xml::Writer::new(writer);
        use quick_xml::events::*;
        xml.write_event(Event::Decl(BytesDecl::new(
            b"1.0",
            Some(b"UTF-8"),
            Some(b"yes"),
        )))?;
        self.write(xml.inner())
    }
    /// Write the standalone xml to path
    fn save_as<P: AsRef<Path>>(&self, path: P) -> Result<(), OoxmlError> {
        let file = File::open(path)?;
        self.write_standalone(file)
    }
    /// Output the xml to an Vec<u8> block.
    fn to_xml_bytes(&self) -> Result<Vec<u8>, OoxmlError> {
        let mut container = Vec::new();
        let mut cursor = std::io::Cursor::new(&mut container);
        self.write(&mut cursor)?;
        Ok(container)
    }
    /// Output the xml to string.
    fn to_xml_string(&self) -> Result<String, OoxmlError> {
        let bytes = self.to_xml_bytes()?;
        Ok(String::from_utf8_lossy(&bytes).to_string())
    }
}

impl<T: serde::ser::Serialize + Sized> ToXml for T {
    fn write<W: Write>(&self, writer: W) -> Result<(), OoxmlError> {
        quick_xml::se::to_writer(writer, self)?;
        Ok(())
    }
}

pub trait OpenXmlElement: FromXml + ToXml {
    fn tag(&self) -> &[u8];
    fn namespace_declarations(&self) -> Vec<Attribute>;
    fn add_namespace_declaration(&mut self, prefix: &str, uri: &str);
    fn remove_namespace_declaration(&mut self, prefix: &str);

    //fn markup_compatibility_attributes(&self) -> ();
    fn extended_attributes(&self) -> Vec<Attribute>;
    fn has_attributes(&self) -> bool;
    fn set_attribute(&mut self, attribute: Attribute);
    fn remove_attribute(&mut self, local_name: &str, namespace_uri: &str);
    fn clear_attributes(&mut self);
    
    fn has_children(&self) -> bool;
    //type Children;
    
    //fn append(&mut self, children: OpenXmlChild);
    //fn clear_children(&mut self);
    fn write_children<W: Write>(&self, writer: W) -> Result<(), OoxmlError>;

    // FIXME(@zitsen): need a OpenXmlAttribute definition.
    fn get_attribute(&self, name: &str);
    // FIXME(@zitsen): there's other implmentations for children elements.
    fn children<'xml, X>(&self) -> Box<dyn Iterator<Item = &'xml X>>
    where
        X: OpenXmlElement,
    {
        unimplemented!()
    }

    fn write<W: Write>(&self, mut writer: W) -> Result<(), OoxmlError> {
        let mut xml = quick_xml::Writer::new(&mut writer);
        // 2. start types element
        let tag = self.tag();
        let mut elem = BytesStart::borrowed_name(tag);
        let attrs = self.extended_attributes();
        if !attrs.is_empty() {
            elem.extend_attributes(attrs);
        }

        xml.write_event(Event::Start(elem))?;

        if self.has_children() {
            self.write_children(xml.inner())?;
        }

        // ends types element.
        let end = BytesEnd::borrowed(tag);
        xml.write_event(Event::End(end))?;
        Ok(())
    }
}

/// Expose trait for an element implemented serde deserialize trait, make it simple and fast.
pub trait OpenXmlElementDeserialize: OpenXmlElement + serde::de::DeserializeOwned {
    fn from_xml_reader<R: BufRead>(reader: R) -> Result<Self, OoxmlError> {
        Ok(quick_xml::de::from_reader(reader)?)
    }
    fn from_xml_str(s: &str) -> Result<Self, OoxmlError> {
        Ok(quick_xml::de::from_str(s)?)
    }
}
