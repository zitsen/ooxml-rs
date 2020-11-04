
//! There's some marker trait and helper trait for XML (de)serializing.
//!
use std::io::BufReader;
use std::marker::Sized;
use std::{borrow::Cow, fs::File, io::prelude::*, path::Path};

use crate::error::OoxmlError;

use quick_xml::events::attributes::Attribute;
use quick_xml::events::*;

use super::namespace::Namespaces;

/// Leaf for plain text, Node for internal xml element, Root for root element of a Part.
pub enum OpenXmlElementType {
    Leaf,
    Node,
    Root,
}

pub trait OpenXmlLeafElement {}
pub trait OpenXmlNodeElement {}
pub trait OpenXmlRootElement {}

/// Static information of an OpenXml element
pub trait OpenXmlElementInfo: Sized {
    /// XML tag name
    fn tag_name() -> &'static str;

    fn as_bytes_start() -> quick_xml::events::BytesStart<'static> {
        assert!(Self::have_tag_name());
        quick_xml::events::BytesStart::borrowed_name(Self::tag_name().as_bytes())
    }
    fn as_bytes_end() -> quick_xml::events::BytesEnd<'static> {
        assert!(Self::have_tag_name());
        quick_xml::events::BytesEnd::borrowed(Self::tag_name().as_bytes())
    }

    fn is_leaf_text_element() -> bool {
        match Self::element_type() {
            OpenXmlElementType::Leaf => true,
            _ => false,
        }
    }

    /// Element type
    fn element_type() -> OpenXmlElementType {
        OpenXmlElementType::Node
    }

    /// If the element have a tag name.
    ///
    /// Specially, plain text or cdata element does not have tag name.
    fn have_tag_name() -> bool {
        match Self::element_type() {
            OpenXmlElementType::Leaf => false,
            _ => true,
        }
    }

    /// If the element can have namespace declartions.
    fn can_have_namespace_declarations() -> bool {
        match Self::element_type() {
            OpenXmlElementType::Leaf => false,
            _ => true,
        }
    }
    /// If the element can have attributes.
    fn can_have_attributes() -> bool {
        match Self::element_type() {
            OpenXmlElementType::Leaf => false,
            _ => true,
        }
    }
    /// If the element can have children.
    ///
    /// Eg. plain text element cannot have children, so all children changes not works.
    fn can_have_children() -> bool {
        match Self::element_type() {
            OpenXmlElementType::Leaf => false,
            _ => true,
        }
    }
}

/// Common element trait.
pub trait OpenXmlElementExt: OpenXmlElementInfo {
    /// Get element attributes, if have.
    fn attributes(&self) -> Option<Vec<Attribute>>;

    /// Get element namespaces, if have.
    fn namespaces(&self) -> Option<Cow<Namespaces>>;

    /// Serialize to writer
    fn write_inner<W: Write>(&self, writer: W)  -> crate::error::Result<()>;

    fn write_outter<W: Write>(
        &self,
        writer: W,
    ) -> crate::error::Result<()> {
        let mut writer = quick_xml::Writer::new(writer);
        use quick_xml::events::*;

        let mut elem = Self::as_bytes_start();
        if Self::can_have_namespace_declarations() {
            if let Some(ns) = self.namespaces() {
                elem.extend_attributes(ns.to_xml_attributes());
            }
        }
        if Self::can_have_attributes() {
            if let Some(attrs) = self.attributes() {
                elem.extend_attributes(attrs);
            }
        }
        if Self::is_leaf_text_element() {
            writer.write_event(Event::Empty(elem))?;
            return Ok(());
        } else {
            writer.write_event(Event::Start(elem))?;
            self.write_inner(writer.inner())?;
            writer.write_event(Event::End(Self::as_bytes_end()))?;
            return Ok(());
        }
    }
}

