use crate::packaging::relationship::*;
use crate::packaging::part::pair::*;

pub trait OpenXmlPartContainer {
    fn data_part_reference_relationships(&self) -> DataPartReferenceRelationships {
        unimplemented!()
    }
    fn external_relationships(&self) -> ExternalRelationships{
        unimplemented!()
    }
    fn hyperlink_relationships(&self) -> HyperlinkRelationships{
        unimplemented!()
    }
    fn parts(&self) -> PartPairs{
        unimplemented!()
    }
}