mod reference;

pub use reference::ReferenceRelationship;
pub struct RelationshipId(String);
pub struct ExternalRelationship(ReferenceRelationship);
pub struct DataPartReferenceRelationship(ReferenceRelationship);

pub struct HyperlinkRelationship(ReferenceRelationship);

pub struct DataPartReferenceRelationships {}
pub struct ExternalRelationships {}

pub struct HyperlinkRelationships {}